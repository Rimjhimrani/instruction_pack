import streamlit as st
import pandas as pd
import numpy as np
import os
import json
import hashlib
import nltk
from datetime import datetime
from difflib import SequenceMatcher
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.drawing.image import Image as OpenpyxlImage
import re
import io
import tempfile
import shutil
from pathlib import Path
from collections import defaultdict
import zipfile
from PIL import Image
import base64
import traceback

# Configure Streamlit page
st.set_page_config(
    page_title="AI Template Mapper - Enhanced with Images",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Try to import optional dependencies
try:
    from nltk.tokenize import word_tokenize
    from nltk.corpus import stopwords
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    
    def initialize_nltk():
        """Initialize NLTK with proper downloads and fallbacks"""
        try:
            required_downloads = [
                ('punkt', 'tokenizers/punkt'),
                ('punkt_tab', 'tokenizers/punkt_tab'), 
                ('stopwords', 'corpora/stopwords')
            ]
            
            for download_name, path in required_downloads:
                try:
                    nltk.data.find(path)
                except LookupError:
                    try:
                        nltk.download(download_name, quiet=True)
                    except Exception as e:
                        print(f"Warning: Could not download {download_name}: {e}")
            
            word_tokenize("test")
            return True
            
        except Exception as e:
            print(f"NLTK initialization failed: {e}")
            return False
    
    NLTK_READY = initialize_nltk()
    
    if NLTK_READY:
        ADVANCED_NLP = True
    else:
        ADVANCED_NLP = False
        st.warning("⚠️ NLTK initialization failed. Using basic text processing.")
        
except ImportError as e:
    ADVANCED_NLP = False
    NLTK_READY = False
    st.warning("⚠️ Advanced NLP features disabled. Install nltk and scikit-learn for better matching.")

class ImageExtractor:
    """Handles image extraction from Excel files with improved duplicate handling and debugging"""
    
    def __init__(self):
        self.supported_formats = ['.png', '.jpg', '.jpeg', '.gif', '.bmp']
        self._placement_counters = defaultdict(int)
        self.current_excel_path = None
    
    def identify_image_upload_areas(self, worksheet):
        """Identify areas in template designated for image uploads with better categorization"""
        upload_areas = []
        processed_cells = set()
        
        try:
            # More specific and non-overlapping image keywords
            image_keywords = {
                'primary': ['primary packaging', 'primary pack', 'primary'],
                'secondary': ['secondary packaging', 'secondary pack', 'secondary'],
                'current': ['current packaging', 'current', 'existing packaging', 'existing'],
                'label': ['label', 'product label', 'labeling', 'labels']
            }
            
            print("=== Scanning for image upload areas ===")
            
            # Search through first few rows for column headers
            for row_num in range(1, min(15, worksheet.max_row + 1)):
                for col_num in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=row_num, column=col_num)
                    cell_coord = f"{row_num}_{col_num}"
                    
                    if cell_coord in processed_cells or not cell.value:
                        continue
                        
                    cell_text = str(cell.value).lower().strip()
                    
                    # Find the best matching category with stricter matching
                    best_match = None
                    best_score = 0
                    
                    for category, keywords in image_keywords.items():
                        for keyword in keywords:
                            # More precise matching to avoid cross-category matches
                            if keyword in cell_text:
                                score = len(keyword)
                                
                                # High bonus for exact match
                                if keyword == cell_text:
                                    score += 20
                                # Bonus for word boundary match
                                elif cell_text.startswith(keyword) or cell_text.endswith(keyword):
                                    score += 10
                                # Penalty for partial matches to avoid confusion
                                elif len(cell_text) > len(keyword) * 2:
                                    score -= 5
                                
                                if score > best_score:
                                    best_match = (category, keyword)
                                    best_score = score
                    
                    if best_match and best_score > 0:
                        area_info = {
                            'position': cell.coordinate,
                            'row': row_num + 1,  # Position image below header
                            'column': col_num,
                            'text': cell.value,
                            'type': best_match[0],
                            'header_text': cell_text,
                            'matched_keyword': best_match[1],
                            'match_score': best_score
                        }
                        upload_areas.append(area_info)
                        processed_cells.add(cell_coord)
                        print(f"Found {best_match[0]} area at {cell.coordinate} (col {col_num}): '{cell.value}' (score: {best_score})")

            # Sort by type priority and then by column to ensure proper order
            type_priority = {'primary': 1, 'secondary': 2, 'current': 3, 'label': 4}
            upload_areas.sort(key=lambda x: (type_priority.get(x['type'], 5), x['column']))
            
            print(f"Total areas found: {len(upload_areas)}")
            return upload_areas
            
        except Exception as e:
            st.error(f"Error identifying image upload areas: {e}")
            return []

    def extract_images_from_excel(self, excel_file_path):
        """Extract unique images from Excel file with enhanced debugging and multiple extraction methods"""
        try:
            self.current_excel_path = excel_file_path
            images = {}
            image_hashes = set()
            
            print("=== ENHANCED IMAGE EXTRACTION ===")
            print(f"📁 File path: {excel_file_path}")
            
            # First, let's check if the file exists and is readable
            if not os.path.exists(excel_file_path):
                print("❌ File does not exist!")
                return {}
            
            file_size = os.path.getsize(excel_file_path)
            print(f"📊 File size: {file_size} bytes")
            
            # Try multiple methods to extract images
            methods_tried = []
            
            # METHOD 1: Standard openpyxl extraction
            try:
                print("\n🔍 METHOD 1: Standard openpyxl extraction")
                result1 = self._extract_with_openpyxl(excel_file_path)
                methods_tried.append(("openpyxl", len(result1)))
                images.update(result1)
                print(f"✅ Method 1 found {len(result1)} images")
            except Exception as e:
                print(f"❌ Method 1 failed: {e}")
                methods_tried.append(("openpyxl", 0))
            
            # METHOD 2: ZIP-based extraction (Excel files are ZIP archives)
            try:
                print("\n🔍 METHOD 2: ZIP-based extraction")
                result2 = self._extract_with_zipfile(excel_file_path)
                methods_tried.append(("zipfile", len(result2)))
                # Only add if we haven't found images yet
                if not images:
                    images.update(result2)
                print(f"✅ Method 2 found {len(result2)} images")
            except Exception as e:
                print(f"❌ Method 2 failed: {e}")
                methods_tried.append(("zipfile", 0))
            
            # METHOD 3: Using python-docx2txt for embedded objects
            try:
                print("\n🔍 METHOD 3: Alternative extraction using xlwings (if available)")
                result3 = self._extract_alternative_method(excel_file_path)
                methods_tried.append(("alternative", len(result3)))
                if not images:
                    images.update(result3)
                print(f"✅ Method 3 found {len(result3)} images")
            except Exception as e:
                print(f"❌ Method 3 failed: {e}")
                methods_tried.append(("alternative", 0))
            
            # Print summary
            print("\n=== EXTRACTION SUMMARY ===")
            for method, count in methods_tried:
                print(f"📊 {method}: {count} images")
            
            print(f"🎯 TOTAL UNIQUE IMAGES EXTRACTED: {len(images)}")
            
            if not images:
                print("\n⚠️ NO IMAGES FOUND - POSSIBLE REASONS:")
                print("1. Excel file contains no embedded images")
                print("2. Images are stored as external links rather than embedded")
                print("3. Images are in unsupported format")
                print("4. Images are stored in drawings/charts rather than as direct images")
                print("5. File might be corrupted or password protected")
                
                # Additional diagnostics
                self._run_diagnostics(excel_file_path)
            
            return {'all_sheets': images}
            
        except Exception as e:
            print(f"❌ CRITICAL ERROR in extract_images_from_excel: {e}")
            import traceback
            traceback.print_exc()
            return {}

    def _extract_with_openpyxl(self, excel_file_path):
        """Standard openpyxl image extraction"""
        images = {}
        
        try:
            workbook = openpyxl.load_workbook(excel_file_path, data_only=False)
            print(f"📋 Workbook loaded. Sheets: {workbook.sheetnames}")
            
            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
                print(f"🔍 Processing sheet: {sheet_name}")
                
                # Check for images in worksheet
                if hasattr(worksheet, '_images'):
                    print(f"📸 _images attribute exists: {len(worksheet._images) if worksheet._images else 0} images")
                    
                    if worksheet._images:
                        for idx, img in enumerate(worksheet._images):
                            try:
                                # Get image data
                                image_data = img._data()
                                
                                # Create hash to avoid duplicates
                                image_hash = hashlib.md5(image_data).hexdigest()
                                
                                # Create PIL Image
                                pil_image = Image.open(io.BytesIO(image_data))
                                
                                # Get position info
                                anchor = img.anchor
                                if hasattr(anchor, '_from') and anchor._from:
                                    col = anchor._from.col
                                    row = anchor._from.row
                                    position = f"{get_column_letter(col + 1)}{row + 1}"
                                else:
                                    position = f"Image_{idx + 1}"
                                
                                # Convert to base64
                                buffered = io.BytesIO()
                                pil_image.save(buffered, format="PNG")
                                img_str = base64.b64encode(buffered.getvalue()).decode()
                                
                                # Classify image type
                                image_type = self._classify_image_type(sheet_name, position, idx)
                                
                                image_key = f"{image_type}_{sheet_name}_{position}_{idx}"
                                images[image_key] = {
                                    'data': img_str,
                                    'format': 'PNG',
                                    'size': pil_image.size,
                                    'position': position,
                                    'sheet': sheet_name,
                                    'index': idx,
                                    'type': image_type,
                                    'hash': image_hash
                                }
                                
                                print(f"✅ Extracted: {image_key} at {position}")
                                
                            except Exception as e:
                                print(f"❌ Failed to extract image {idx} from sheet {sheet_name}: {e}")
                else:
                    print(f"⚠️ No _images attribute found in sheet {sheet_name}")
                
                # Also check for drawing parts (charts, shapes, etc.)
                if hasattr(worksheet, '_charts') and worksheet._charts:
                    print(f"📊 Found {len(worksheet._charts)} charts in {sheet_name}")
                
                if hasattr(worksheet, '_drawing') and worksheet._drawing:
                    print(f"🎨 Found drawing elements in {sheet_name}")
            
            workbook.close()
            
        except Exception as e:
            print(f"❌ Error in openpyxl extraction: {e}")
            raise
        
        return images

    def _extract_with_zipfile(self, excel_file_path):
        """Extract images by treating Excel file as ZIP archive"""
        images = {}
        
        try:
            import zipfile
            
            with zipfile.ZipFile(excel_file_path, 'r') as zip_ref:
                # List all files in the archive
                file_list = zip_ref.namelist()
                print(f"📁 ZIP contents: {len(file_list)} files")
                
                # Look for media files
                media_files = [f for f in file_list if '/media/' in f.lower()]
                image_files = [f for f in file_list if any(f.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp'])]
                
                print(f"📸 Media files found: {len(media_files)}")
                print(f"🖼️ Image files found: {len(image_files)}")
                
                # Extract images from media folder
                for media_file in media_files:
                    try:
                        with zip_ref.open(media_file) as img_file:
                            image_data = img_file.read()
                            
                            # Create PIL Image
                            pil_image = Image.open(io.BytesIO(image_data))
                            
                            # Convert to base64
                            buffered = io.BytesIO()
                            pil_image.save(buffered, format="PNG")
                            img_str = base64.b64encode(buffered.getvalue()).decode()
                            
                            # Create hash
                            image_hash = hashlib.md5(image_data).hexdigest()
                            
                            # Generate key
                            filename = os.path.basename(media_file)
                            image_key = f"zip_{filename}_{len(images)}"
                            
                            images[image_key] = {
                                'data': img_str,
                                'format': 'PNG',
                                'size': pil_image.size,
                                'position': f"ZIP_{len(images)}",
                                'sheet': 'ZIP_EXTRACTED',
                                'index': len(images),
                                'type': 'current',  # Default type
                                'hash': image_hash,
                                'source_path': media_file
                            }
                            
                            print(f"✅ ZIP extracted: {image_key}")
                            
                    except Exception as e:
                        print(f"❌ Failed to extract {media_file}: {e}")
                
                # Also check for direct image files
                for img_file in image_files:
                    if img_file not in media_files:  # Avoid duplicates
                        try:
                            with zip_ref.open(img_file) as f:
                                image_data = f.read()
                                
                                pil_image = Image.open(io.BytesIO(image_data))
                                
                                buffered = io.BytesIO()
                                pil_image.save(buffered, format="PNG")
                                img_str = base64.b64encode(buffered.getvalue()).decode()
                                
                                image_hash = hashlib.md5(image_data).hexdigest()
                                
                                filename = os.path.basename(img_file)
                                image_key = f"direct_{filename}_{len(images)}"
                                
                                images[image_key] = {
                                    'data': img_str,
                                    'format': 'PNG',
                                    'size': pil_image.size,
                                    'position': f"DIRECT_{len(images)}",
                                    'sheet': 'DIRECT_EXTRACTED',
                                    'index': len(images),
                                    'type': 'primary',  # Default type
                                    'hash': image_hash,
                                    'source_path': img_file
                                }
                                
                                print(f"✅ Direct extracted: {image_key}")
                                
                        except Exception as e:
                            print(f"❌ Failed to extract direct image {img_file}: {e}")
        
        except Exception as e:
            print(f"❌ Error in ZIP extraction: {e}")
            raise
        
        return images

    def _extract_alternative_method(self, excel_file_path):
        """Alternative extraction method using other libraries if available"""
        images = {}
        
        try:
            # Try using xlrd for older Excel files
            print("🔍 Attempting xlrd-based extraction...")
            # This is a placeholder - xlrd doesn't directly support image extraction
            # but we can try to detect if it's an older format
            
        except Exception as e:
            print(f"❌ Alternative method failed: {e}")
        
        return images

    def _run_diagnostics(self, excel_file_path):
        """Run diagnostic checks on the Excel file"""
        try:
            print("\n🔍 RUNNING DIAGNOSTICS...")
            
            # Check file extension
            _, ext = os.path.splitext(excel_file_path)
            print(f"📄 File extension: {ext}")
            
            # Try to open with different methods
            try:
                import openpyxl
                wb = openpyxl.load_workbook(excel_file_path)
                print(f"✅ Openpyxl can open file")
                print(f"📋 Sheets: {wb.sheetnames}")
                
                # Check each sheet for any objects
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    print(f"\n🔍 Sheet '{sheet_name}' diagnostics:")
                    print(f"   - Max row: {ws.max_row}")
                    print(f"   - Max column: {ws.max_column}")
                    print(f"   - Has _images: {hasattr(ws, '_images')}")
                    print(f"   - Has _charts: {hasattr(ws, '_charts')}")
                    print(f"   - Has _drawing: {hasattr(ws, '_drawing')}")
                    
                    if hasattr(ws, '_images') and ws._images:
                        print(f"   - Images count: {len(ws._images)}")
                    
                    if hasattr(ws, '_charts') and ws._charts:
                        print(f"   - Charts count: {len(ws._charts)}")
                
                wb.close()
                
            except Exception as e:
                print(f"❌ Openpyxl cannot open file: {e}")
            
            # Check ZIP structure
            try:
                import zipfile
                with zipfile.ZipFile(excel_file_path, 'r') as zip_ref:
                    files = zip_ref.namelist()
                    print(f"\n📁 ZIP structure analysis:")
                    print(f"   - Total files: {len(files)}")
                    
                    media_files = [f for f in files if 'media' in f.lower()]
                    drawing_files = [f for f in files if 'drawing' in f.lower()]
                    chart_files = [f for f in files if 'chart' in f.lower()]
                    
                    print(f"   - Media files: {len(media_files)}")
                    print(f"   - Drawing files: {len(drawing_files)}")
                    print(f"   - Chart files: {len(chart_files)}")
                    
                    if media_files:
                        print(f"   - Media files found: {media_files}")
                    
            except Exception as e:
                print(f"❌ Cannot analyze as ZIP: {e}")
                
        except Exception as e:
            print(f"❌ Diagnostics failed: {e}")

    def _classify_image_type(self, sheet_name, position, index):
        """Fixed version that doesn't reload the workbook"""
        print(f"Classifying image {index} from sheet '{sheet_name}' at position '{position}'")
        try:
            # Simple fallback classification based on index to avoid workbook reloading issues
            fallback_types = ['current', 'primary', 'secondary', 'label']
            fallback_type = fallback_types[index % len(fallback_types)]
        
            # Try to get column info if position is valid
            if position and re.match(r'^[A-Z]+\d+$', position):
                col_letter = re.sub(r'\d+', '', position)
            
                # Simple heuristic based on column position
                try:
                    col_index = column_index_from_string(col_letter)
                    if col_index <= 5:  # Early columns = current
                        return 'current'
                    elif col_index <= 10:  # Middle columns = primary
                        return 'primary'
                    elif col_index <= 15:  # Later columns = secondary
                        return 'secondary'
                    else:  # Far right = label
                        return 'label'
                except:
                    pass
        
            print(f"-> Using fallback classification: {fallback_type}")
            return fallback_type

        except Exception as e:
            print(f"❌ Error in _classify_image_type: {e}")
            return 'unknown'

    def add_images_to_template(self, worksheet, uploaded_images, image_areas):
        """Add uploaded images to template - ENHANCED WITH BETTER ERROR HANDLING"""
        try:
            added_images = 0
            temp_image_paths = []
    
            print("=== Adding images to template ===")
            print(f"Available images: {len(uploaded_images)}")
    
            # Debug: Print all available images
            for img_key, img_data in uploaded_images.items():
                print(f"Available: {img_key} -> type: {img_data.get('type', 'unknown')}")
    
            # Process EACH image type separately and ensure they all get added
            row_42_column_position = 1  # Start at column A for row 42

            # 1. CURRENT PACKAGING - Always goes to T3
            current_images = [
                (k, v) for k, v in uploaded_images.items() 
                if v.get('type', '').lower() == 'current'
            ]
            print(f"\n--- CURRENT PACKAGING ({len(current_images)} images) ---")
            for img_key, img_data in current_images:
                success = self._place_image_at_position(
                    worksheet, img_key, img_data, 'T3', 
                    width_cm=8.3, height_cm=8.3, temp_image_paths=temp_image_paths
                )
                if success:
                    added_images += 1
                    print(f"✅ CURRENT placed at T3: {img_key}")
                else:
                    print(f"❌ CURRENT failed: {img_key}")

            # 2. PRIMARY PACKAGING - Goes to row 42, column A
            primary_images = [
                (k, v) for k, v in uploaded_images.items() 
                if v.get('type', '').lower() == 'primary'
            ]
            print(f"\n--- PRIMARY PACKAGING ({len(primary_images)} images) ---")
            for img_key, img_data in primary_images:
                cell_pos = f"{get_column_letter(row_42_column_position)}42"
                success = self._place_image_at_position(
                    worksheet, img_key, img_data, cell_pos,
                    width_cm=4.3, height_cm=4.3, temp_image_paths=temp_image_paths
                )
                if success:
                    added_images += 1
                    print(f"✅ PRIMARY placed at {cell_pos}: {img_key}")
                    # Move to next position for row 42
                    image_width_cols = int(4.3 * 1.162)  # ≈ 5 columns
                    gap_cols = int(1.162 * 1.162)         # ≈ 3 columns gap  
                    row_42_column_position += image_width_cols + gap_cols
                else:
                    print(f"❌ PRIMARY failed: {img_key}")

            # 3. SECONDARY PACKAGING - Goes to row 42, next position
            secondary_images = [
                (k, v) for k, v in uploaded_images.items() 
                if v.get('type', '').lower() == 'secondary'
            ]
            print(f"\n--- SECONDARY PACKAGING ({len(secondary_images)} images) ---")
            for img_key, img_data in secondary_images:
                cell_pos = f"{get_column_letter(row_42_column_position)}42"
                success = self._place_image_at_position(
                    worksheet, img_key, img_data, cell_pos,
                    width_cm=4.3, height_cm=4.3, temp_image_paths=temp_image_paths
                )
                if success:
                    added_images += 1
                    print(f"✅ SECONDARY placed at {cell_pos}: {img_key}")
                    # Move to next position for row 42
                    image_width_cols = int(4.3 * 1.162)  # ≈ 5 columns
                    gap_cols = int(1.162 * 1.162)         # ≈ 3 columns gap
                    row_42_column_position += image_width_cols + gap_cols
                else:
                    print(f"❌ SECONDARY failed: {img_key}")

            # 4. LABEL - Goes to row 42, next position
            label_images = [
                (k, v) for k, v in uploaded_images.items() 
                if v.get('type', '').lower() == 'label'
            ]
            print(f"\n--- LABEL ({len(label_images)} images) ---")
            for img_key, img_data in label_images:
                cell_pos = f"{get_column_letter(row_42_column_position)}42"
                success = self._place_image_at_position(
                    worksheet, img_key, img_data, cell_pos,
                    width_cm=4.3, height_cm=4.3, temp_image_paths=temp_image_paths
                )
                if success:
                    added_images += 1
                    print(f"✅ LABEL placed at {cell_pos}: {img_key}")
                else:
                    print(f"❌ LABEL failed: {img_key}")

            print(f"\n✅ TOTAL IMAGES ADDED: {added_images}")
            print(f"📁 Temporary files created: {len(temp_image_paths)}")
    
            return added_images, temp_image_paths

        except Exception as e:
            st.error(f"Error adding images to template: {e}")
            print(f"CRITICAL ERROR in add_images_to_template: {e}")
            traceback.print_exc()
            return 0, []

    def _place_image_at_position(self, worksheet, img_key, img_data, cell_position, width_cm, height_cm, temp_image_paths):
        """Place a single image at the specified cell position - FIXED VERSION"""
        try:
            print(f"  Placing {img_key} at {cell_position} ({width_cm}x{height_cm}cm)")
        
            # Create temporary image file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                image_bytes = base64.b64decode(img_data['data'])
                tmp_img.write(image_bytes)
                tmp_img.flush()  # Ensure data is written
                tmp_img_path = tmp_img.name
        
            print(f"    Created temp file: {tmp_img_path}")
        
            # Verify temp file exists and has content
            if not os.path.exists(tmp_img_path):
                print(f"    ❌ Temp file doesn't exist: {tmp_img_path}")
                return False
            
            file_size = os.path.getsize(tmp_img_path)
            if file_size == 0:
                print(f"    ❌ Temp file is empty: {tmp_img_path}")
                return False
            
            print(f"    Temp file size: {file_size} bytes")
        
            # Create openpyxl image object
            try:
                img = OpenpyxlImage(tmp_img_path)
                print(f"    ✅ Created OpenpyxlImage object")
            except Exception as img_err:
                print(f"    ❌ Failed to create OpenpyxlImage: {img_err}")
                return False
        
            # Set image size (converting cm to pixels: 1cm ≈ 37.8 pixels)
            img.width = int(width_cm * 37.8)
            img.height = int(height_cm * 37.8)
            print(f"    Set size: {img.width}x{img.height} pixels")
        
            # Set position using simple anchor
            img.anchor = cell_position
            print(f"    Set anchor: {cell_position}")
        
            # Add image to worksheet
            try:
                worksheet.add_image(img)
                print(f"    ✅ Added image to worksheet")
            except Exception as add_err:
                print(f"    ❌ Failed to add image to worksheet: {add_err}")
                return False
        
            # Track temporary file for cleanup (but don't delete yet!)
            temp_image_paths.append(tmp_img_path)
        
            print(f"    ✅ Successfully placed {img_key} at {cell_position}")
            return True
        
        except Exception as e:
            print(f"    ❌ Failed to place {img_key} at {cell_position}: {e}")
            import traceback
            traceback.print_exc()
            return False

    def reclassify_extracted_images(self, extracted_images, classification_rules=None):
        """Reclassify extracted images based on new rules or manual assignment"""
        if not extracted_images or 'all_sheets' not in extracted_images:
            return extracted_images
            
        print("=== Reclassifying extracted images ===")
        
        # Default classification: first image = current, then cycle through others
        default_rules = {
            0: 'current',    # First image = current packaging
            1: 'primary',    # Second image = primary
            2: 'secondary',  # Third image = secondary  
            3: 'label'       # Fourth image = label
        }
        
        rules = classification_rules or default_rules
        
        # Sort images by their original index to maintain order
        sorted_images = []
        for img_key, img_data in extracted_images['all_sheets'].items():
            # Extract original index from the key or image data
            original_index = img_data.get('index', 0)
            sorted_images.append((img_key, img_data, original_index))
        
        # Sort by original index
        sorted_images.sort(key=lambda x: x[2])
        
        # Reclassified images
        reclassified_images = {}
        for position, (img_key, img_data, original_index) in enumerate(sorted_images):
            # Determine new type based on position in sorted list
            if position in rules:
                new_type = rules[position]
            else:
                # For additional images, cycle through non-current types
                type_cycle = ['primary', 'secondary', 'label']
                new_type = type_cycle[(position - 1) % len(type_cycle)]
            
            # Update image data with new type
            img_data['type'] = new_type
            
            # Create new key with correct type
            parts = img_key.split('_')
            if len(parts) >= 4:
                new_key = f"{new_type}_{parts[1]}_{parts[2]}_{parts[3]}"
            else:
                new_key = f"{new_type}_{img_key.split('_', 1)[1] if '_' in img_key else img_key}"
            
            reclassified_images[new_key] = img_data
            print(f"Reclassified: {img_key} -> {new_key} (type: {new_type})")
        
        return {'all_sheets': reclassified_images}
    
class EnhancedTemplateMapperWithImages:
    def __init__(self):
        self.similarity_threshold = 0.3
        self.image_extractor = ImageExtractor()
        self.stop_words = {
            'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from',
            'has', 'he', 'in', 'is', 'it', 'its', 'of', 'on', 'that', 'the',
            'to', 'was', 'will', 'with', 'or', 'but', 'not', 'this', 'have',
            'had', 'what', 'when', 'where', 'who', 'which', 'why', 'how'
        }
        self.packaging_procedures = {
            "BOX IN BOX SENSITIVE": [
                "Pick up 1 quantity of part and apply bubble wrapping over it",
                "Apply tape and Put 1 such bubble wrapped part into a carton box [L-{Inner L} mm, W-{Inner W} mm, H-{Inner H} mm]",
                "Seal carton box and put {Inner Qty/Pack} such carton boxes into another carton box [L-{Inner L} mm, W-{Inner W} mm, H-{Inner H} mm]",
                "Seal carton box and put Traceability label as per PMSPL standard guideline",
                "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Inner Qty/Pack})",
                "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet.",
                "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level (max height including pallet -1000 mm)",
                "Put corner / edge protector and apply pet strap (2 times -- cross way)",
                "Apply traceability label on complete pack",
                "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only."
            ],
            
            "BOX IN BOX": [
                "Pick up 1 quantity of part and put it in a polybag",
                "seal the polybag and put it into a carton box [L-{Inner L} mm, W-{Inner W} mm, H-{Inner H} mm]",
                "Put {Inner Qty/Pack} such carton boxes into another carton box [L-{Inner L} mm, W-{Inner W} mm, H-{Inner H} mm]",
                "Seal carton box and put Traceability label as per PMSPL standard guideline",
                "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Inner Qty/Pack})",
                "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet.",
                "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level (max height including pallet -1000 mm)",
                "Put corner / edge protector and apply pet strap (2 times -- cross way)",
                "Apply traceability label on complete pack",
                "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only."
            ],
            
            "CARTON BOX WITH SEPARATOR FOR ONE PART": [
                "Pick up {Qty/Veh} parts and apply bubble wrapping over it (individually)",
                "Apply tape and Put bubble wrapped part into a carton box. Apply part separator & filler material between two parts to arrest part movement during handling",
                "Seal carton box and put Traceability label as per PMSPL standard guideline",
                "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Inner Qty/Pack})",
                "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level",
                "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet.",
                "Put corner / edge protector and apply pet strap (2 times -- cross way)",
                "Apply traceability label on complete pack",
                "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only.",
                ""
            ],
            
            "INDIVIDUAL NOT SENSITIVE": [
                "Pick up one part and put it into a polybag",
                "Seal polybag and Put polybag into a carton box",
                "Seal carton box and put Traceability label as per PMSPL standard guideline",
                "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Inner Qty/Pack})",
                "Load carton boxes on base wooden pallet -- Maximum {Layer} boxes per layer & Maximum {Level} level (max height including pallet - 1000 mm)",
                "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet.",
                "Put corner / edge protector and apply pet strap (2 times -- cross way)",
                "Apply traceability label on complete pack",
                "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only.",
                ""
            ],

            "INDIVIDUAL PROTECTION FOR EACH PART MANY TYPE": [
                "Pick up {Qty/Veh} parts and apply bubble wrapping over it (individually)",
                "Apply tape and Put bubble wrapped part into a carton box. Apply part separator &  filler material between two parts to arrest part movement during handling",															
		"Seal carton box and put Traceability label as per PMSPL standard guideline",														
                "Prepare additional carton boxes in line with procurement schedule ( multiple of  primary pack quantity – {Qty/Pack})",														
		"Load carton boxes on base wooden pallet – {Layer} boxes per layer & max {Level} level",														
                "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",															
		"Put corner / edge protector and apply pet strap ( 2 times – cross way)",															
                "Apply traceability label on complete pack",														
                "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",															
		"Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only",
            ],

            "INDIVIDUAL PROTECTION FOR EACH PART": [
                "Pick up {Qty/Veh} parts and apply bubble wrapping over it (individually)",
                "Apply tape and Put bubble wrapped part into a carton box. Apply part separator & filler material between two parts to arrest part movement during handling",
                "Seal carton box and put Traceability label as per PMSPL standard guideline",
                "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Inner Qty/Pack})",
                "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level (max height including pallet - 1000 mm)",
                "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet.",
                "Put corner / edge protector and apply pet strap (2 times -- cross way)",
                "Apply traceability label on complete pack",
                "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only.",
                ""
            ],
            
            "INDIVIDUAL SENSITIVE": [
                "Pick up one part and apply bubble wrapping over it",
                "Apply tape and Put bubble wrapped part into a carton box",
                "Seal carton box and put Traceability label as per PMSPL standard guideline",
                "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Inner Qty/Pack})",
                "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level (max height including pallet - 1000 mm)",
                "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet.",
                "Put corner / edge protector and apply pet strap (2 times -- cross way)",
                "Apply traceability label on complete pack",
                "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only.",
                ""
            ],
            
            "MANY IN ONE TYPE": [
                "Pick up {Qty/Veh} quantity of part and put it in a polybag",
                "Seal polybag and Put it into a carton box",
                "Seal carton box and put Traceability label as per PMSPL standard guideline",
                "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Inner Qty/Pack})",
                "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet.",
                "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level (max height including pallet - 1000 mm)",
                "Put corner / edge protector and apply pet strap (2 times -- cross way)",
                "Apply traceability label on complete pack",
                "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only.",
                ""
            ],
            
            "SINGLE BOX": [
                "Pick up 1 quantity of part and put it in a polybag",
                "Put into a carton box",
                "Seal carton box and put Traceability label as per PMSPL standard guideline",
                "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Inner Qty/Pack})",
                "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet.",
                "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level",
                "Put corner / edge protector and apply pet strap (2 times -- cross way) and stretch wrap it",
                "Apply traceability label on complete pack",
                "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only.",
                ""
            ]
        }
        
        # Enhanced section-based mapping rules
        self.section_mappings = {
            'primary_packaging': {
                'section_keywords': [
                    'primary packaging instruction', 'primary packaging', 'primary', 
                    'internal', '( primary / internal )', 'primary / internal'
                ],
                'field_mappings': {
                    'primary packaging type': 'Primary Packaging Type',
                    'packaging type': 'Primary Packaging Type',
                    'l-mm': 'Primary L-mm',
                    'l mm': 'Primary L-mm',
                    'length': 'Primary L-mm',
                    'w-mm': 'Primary W-mm',
                    'w mm': 'Primary W-mm', 
                    'width': 'Primary W-mm',
                    'h-mm': 'Primary H-mm',
                    'h mm': 'Primary H-mm',
                    'height': 'Primary H-mm',
                    'qty/pack': 'Primary Qty/Pack',
                    'quantity': 'Primary Qty/Pack',
                    'empty weight': 'Primary Empty Weight',
                    'pack weight': 'Primary Pack Weight'
                }
            },
            'secondary_packaging': {
                'section_keywords': [
                    'secondary packaging instruction', 'secondary packaging', 'secondary', 
                    'outer', 'external', '( outer / external )', 'outer / external'
                ],
                'field_mappings': {
                    'secondary packaging type': 'Secondary Packaging Type',
                    'packaging type': 'Secondary Packaging Type',
                    'type': 'Secondary Packaging Type',  # ← ADD THIS LINE
                    'l-mm': 'Secondary L-mm',
                    'l mm': 'Secondary L-mm',
                    'length': 'Secondary L-mm',
                    'w-mm': 'Secondary W-mm',
                    'w mm': 'Secondary W-mm',
                    'width': 'Secondary W-mm',
                    'h-mm': 'Secondary H-mm',
                    'h mm': 'Secondary H-mm',
                    'height': 'Secondary H-mm',
                    'qty/pack': 'Secondary Qty/Pack',
                    'quantity': 'Secondary Qty/Pack',
                    'empty weight': 'Secondary Empty Weight',
                    'pack weight': 'Secondary Pack Weight'
                }
            },
            'part_information': {
                'section_keywords': [
                    'part information', 'part info', 'part', 'component', 'item', 'component information'
                ],
                'field_mappings': {
                    # Enhanced part dimension mappings
                    'L': 'Part L',
                    'l': 'Part L',
                    'length': 'Part L',
                    'part l': 'Part L',
                    'component l': 'Part L',
                    'W': 'Part W',
                    'w': 'Part W',
                    'width': 'Part W',
                    'part w': 'Part W',
                    'component w': 'Part W',
                    'H': 'Part H',
                    'h': 'Part H',
                    'height': 'Part H',
                    'part h': 'Part H',
                    'component h': 'Part H',
                    # Other part fields
                    'part no': 'Part No',
                    'part number': 'Part No',
                    'description': 'Part Description',
                    'unit weight': 'Part Unit Weight'
                }
            },
            'vendor_information': {
                'section_keywords': [
                    'vendor information', 'vendor info', 'vendor', 'supplier', 'supplier information', 'supplier info'
                ],
                'field_mappings': {
                    # Enhanced vendor field mappings
                    'vendor name': 'Vendor Name',
                    'name': 'Vendor Name',
                    'supplier name': 'Vendor Name',
                    'vendor code': 'Vendor Code',
                    'supplier code': 'Vendor Code',
                    'code': 'Vendor Code',
                    'vendor location': 'Vendor Location',
                    'location': 'Vendor Location',
                    'supplier location': 'Vendor Location',
                    'address': 'Vendor Location'
                }
            }
        }
        
        if ADVANCED_NLP:
            try:
                self.stop_words = set(stopwords.words('english'))
                self.vectorizer = TfidfVectorizer(stop_words='english', ngram_range=(1, 2))
            except:
                pass
    
    def preprocess_text(self, text):
        """Preprocess text for better matching"""
        try:
            if pd.isna(text) or text is None:
                return ""
            
            text = str(text).lower()
            # Remove parentheses and special characters but keep spaces
            text = re.sub(r'[()[\]{}]', ' ', text)
            text = re.sub(r'[^\w\s/-]', ' ', text)
            text = re.sub(r'\s+', ' ', text).strip()
            
            return text
        except Exception as e:
            st.error(f"Error in preprocess_text: {e}")
            return ""
    
    def extract_keywords(self, text):
        """Extract keywords from text with improved error handling"""
        try:
            text = self.preprocess_text(text)
            if not text:
                return []
                
            if ADVANCED_NLP and NLTK_READY:
                try:
                    tokens = word_tokenize(text)
                    keywords = [token for token in tokens if token not in self.stop_words and len(token) > 1]
                    return keywords
                except Exception as e:
                    print(f"NLTK tokenization failed, using fallback: {e}")
            
            tokens = text.split()
            keywords = [token for token in tokens if token not in self.stop_words and len(token) > 1]
            return keywords
        except Exception as e:
            st.error(f"Error in extract_keywords: {e}")
            return []
    
    def identify_section_context(self, worksheet, row, col, max_search_rows=15):
        """Enhanced section identification with better pattern matching"""
        try:
            section_context = None
        
            # Search upwards and in nearby cells for section headers
            for search_row in range(max(1, row - max_search_rows), row + 2):  # Include current row + 1
                for search_col in range(max(1, col - 15), min(worksheet.max_column + 1, col + 15)):
                    try:
                        cell = worksheet.cell(row=search_row, column=search_col)
                        if cell.value:
                            cell_text = self.preprocess_text(str(cell.value))
                        
                            # Check for section keywords with more flexible matching
                            for section_name, section_info in self.section_mappings.items():
                                for keyword in section_info['section_keywords']:
                                    keyword_processed = self.preprocess_text(keyword)
                                
                                    # Exact match
                                    if keyword_processed == cell_text:
                                        return section_name
                                
                                    # Partial match for key phrases
                                    if keyword_processed in cell_text or cell_text in keyword_processed:
                                        return section_name
                                
                                    # Enhanced context matching
                                    if section_name == 'primary_packaging':
                                        if ('primary' in cell_text and ('packaging' in cell_text or 'internal' in cell_text)):
                                            return section_name
                                    elif section_name == 'secondary_packaging':
                                        if ('secondary' in cell_text and ('packaging' in cell_text or 'outer' in cell_text or 'external' in cell_text)):
                                            return section_name
                                    elif section_name == 'part_information':
                                        if (('part' in cell_text and ('information' in cell_text or 'info' in cell_text)) or
                                            ('component' in cell_text and ('information' in cell_text or 'info' in cell_text))):
                                            return section_name
                                    elif section_name == 'vendor_information':
                                        if (('vendor' in cell_text and ('information' in cell_text or 'info' in cell_text)) or
                                            ('supplier' in cell_text and ('information' in cell_text or 'info' in cell_text))):
                                            return section_name
                    except:
                        continue
        
            return section_context
        
        except Exception as e:
            st.error(f"Error in identify_section_context: {e}")
            return None
    
    def calculate_similarity(self, text1, text2):
        """Calculate similarity between two texts"""
        try:
            if not text1 or not text2:
                return 0.0
            
            text1 = self.preprocess_text(text1)
            text2 = self.preprocess_text(text2)
            
            if not text1 or not text2:
                return 0.0
            
            # Sequence similarity
            sequence_sim = SequenceMatcher(None, text1, text2).ratio()
            
            # TF-IDF similarity (if available)
            tfidf_sim = 0.0
            if ADVANCED_NLP:
                try:
                    tfidf_matrix = self.vectorizer.fit_transform([text1, text2])
                    tfidf_sim = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0]
                except:
                    tfidf_sim = 0.0
            
            # Keyword overlap
            keywords1 = set(self.extract_keywords(text1))
            keywords2 = set(self.extract_keywords(text2))
            
            if keywords1 and keywords2:
                keyword_sim = len(keywords1.intersection(keywords2)) / len(keywords1.union(keywords2))
            else:
                keyword_sim = 0.0
            
            # Weighted average
            if ADVANCED_NLP:
                final_similarity = (sequence_sim * 0.4) + (tfidf_sim * 0.4) + (keyword_sim * 0.2)
            else:
                final_similarity = (sequence_sim * 0.7) + (keyword_sim * 0.3)
            
            return final_similarity
        except Exception as e:
            st.error(f"Error in calculate_similarity: {e}")
            return 0.0
    
    def is_mappable_field(self, text):
        """Enhanced field detection for packaging templates"""
        try:
            if not text or pd.isna(text):
                return False
            
            text = str(text).lower().strip()
            if not text:
                return False
        
            # DEBUG: Print what we're checking
            print(f"DEBUG is_mappable_field: Checking '{text}'")
        
            # Skip header-like patterns that should not be treated as fields
            header_exclusions = [
                'vendor information', 'part information', 'primary packaging', 'secondary packaging',
                'packaging instruction', 'procedure', 'steps', 'process'
            ]
        
            for exclusion in header_exclusions:
                if exclusion in text and 'type' not in text:
                    print(f"DEBUG: Excluding '{text}' as header")
                    return False
        
            # Define mappable field patterns for packaging templates
            mappable_patterns = [
                # Packaging type fields
                r'packaging\s+type', r'\btype\b',
            
                # Dimension fields
                r'\bl[-\s]*mm\b', r'\bw[-\s]*mm\b', r'\bh[-\s]*mm\b',
                r'\bl\b', r'\bw\b', r'\bh\b',  # Single letter dimensions
            
                # Part-specific dimension fields
                r'part\s+l\b', r'part\s+w\b', r'part\s+h\b',
            
                # Basic dimensions
                r'\blength\b', r'\bwidth\b', r'\bheight\b',
            
                # Other fields
                r'qty[/\s]*pack', r'quantity\b', r'weight\b', r'empty\s+weight',
                r'\bcode\b', r'\bname\b', r'\bdescription\b', r'\blocation\b',
                r'part\s+no\b', r'part\s+number\b'
            ]
        
            for pattern in mappable_patterns:
                if re.search(pattern, text):
                    print(f"DEBUG: '{text}' matches pattern '{pattern}'")
                    return True
        
            # Check if it ends with colon
            if text.endswith(':'):
                print(f"DEBUG: '{text}' ends with colon")
                return True
                
            print(f"DEBUG: '{text}' is NOT mappable")
            return False
        except Exception as e:
            st.error(f"Error in is_mappable_field: {e}")
            return False
    
    def find_procedure_step_area(self, worksheet):
        """Find area in template where procedure steps should be written"""
        try:
            procedure_keywords = [
                'procedure', 'steps', 'process', 'instruction', 'method',
                'packaging procedure', 'packing steps', 'process steps',
                'step 1', 'step 2', 'step 3', 'step 4', 'step 5',
                'step 6', 'step 7', 'step 8', 'step 9', 'step 10', 'step 11'
            ]
        
            # Search for procedure area indicators
            for row_num in range(1, min(50, worksheet.max_row + 1)):
                for col_num in range(1, min(20, worksheet.max_column + 1)):
                    cell = worksheet.cell(row=row_num, column=col_num)
                    if not cell.value:
                        continue
                
                    cell_text = str(cell.value).lower().strip()
                
                    # Check for procedure keywords
                    for keyword in procedure_keywords:
                        if keyword in cell_text:
                            print(f"Found procedure area indicator at {cell.coordinate}: '{cell.value}'")
                            # Return fixed position: Row 28, Column B (2)
                            return {
                                'start_row': 28,
                                'start_col': 2,  # Column B
                                'header_text': cell.value,
                                'header_position': cell.coordinate
                            }
        
            # If no specific procedure area found, use fixed default location
            print("No procedure area found, using fixed default location (Row 28, Column B)")
            return {
                'start_row': 28,  # Fixed row 28
                'start_col': 2,   # Column B (2)
                'header_text': 'Packaging Procedure Steps',
                'header_position': 'B27'  # Header one row above
            }
        except Exception as e:
            st.error(f"Error finding procedure step area: {e}")
            return None

    
    def write_procedure_steps_to_template(self, worksheet, packaging_type, data_dict=None):
        """Write packaging procedure steps in Column B starting from Row 28 (Step numbers already exist in Column A)"""
        try:
            from openpyxl.cell import MergedCell
            from openpyxl.styles import Font, Alignment

            print(f"\n=== WRITING PROCEDURE STEPS FOR {packaging_type} ===")

            # Get the procedure steps
            steps = self.get_procedure_steps(packaging_type, data_dict)
            if not steps:
                print(f"❌ No procedure steps found for packaging type: {packaging_type}")
                return 0

            print(f"📋 Retrieved {len(steps)} procedure steps")

            # Fixed column and starting row
            start_row = 28      # Start from Row 28
            target_col = 2      # Column B (step content)
    
            # Filter out empty or blank steps
            non_empty_steps = [step for step in steps if step and step.strip()]
            steps_to_write = non_empty_steps

            print(f"✏️  Will write {len(steps_to_write)} non-empty steps")

            steps_written = 0

            for i, step in enumerate(steps_to_write):
                step_row = start_row + i
                step_text = step.strip()
                target_cell = worksheet.cell(row=step_row, column=target_col)
                print(f"📝 Writing step {i + 1} to B{step_row}: {step_text[:50]}...")

                # 🔧 HARD FIX ONLY FOR ROW 37 (Step 10)
                if step_row == 37:
                    for merged_range in worksheet.merged_cells.ranges:
                        if "B37" in str(merged_range):
                            print(f"🔧 Forcing unmerge of B37 range: {merged_range}")
                            worksheet.unmerge_cells(str(merged_range))
                            break
                    target_cell = worksheet.cell(row=37, column=2)  # re-fetch after unmerge

                # Write step content
                target_cell.value = step_text
                target_cell.font = Font(name='Calibri', size=10)
                target_cell.alignment = Alignment(wrap_text=True, vertical='top')

                # 🔧 RE-MERGE ROW 37 AFTER WRITING CONTENT
                if step_row == 37:
                    try:
                        # Merge B37 with adjacent cells (typically B37:P37 for procedure steps)
                        merge_range = f"B37:P37"
                        worksheet.merge_cells(merge_range)
                        print(f"✅ Re-merged row 37: {merge_range}")
                    except Exception as merge_error:
                        print(f"⚠️ Warning: Could not re-merge B37: {merge_error}")

                # Adjust height manually based on estimated lines
                max_chars_per_line = 100
                num_lines = max(1, len(step_text) // max_chars_per_line + 1)
                estimated_height = 15 + (num_lines - 1) * 15
                worksheet.row_dimensions[step_row].height = estimated_height

                steps_written += 1

            print(f"\n✅ PROCEDURE STEPS COMPLETED")
            print(f"   Total steps written: {steps_written}")
            print(f"   Location: Column B, starting from Row 28")

            return steps_written

        except Exception as e:
            print(f"💥 Critical error in write_procedure_steps_to_template: {e}")
            traceback.print_exc()
            return 0
            
    def get_procedure_steps(self, packaging_type, data_dict=None):
        procedures = self.packaging_procedures.get(packaging_type, [""] * 11)
        if data_dict:
            filled_procedures = []
            for procedure in procedures:
                filled_procedure = procedure
                replacements = {
                    '{Inner L}': str(data_dict.get('Inner L', 'XXX')),
                    '{Inner W}': str(data_dict.get('Inner W', 'XXX')),
                    '{Inner H}': str(data_dict.get('Inner H', 'XXX')),
                    '{Inner Qty/Pack}': str(data_dict.get('Inner Qty/Pack', 'XXX')),
                    '{Qty/Pack}': str(data_dict.get('Inner Qty/Pack', data_dict.get('Qty/Pack', 'XXX'))),
                    '{Qty/Veh}': str(data_dict.get('Qty/Veh', 'XXX')),
                    '{Layer}': str(data_dict.get('Layer', 'XXX')),
                    '{Level}': str(data_dict.get('Level', 'XXX')),
                }
                for placeholder, value in replacements.items():
                    filled_procedure = filled_procedure.replace(placeholder, value)
                filled_procedures.append(filled_procedure)
            return filled_procedures
        else:
            return procedures
            
    def find_template_fields_with_context_and_images(self, template_file):
        """Find template fields and image upload areas"""
        fields = {}
        image_areas = []
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
        
            merged_ranges = worksheet.merged_cells.ranges
        
            # Find mappable fields
            for row in worksheet.iter_rows():
                for cell in row:
                    try:
                        if cell.value is not None:
                            cell_value = str(cell.value).strip()
                        
                            if cell_value and self.is_mappable_field(cell_value):
                                cell_coord = cell.coordinate
                                merged_range = None
                            
                                for merge_range in merged_ranges:
                                    if cell.coordinate in merge_range:
                                        merged_range = str(merge_range)
                                        break
                            
                                # Identify section context
                                section_context = self.identify_section_context(
                                    worksheet, cell.row, cell.column
                                )
                            
                                # DEBUG PRINTS - ADD THESE LINES
                                print(f"DEBUG: Found field '{cell_value}' at {cell_coord}")
                                print(f"DEBUG: Section context: {section_context}")
                                print(f"DEBUG: Is mappable: {self.is_mappable_field(cell_value)}")
                                print("---")
                            
                                fields[cell_coord] = {
                                    'value': cell_value,
                                    'row': cell.row,
                                    'column': cell.column,
                                    'merged_range': merged_range,
                                    'section_context': section_context,
                                    'is_mappable': True
                                }
                    except Exception as e:
                        continue
        
            # Find image upload areas
            image_areas = self.image_extractor.identify_image_upload_areas(worksheet)
        
            workbook.close()
        
        except Exception as e:
            st.error(f"Error reading template: {e}")
    
        return fields, image_areas
    
    def map_data_with_section_context(self, template_fields, data_df):
        """Enhanced mapping with better section-aware logic"""
        mapping_results = {}
        used_columns = set()

        try:
            data_columns = data_df.columns.tolist()
            print(f"DEBUG: Available data columns: {data_columns}")  # ADD THIS

            for coord, field in template_fields.items():
                try:
                    best_match = None
                    best_score = 0.0
                    field_value = field['value']
                    section_context = field.get('section_context')

                    print(f"DEBUG: Mapping field '{field_value}' with section '{section_context}'")  # ADD THIS

                    # If section context exists, use its field mappings
                    if section_context and section_context in self.section_mappings:
                        section_mappings = self.section_mappings[section_context]['field_mappings']
                        print(f"DEBUG: Section mappings: {section_mappings}")  # ADD THIS

                        for template_field_key, data_column_pattern in section_mappings.items():
                            normalized_field_value = self.preprocess_text(field_value)
                            normalized_template_key = self.preprocess_text(template_field_key)

                            print(f"DEBUG: Comparing '{normalized_field_value}' with '{normalized_template_key}'")  # ADD THIS

                            if normalized_field_value == normalized_template_key:
                                # Prefer section-prefixed column
                                section_prefix = section_context.split('_')[0].capitalize()
                                expected_column = f"{section_prefix} {data_column_pattern}".strip()
                            
                                print(f"DEBUG: Looking for expected column: '{expected_column}'")  # ADD THIS

                                for data_col in data_columns:
                                    if data_col in used_columns:
                                        continue
                                    if self.preprocess_text(data_col) == self.preprocess_text(expected_column):
                                        best_match = data_col
                                        best_score = 1.0
                                        print(f"DEBUG: EXACT MATCH FOUND: {data_col}")  # ADD THIS
                                        break

                                # Fallback to similarity match if no exact match
                                if not best_match:
                                    for data_col in data_columns:
                                        if data_col in used_columns:
                                            continue
                                        similarity = self.calculate_similarity(expected_column, data_col)
                                        if similarity > best_score and similarity >= self.similarity_threshold:
                                            best_score = similarity
                                            best_match = data_col
                                            print(f"DEBUG: SIMILARITY MATCH: {data_col} (score: {similarity})")  # ADD THIS
                                break
                    # 🔧 Fallback 1: If 'type' and no section, assume secondary packaging
                    if not section_context and self.preprocess_text(field_value) == 'type':
                        section_context = 'secondary_packaging'
                        section_mappings = self.section_mappings[section_context]['field_mappings']
                        print(f"⚠️ Fallback: Assuming 'secondary_packaging' for 'Type' at {coord}")

                        for template_field_key, data_column_pattern in section_mappings.items():
                            if self.preprocess_text(template_field_key) == 'type':
                                expected_column = data_column_pattern
                                for data_col in data_columns:
                                    if data_col in used_columns:
                                        continue
                                    if self.preprocess_text(data_col) == self.preprocess_text(expected_column):
                                        best_match = data_col
                                        best_score = 1.0
                                        break
                                break

                    # 🔧 Fallback 2: If 'L', 'W', 'H', etc. and no section, assume part_information
                    if not section_context and self.preprocess_text(field_value) in ['l', 'w', 'h', 'length', 'width', 'height']:
                        section_context = 'part_information'
                        section_mappings = self.section_mappings[section_context]['field_mappings']
                        print(f"⚠️ Fallback: Assuming 'part_information' for '{field_value}' at {coord}")

                        for template_field_key, data_column_pattern in section_mappings.items():
                            normalized_field_value = self.preprocess_text(field_value)
                            normalized_template_key = self.preprocess_text(template_field_key)

                            if normalized_field_value == normalized_template_key:
                                expected_column = data_column_pattern
                                for data_col in data_columns:
                                    if data_col in used_columns:
                                        continue
                                    if self.preprocess_text(data_col) == self.preprocess_text(expected_column):
                                        best_match = data_col
                                        best_score = 1.0
                                        break
                                break

                    # Final fallback if section mapping didn't resolve
                    if not best_match:
                        for data_col in data_columns:
                            if data_col in used_columns:
                                continue
                            similarity = self.calculate_similarity(field_value, data_col)
                            if similarity > best_score and similarity >= self.similarity_threshold:
                                best_score = similarity
                                best_match = data_col

                    print(f"DEBUG: Final mapping result - Field: '{field_value}' -> Column: '{best_match}' (Score: {best_score})")  # ADD THIS
                    print("=" * 50)  # ADD THIS

                    # Save mapping
                    mapping_results[coord] = {
                        'template_field': field_value,
                        'data_column': best_match,
                        'similarity': best_score,
                        'field_info': field,
                        'section_context': section_context,
                        'is_mappable': best_match is not None
                    }

                    # Prevent reuse of the same column
                    if best_match:
                        used_columns.add(best_match)

                except Exception as e:
                    st.error(f"Error mapping field {coord}: {e}")
                    continue

        except Exception as e:
            st.error(f"Error in map_data_with_section_context: {e}")

        return mapping_results

    
    def find_data_cell_for_label(self, worksheet, field_info):
        """Find data cell for a label with improved merged cell handling"""
        try:
            row = field_info['row']
            col = field_info['column']
            merged_ranges = list(worksheet.merged_cells.ranges)
        
            def is_suitable_data_cell(cell_coord):
                """Check if a cell is suitable for data entry"""
                try:
                    cell = worksheet[cell_coord]
                    if hasattr(cell, '__class__') and cell.__class__.__name__ == 'MergedCell':
                        return False
                    if cell.value is None or str(cell.value).strip() == "":
                        return True
                    # Check for data placeholder patterns
                    cell_text = str(cell.value).lower().strip()
                    data_patterns = [r'^_+$', r'^\.*$', r'^-+$', r'enter', r'fill', r'data']
                    return any(re.search(pattern, cell_text) for pattern in data_patterns)
                except:
                    return False
            
            # Strategy 1: Look right of label (most common pattern)
            for offset in range(1, 6):
                target_col = col + offset
                if target_col <= worksheet.max_column:
                    cell_coord = worksheet.cell(row=row, column=target_col).coordinate
                    if is_suitable_data_cell(cell_coord):
                        return cell_coord
            
            # Strategy 2: Look below label
            for offset in range(1, 4):
                target_row = row + offset
                if target_row <= worksheet.max_row:
                    cell_coord = worksheet.cell(row=target_row, column=col).coordinate
                    if is_suitable_data_cell(cell_coord):
                        return cell_coord
            
            # Strategy 3: Look in nearby area
            for r_offset in range(-1, 3):
                for c_offset in range(-1, 6):
                    if r_offset == 0 and c_offset == 0:
                        continue
                    target_row = row + r_offset
                    target_col = col + c_offset
                
                    if (target_row > 0 and target_row <= worksheet.max_row and 
                        target_col > 0 and target_col <= worksheet.max_column):
                            cell_coord = worksheet.cell(row=target_row, column=target_col).coordinate
                            if is_suitable_data_cell(cell_coord):
                                return cell_coord
            
            return None
            
        except Exception as e:
            st.error(f"Error in find_data_cell_for_label: {e}")
            return None
    
    def add_images_to_template(self, worksheet, uploaded_images, image_areas):
        """Add uploaded images to template in designated areas only if part number or description matches"""
        try:
            added_images = 0
            temp_image_paths = []
            used_images = set()

            # Get current part number and description
            part_no = str(data_dict.get('Part No', '')).lower()
            desc = str(data_dict.get('Part Description', '')).lower()

            for area in image_areas:
                area_type = area['type']
                label_text = area.get('text', '').lower()
                matching_image = None

                for label, img_data in uploaded_images.items():
                    if label in used_images:
                        continue

                    label_lower = label.lower()

                    # Match image if part number or description is in label
                    if (
                        part_no and part_no in label_lower
                    ) or (
                        desc and desc in label_lower
                    ) or (
                        area_type in label_lower
                    ) or (
                        area_type.replace('_', ' ') in label_lower
                    ) or (
                        label_lower in label_text or label_text in label_lower
                    ):
                        matching_image = img_data
                        used_images.add(label)
                        break

                # fallback to any unused image
                if not matching_image:
                    for label, img_data in uploaded_images.items():
                        if label not in used_images:
                            matching_image = img_data
                            used_images.add(label)
                            break

                if matching_image:
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                            image_bytes = base64.b64decode(matching_image['data'])
                            tmp_img.write(image_bytes)
                            tmp_img_path = tmp_img.name
                        img = OpenpyxlImage(tmp_img_path)
                        img.width = 250
                        img.height = 150

                        cell_coord = f"{get_column_letter(area['column'])}{area['row']}"
                        worksheet.add_image(img, cell_coord)

                        temp_image_paths.append(tmp_img_path)
                        added_images += 1
                    except Exception as e:
                        st.warning(f"Could not add image to {area['position']}: {e}")
                        continue

            return added_images, temp_image_paths

        except Exception as e:
            st.error(f"Error adding images to template: {e}")
            return 0, []
    
    def fill_template_with_data_and_images(self, template_file, mapping_results, data_df, uploaded_images=None, packaging_type=None):
        print("🛠️ Entered fill_template_with_data_and_images()")
        print(f"📂 Template file: {template_file}")
        print(f"📊 DataFrame shape: {data_df.shape}")
        print(f"🧩 Number of mappings: {len(mapping_results)}")
        print(f"🖼️ Uploaded images: {list(uploaded_images.keys()) if uploaded_images else 'None'}")
        print(f"📦 Packaging type: {packaging_type}")

        try:
            # ✅ Load template from UploadedFile or path
            if hasattr(template_file, "read"):
                print("📥 Reading template from UploadedFile")
                template_bytes = template_file.read()
                workbook = openpyxl.load_workbook(BytesIO(template_bytes))
            else:
                print("📁 Reading template from file path")
                workbook = openpyxl.load_workbook(template_file)

            filled_count = 0
            images_added = 0
            procedure_steps_added = 0
            temp_image_paths = []

            # ✅ Fill mapped fields
            for mapping in mapping_results:
                field_name = mapping.get('template_field')
                column = mapping.get('data_column')

                if column and field_name and column in data_df.columns:
                    try:
                        value = data_df.iloc[0][column]
                        print(f"✍️ Writing value '{value}' to field '{field_name}'")
                        for sheet in workbook.worksheets:
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value == field_name:
                                        cell.value = value
                                        filled_count += 1
                    except Exception as e:
                        print(f"⚠️ Failed to fill field '{field_name}': {e}")

            # ✅ Add procedure steps
            try:
                procedure_steps_added = self.add_procedure_steps_to_template(workbook, data_df, packaging_type)
                print(f"📜 Procedure steps added: {procedure_steps_added}")
            except Exception as pe:
                print(f"❌ Error adding procedure steps: {pe}")

            # ✅ Add images
            try:
                images_added = self.add_images_to_template(workbook, uploaded_images)
                print(f"🖼️ Images added: {images_added}")
            except Exception as ie:
                print(f"❌ Error adding images: {ie}")

            # ✅ Final debug before return
            print(f"✅ Workbook ready: {workbook is not None}")
            print(f"📊 Summary — Fields: {filled_count}, Images: {images_added}, Procedure Steps: {procedure_steps_added}")
            return workbook, filled_count, images_added, temp_image_paths, procedure_steps_added

        except Exception as e:
            print(f"❌ Critical error: {e}")
            return None, 0, 0, [], 0


# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'user_role' not in st.session_state:
    st.session_state.user_role = None
if 'templates' not in st.session_state:
    st.session_state.templates = {}
if 'enhanced_mapper' not in st.session_state:
    st.session_state.enhanced_mapper = EnhancedTemplateMapperWithImages()

# User management functions
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(password, hashed):
    return hash_password(password) == hashed

DEFAULT_USERS = {
    "admin": {
        "password": hash_password("admin123"),
        "role": "admin",
        "name": "Administrator"
    },
    "user1": {
        "password": hash_password("user123"),
        "role": "user",
        "name": "Regular User"
    }
}

def authenticate_user(username, password):
    if username in DEFAULT_USERS:
        if verify_password(password, DEFAULT_USERS[username]['password']):
            return DEFAULT_USERS[username]['role'], DEFAULT_USERS[username]['name']
    return None, None

def show_login():
    st.title("🤖 Enhanced AI Template Mapper with Images")
    st.markdown("### Advanced packaging template processing with image support")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        with st.form("login_form"):
            st.subheader("Login")
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Login", use_container_width=True)
            
            if submit:
                role, name = authenticate_user(username, password)
                if role:
                    st.session_state.authenticated = True
                    st.session_state.user_role = role
                    st.session_state.username = username
                    st.session_state.name = name
                    st.rerun()
                else:
                    st.error("Invalid credentials")
        
        st.info("**Demo Credentials:**\n- Admin: admin/admin123\n- User: user1/user123")
def generate_single_template(enhanced_mapper, template_path, mapping_results, single_row_df, images_to_use, debug_mode=False):
    """Generate a single template with comprehensive error handling and deep debugging"""
    try:
        if debug_mode:
            print(f"=== GENERATE_SINGLE_TEMPLATE DEBUG ===")
            print(f"Template path exists: {os.path.exists(template_path)}")
            print(f"Mapping results count: {len(mapping_results) if mapping_results else 0}")
            print(f"Data rows: {len(single_row_df)}")
            print(f"Images available: {len(images_to_use.get('all_sheets', {}))}")
            print(f"Enhanced mapper type: {type(enhanced_mapper)}")
            
            # Check if the method exists
            if hasattr(enhanced_mapper, 'fill_template_with_data_and_images'):
                print("✅ fill_template_with_data_and_images method found")
            else:
                print("❌ fill_template_with_data_and_images method NOT found")
                available_methods = [method for method in dir(enhanced_mapper) if 'fill' in method.lower()]
                print(f"Available 'fill' methods: {available_methods}")
        
        # Validate inputs with detailed logging
        if not template_path or not os.path.exists(template_path):
            error_msg = f'Template file not accessible: {template_path}'
            if debug_mode:
                print(f"❌ {error_msg}")
            return {'success': False, 'error': error_msg}
        
        if single_row_df.empty:
            error_msg = 'No data provided'
            if debug_mode:
                print(f"❌ {error_msg}")
            return {'success': False, 'error': error_msg}
        
        if not hasattr(enhanced_mapper, 'fill_template_with_data_and_images'):
            error_msg = 'Template filler method not available'
            if debug_mode:
                print(f"❌ {error_msg}")
            return {'success': False, 'error': error_msg}
        
        # Additional validation checks
        if not mapping_results:
            if debug_mode:
                print("⚠️ WARNING: No mapping results provided")
        
        # Try to get more info about the data
        if debug_mode:
            print(f"Single row data preview:")
            for col, val in single_row_df.iloc[0].items():
                print(f"  {col}: {val}")
        
        # Call the template filling method with detailed error tracking
        if debug_mode:
            print("🔄 Calling fill_template_with_data_and_images...")
        
        try:
            result = enhanced_mapper.fill_template_with_data_and_images(
                template_path, 
                mapping_results, 
                single_row_df, 
                images_to_use, 
                None  # No packaging type for now
            )
            
            if debug_mode:
                print(f"✅ Method call completed. Result type: {type(result)}")
                print(f"Result length: {len(result) if result and hasattr(result, '__len__') else 'N/A'}")
                
        except Exception as method_error:
            error_msg = f'Error calling fill_template_with_data_and_images: {str(method_error)}'
            if debug_mode:
                print(f"❌ {error_msg}")
                traceback.print_exc()
            return {'success': False, 'error': error_msg}
        
        # Validate result structure
        if not result:
            error_msg = 'Template filler returned None/empty result'
            if debug_mode:
                print(f"❌ {error_msg}")
            return {'success': False, 'error': error_msg}
        
        # Check if result is iterable and has expected structure
        try:
            if not hasattr(result, '__getitem__'):
                error_msg = f'Template filler returned non-indexable result: {type(result)}'
                if debug_mode:
                    print(f"❌ {error_msg}")
                return {'success': False, 'error': error_msg}
            
            if len(result) < 1:
                error_msg = 'Template filler returned empty result tuple'
                if debug_mode:
                    print(f"❌ {error_msg}")
                return {'success': False, 'error': error_msg}
                
        except Exception as structure_error:
            error_msg = f'Error checking result structure: {str(structure_error)}'
            if debug_mode:
                print(f"❌ {error_msg}")
            return {'success': False, 'error': error_msg}
        
        # Extract workbook with detailed checking
        try:
            workbook = result[0]
            if debug_mode:
                print(f"Workbook extracted. Type: {type(workbook)}")
                print(f"Workbook is None: {workbook is None}")
                
                if workbook:
                    print(f"Workbook has worksheets: {hasattr(workbook, 'worksheets')}")
                    if hasattr(workbook, 'worksheets'):
                        print(f"Number of worksheets: {len(workbook.worksheets)}")
                
        except (IndexError, TypeError) as extract_error:
            error_msg = f'Error extracting workbook from result: {str(extract_error)}'
            if debug_mode:
                print(f"❌ {error_msg}")
            return {'success': False, 'error': error_msg}
        
        if not workbook:
            error_msg = 'Template filler returned None workbook'
            if debug_mode:
                print(f"❌ {error_msg}")
                # Try to get more info about what was returned
                print(f"Full result: {result}")
            return {'success': False, 'error': error_msg}
        
        # Extract additional info from result with safe indexing
        try:
            filled_count = result[1] if len(result) > 1 else 0
            images_added = result[2] if len(result) > 2 else 0
            temp_files = result[3] if len(result) > 3 else []
            
            if debug_mode:
                print(f"✅ Template filled successfully:")
                print(f"  Filled count: {filled_count}")
                print(f"  Images added: {images_added}")
                print(f"  Temp files: {len(temp_files)}")
                
        except Exception as info_error:
            # If we can't extract additional info, that's okay as long as we have the workbook
            if debug_mode:
                print(f"⚠️ Warning: Could not extract additional info: {info_error}")
            filled_count = 0
            images_added = 0
            temp_files = []
        
        return {
            'success': True,
            'workbook': workbook,
            'filled_count': filled_count,
            'images_added': images_added,
            'temp_files': temp_files
        }
        
    except Exception as e:
        error_msg = f'Unexpected error in generate_single_template: {str(e)}'
        if debug_mode:
            print(f"❌ {error_msg}")
            traceback.print_exc()
        return {'success': False, 'error': error_msg}


def show_main_app():
    """Main application interface - ENHANCED DEBUGGING VERSION"""
    if 'enhanced_mapper' not in st.session_state:
        st.session_state.enhanced_mapper = EnhancedExcelMapper()
    st.title("🤖 Enhanced AI Template Mapper with Images")
    
    # Header with user info
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown(f"Welcome, **{st.session_state.name}** ({st.session_state.user_role})")
    with col3:
        if st.button("Logout"):
            st.session_state.authenticated = False
            st.session_state.user_role = None
            st.rerun()
    
    st.markdown("---")
    
    # Sidebar for file uploads
    with st.sidebar:
        st.header("📁 File Upload")
        
        # Template upload
        st.subheader("Excel Template")
        template_file = st.file_uploader(
            "Upload Excel Template",
            type=['xlsx', 'xls'],
            help="Upload the Excel template file"
        )
        
        # Data upload
        st.subheader("Data File")
        data_file = st.file_uploader(
            "Upload Data File",
            type=['xlsx', 'xls', 'csv'],
            help="Upload the data file to map to template"
        )
        
        # Image Source Selection
        st.markdown("---")
        st.header("🖼️ Image Processing Options")
        st.info("Choose how you want to handle images for your templates")
        
        image_source = st.radio(
            "Select Image Source:",
            [
                "🚫 No Images (Fastest)",
                "📤 Upload Images (Same for All Templates)", 
                "📊 Extract from Data File (Row-Specific)",
                "🔄 Both Upload + Extract (Advanced)"
            ],
            index=0,
            help="Choose your preferred image handling method"
        )
        
        # Show relevant options based on selection
        uploaded_images = {}
        
        if image_source == "📤 Upload Images (Same for All Templates)":
            st.subheader("🖼️ Bulk Image Upload")
            st.success("✅ BULK MODE: Upload images that apply to ALL templates")
            
            current_img = st.file_uploader("Current Packaging Image", type=['png', 'jpg', 'jpeg', 'gif', 'bmp'])
            primary_img = st.file_uploader("Primary Packaging Image", type=['png', 'jpg', 'jpeg', 'gif', 'bmp'])
            secondary_img = st.file_uploader("Secondary Packaging Image", type=['png', 'jpg', 'jpeg', 'gif', 'bmp'])
            label_img = st.file_uploader("Label Image", type=['png', 'jpg', 'jpeg', 'gif', 'bmp'])
            
            uploaded_images = {'current': current_img, 'primary': primary_img, 'secondary': secondary_img, 'label': label_img}
            
        elif image_source == "📊 Extract from Data File (Row-Specific)":
            st.subheader("📊 Auto-Extract Images")
            st.success("✅ EXTRACT MODE: Images will be extracted from Excel data file")
            st.info("💡 Make sure your data file is Excel format (.xlsx/.xls) with embedded images")
            
        elif image_source == "🔄 Both Upload + Extract (Advanced)":
            st.subheader("🔄 Advanced: Upload + Extract")
            st.warning("⚠️ HYBRID MODE: Both uploaded and extracted images will be used")
            
            current_img = st.file_uploader("Current Packaging Image (Bulk)", type=['png', 'jpg', 'jpeg', 'gif', 'bmp'])
            primary_img = st.file_uploader("Primary Packaging Image (Bulk)", type=['png', 'jpg', 'jpeg', 'gif', 'bmp'])
            secondary_img = st.file_uploader("Secondary Packaging Image (Bulk)", type=['png', 'jpg', 'jpeg', 'gif', 'bmp'])
            label_img = st.file_uploader("Label Image (Bulk)", type=['png', 'jpg', 'jpeg', 'gif', 'bmp'])
            
            uploaded_images = {'current': current_img, 'primary': primary_img, 'secondary': secondary_img, 'label': label_img}
            
        else:  # No Images
            st.subheader("🚫 No Images")
            st.info("✅ NO IMAGE MODE: Templates will be generated without images")
        
        # Enhanced Settings
        st.markdown("---")
        st.subheader("⚙️ Enhanced Debug Settings")
        
        # Debug mode toggle with more detail
        debug_mode = st.checkbox("🐛 Deep Debug Mode", value=True, help="Enable comprehensive debugging and error reporting")
        
        # Test with limited rows
        test_mode = st.checkbox("🧪 Test Mode (First 2 rows)", value=True, help="Test with only first 2 rows for faster debugging")
        
        # Enhanced mapper validation
        if st.checkbox("🔍 Validate Enhanced Mapper", value=False, help="Check enhanced mapper initialization"):
            if hasattr(st.session_state, 'enhanced_mapper'):
                st.success("✅ Enhanced mapper is available")
                mapper_methods = [method for method in dir(st.session_state.enhanced_mapper) if not method.startswith('_')]
                with st.expander("Available Methods"):
                    for method in sorted(mapper_methods):
                        st.write(f"• {method}")
            else:
                st.error("❌ Enhanced mapper not found in session state!")
                st.warning("This is likely the root cause of your issue. The enhanced_mapper needs to be initialized.")
        
        similarity_threshold = st.slider(
            "Similarity Threshold", min_value=0.1, max_value=1.0, value=0.3, step=0.1,
            help="Minimum similarity score for field matching"
        )
        
        if hasattr(st.session_state, 'enhanced_mapper'):
            st.session_state.enhanced_mapper.similarity_threshold = similarity_threshold
    
    # Main processing logic with enhanced debugging
    if template_file and data_file:
        st.header("🔍 Processing Files")
        
        # DIAGNOSTIC SECTION
        if debug_mode:
            st.subheader("🔧 System Diagnostics")
            
            # Check session state
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Session State Keys:**")
                for key in sorted(st.session_state.keys()):
                    st.write(f"• {key}")
            
            with col2:
                st.write("**Enhanced Mapper Status:**")
                if hasattr(st.session_state, 'enhanced_mapper'):
                    mapper = st.session_state.enhanced_mapper
                    st.write(f"• Type: {type(mapper)}")
                    st.write(f"• Has fill method: {hasattr(mapper, 'fill_template_with_data_and_images')}")
                    if hasattr(mapper, 'similarity_threshold'):
                        st.write(f"• Similarity threshold: {mapper.similarity_threshold}")
                else:
                    st.error("❌ enhanced_mapper not found!")
        
        # Continue with existing file processing logic...
        # [The rest of your existing code continues here]
        
        # Initialize variables
        extracted_images = {}
        data_df = pd.DataFrame()
        template_path = None
        bulk_images = {}
        processing_errors = []

        # STEP 1: Read and validate data file
        try:
            st.info("📖 Reading data file...")
            
            if data_file.name.endswith('.csv'):
                data_df = pd.read_csv(data_file)
            else:
                data_df = pd.read_excel(data_file)

            if data_df.empty:
                st.error("❌ Data file is empty!")
                return
            
            # Limit rows in test mode (reduced to 2 for faster testing)
            if test_mode:
                original_rows = len(data_df)
                data_df = data_df.head(2)
                st.warning(f"🧪 TEST MODE: Processing only {len(data_df)} rows (out of {original_rows})")
            
            st.success(f"✅ Data file loaded: {len(data_df)} rows, {len(data_df.columns)} columns")
            
            if debug_mode:
                st.write("**Data columns:**", list(data_df.columns))
                st.write("**First row sample:**", data_df.iloc[0].to_dict())

        except Exception as read_err:
            st.error(f"❌ Failed to read data file: {read_err}")
            if debug_mode:
                st.code(traceback.format_exc())
            return

        # STEP 2: Process template file with enhanced validation
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_template:
                tmp_template.write(template_file.getvalue())
                template_path = tmp_template.name
            
            st.info("📋 Analyzing template...")
            
            if debug_mode:
                st.write(f"**Template path:** {template_path}")
                st.write(f"**Template file exists:** {os.path.exists(template_path)}")
        
            # Critical check: enhanced_mapper availability
            if not hasattr(st.session_state, 'enhanced_mapper'):
                st.error("❌ CRITICAL: Enhanced mapper not initialized!")
                st.error("This is the root cause of the 'Template filler returned no workbook' error.")
                st.info("**Solution:** Ensure the enhanced_mapper is properly initialized in your session state before calling this function.")
                return
            
            # Verify the mapper has required methods
            mapper = st.session_state.enhanced_mapper
            required_methods = ['find_template_fields_with_context_and_images', 'map_data_with_section_context', 'fill_template_with_data_and_images']
            missing_methods = []
            
            for method in required_methods:
                if not hasattr(mapper, method):
                    missing_methods.append(method)
            
            if missing_methods:
                st.error(f"❌ Enhanced mapper is missing required methods: {missing_methods}")
                return
            else:
                if debug_mode:
                    st.success("✅ All required methods found on enhanced mapper")
        
            with st.spinner("Analyzing template fields..."):
                try:
                    template_fields, image_areas = mapper.find_template_fields_with_context_and_images(template_path)
                    
                    if debug_mode:
                        st.write(f"**Template analysis result:**")
                        st.write(f"• Fields found: {len(template_fields) if template_fields else 0}")
                        st.write(f"• Image areas found: {len(image_areas) if image_areas else 0}")
                        
                except Exception as template_err:
                    st.error(f"❌ Error analyzing template: {template_err}")
                    if debug_mode:
                        st.code(traceback.format_exc())
                    return
        
            if not template_fields:
                st.warning("⚠️ No mappable fields found in template")
                template_fields = {}
            else:
                st.success(f"✅ Found {len(template_fields)} mappable fields in template")

            # STEP 3: Map data to template
            st.info("🔗 Mapping data to template fields...")
            
            try:
                mapping_results = mapper.map_data_with_section_context(template_fields, data_df)
                
                if debug_mode:
                    st.write(f"**Mapping results:**")
                    st.write(f"• Mappings created: {len(mapping_results) if mapping_results else 0}")
                    
            except Exception as mapping_err:
                st.error(f"❌ Error during field mapping: {mapping_err}")
                if debug_mode:
                    st.code(traceback.format_exc())
                return
            
            if not mapping_results:
                st.warning("⚠️ No field mappings created")
                mapping_results = {}
            else:
                mapped_count = sum(1 for mapping in mapping_results.values() if mapping.get('is_mappable', False))
                st.success(f"✅ Successfully mapped {mapped_count} fields")

            # STEP 4: Generate templates with ENHANCED debugging
            st.subheader("🎯 Generate Templates")
            
            if st.button("🚀 Generate Templates (Enhanced Debug)", type="primary", use_container_width=True):
                
                # Enhanced pre-flight checks
                st.info("🔍 Running enhanced pre-flight checks...")
                
                preflight_errors = []
                
                if not hasattr(st.session_state, 'enhanced_mapper'):
                    preflight_errors.append("Enhanced mapper not available in session state")
                else:
                    mapper = st.session_state.enhanced_mapper
                    if not hasattr(mapper, 'fill_template_with_data_and_images'):
                        preflight_errors.append("fill_template_with_data_and_images method not found on mapper")
                
                if not template_path or not os.path.exists(template_path):
                    preflight_errors.append(f"Template file not accessible: {template_path}")
                
                if data_df.empty:
                    preflight_errors.append("No data to process")
                
                if preflight_errors:
                    st.error("❌ Enhanced pre-flight check failed:")
                    for error in preflight_errors:
                        st.write(f"• {error}")
                    return
                else:
                    st.success("✅ All pre-flight checks passed!")
                
                # Start generation with enhanced debugging
                with st.spinner(f"Generating {len(data_df)} templates with enhanced debugging..."):
                    try:
                        zip_buffer = io.BytesIO()
                        
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            successful_templates = 0
                            failed_templates = []
                            detailed_errors = []
                            
                            progress_bar = st.progress(0)
                            status_placeholder = st.empty()
                            debug_placeholder = st.empty() if debug_mode else None
                            
                            for index, row in data_df.iterrows():
                                try:
                                    # Update progress
                                    progress = (index + 1) / len(data_df)
                                    progress_bar.progress(progress)
                                    status_placeholder.text(f"Processing row {index + 1} of {len(data_df)}...")
                                    
                                    if debug_mode:
                                        debug_placeholder.info(f"🔍 DEBUG: Processing row {index + 1}")
                                    
                                    # Create single-row DataFrame
                                    single_row_df = pd.DataFrame([row])
                                    
                                    # Prepare images for this row (simplified for debugging)
                                    images_to_use = {}  # Start with no images for debugging
                                    
                                    # Generate template with enhanced debugging
                                    try:
                                        result = generate_single_template(
                                            st.session_state.enhanced_mapper,
                                            template_path,
                                            mapping_results,
                                            single_row_df,
                                            images_to_use,
                                            debug_mode
                                        )
                                        
                                        if debug_mode:
                                            debug_placeholder.write(f"Row {index + 1} result: {result.get('success', False)}")
                                        
                                        if result and result.get('success', False):
                                            workbook = result['workbook']
                                            
                                            # Generate filename
                                            filename = generate_safe_filename(row, data_df.columns, index)
                                            
                                            # Save to zip
                                            template_buffer = io.BytesIO()
                                            workbook.save(template_buffer)
                                            template_buffer.seek(0)
                                            
                                            zip_file.writestr(filename, template_buffer.getvalue())
                                            
                                            # Cleanup
                                            cleanup_temp_files(result.get('temp_files', []))
                                            workbook.close()
                                            
                                            successful_templates += 1
                                        
                                        else:
                                            error_msg = result.get('error', 'Unknown error') if result else 'No result returned'
                                            failed_templates.append(index + 1)
                                            detailed_errors.append(f"Row {index + 1}: {error_msg}")
                                    
                                    except Exception as template_err:
                                        error_msg = f"Template generation exception: {str(template_err)}"
                                        failed_templates.append(index + 1)
                                        detailed_errors.append(f"Row {index + 1}: {error_msg}")
                                        
                                        if debug_mode:
                                            debug_placeholder.error(f"Row {index + 1} exception: {template_err}")
                                
                                except Exception as row_err:
                                    error_msg = f"Row processing exception: {str(row_err)}"
                                    failed_templates.append(index + 1)
                                    detailed_errors.append(f"Row {index + 1}: {error_msg}")

                        # Finalize results
                        progress_bar.progress(1.0)
                        status_placeholder.text("Finalizing...")
                        
                        if debug_mode and debug_placeholder:
                            debug_placeholder.empty()

                        if successful_templates > 0:
                            zip_buffer.seek(0)
                            
                            st.success(f"✅ Successfully generated {successful_templates} templates!")
                            
                            if failed_templates:
                                st.warning(f"⚠️ Failed to generate {len(failed_templates)} templates")
                                
                                with st.expander("Detailed Error Report", expanded=True):
                                    for error in detailed_errors:
                                        st.write(f"• {error}")
                            
                            # Download button
                            st.download_button(
                                label=f"📦 Download {successful_templates} Templates (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name=f"filled_templates_{successful_templates}_files.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                        
                        else:
                            st.error("❌ No templates were successfully generated!")
                            
                            st.error("**Detailed Error Analysis:**")
                            for error in detailed_errors:
                                st.write(f"• {error}")

                    except Exception as e:
                        st.error(f"❌ Critical error during template generation: {e}")
                        if debug_mode:
                            st.code(traceback.format_exc())

        except Exception as e:
            st.error(f"❌ Error processing template file: {e}")
            if debug_mode:
                st.code(traceback.format_exc())
        
        finally:
            # Clean up template file
            if template_path and os.path.exists(template_path):
                try:
                    os.unlink(template_path)
                except Exception as cleanup_err:
                    if debug_mode:
                        st.warning(f"⚠️ Could not delete temp template file: {cleanup_err}")
      
    else:
        # Show enhanced instructions when no files uploaded
        st.info("👆 Please upload both an Excel template and a data file to begin")
        
        st.markdown("""
        ### 🚨 **ROOT CAUSE ANALYSIS**
        
        The error "Template filler returned no workbook" typically occurs when:
        
        1. **❌ Enhanced Mapper Not Initialized**: The `st.session_state.enhanced_mapper` is not properly set up
        2. **❌ Missing Required Methods**: The mapper doesn't have the `fill_template_with_data_and_images` method
        3. **❌ Template File Issues**: The template file is corrupted or has an unexpected format
        4. **❌ Method Return Format**: The fill method is returning an unexpected result structure
        
        ### 🔧 **DEBUGGING STEPS**
        
        1. **Enable Deep Debug Mode** in the sidebar
        2. **Validate Enhanced Mapper** using the checkbox in settings
        3. **Use Test Mode** to process only 2 rows
        4. **Check the System Diagnostics** section when files are uploaded
        
        ### 🎯 **RECOMMENDED FIX**
        
        Make sure your `enhanced_mapper` is properly initialized before calling this function:
        
        ```python
        # Ensure this is done BEFORE calling show_main_app()
        if 'enhanced_mapper' not in st.session_state:
            st.session_state.enhanced_mapper = YourEnhancedMapperClass()
        ```
        """)


# Keep all your existing utility functions
def prepare_images_for_row(image_source, bulk_images, extracted_images, row, columns, debug_mode=False):
    """Prepare images for a specific row based on the selected mode"""
    try:
        if debug_mode:
            print(f"Preparing images - Mode: {image_source}")
        
        if image_source == "🚫 No Images (Fastest)":
            return {}
        
        elif image_source == "📤 Upload Images (Same for All Templates)":
            if bulk_images:
                return {'all_sheets': bulk_images.copy()}
            else:
                return {}
        
        elif image_source == "📊 Extract from Data File (Row-Specific)":
            if extracted_images:
                return filter_images_for_row(extracted_images, row, columns)
            else:
                return {}
        
        elif image_source == "🔄 Both Upload + Extract (Advanced)":
            combined_images = {}
            
            # Add bulk images
            if bulk_images:
                combined_images.update(bulk_images)
            
            # Add extracted images for this row
            if extracted_images:
                row_images = filter_images_for_row(extracted_images, row, columns)
                if row_images and 'all_sheets' in row_images:
                    for img_key, img_data in row_images['all_sheets'].items():
                        img_type = img_data.get('type', 'unknown')
                        
                        # Avoid conflicts with bulk images
                        bulk_has_type = any(
                            bulk_img.get('type', '') == img_type 
                            for bulk_img in bulk_images.values()
                        ) if bulk_images else False
                        
                        if bulk_has_type:
                            new_key = f"extracted_{img_type}_{img_key}"
                            combined_images[new_key] = img_data
                        else:
                            combined_images[img_key] = img_data
            
            return {'all_sheets': combined_images} if combined_images else {}
        
        else:
            return {}
    
    except Exception as e:
        print(f"Error preparing images: {e}")
        return {}


def cleanup_temp_files(temp_files):
    """Safely clean up temporary files"""
    for file_path in temp_files:
        try:
            if os.path.exists(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f"Warning: Could not delete temp file {file_path}: {e}")


def generate_safe_filename(row, columns, index):
    """Generate a safe filename with fallback options"""
    try:
        # Try to get meaningful identifiers
        identifiers = []
        
        # Look for common ID fields
        id_fields = ['id', 'part_no', 'partno', 'part_number', 'code', 'sku']
        for field in id_fields:
            for col in columns:
                if field in col.lower():
                    value = row.get(col)
                    if pd.notna(value) and str(value).strip():
                        clean_value = re.sub(r'[^\w\-_\.]', '_', str(value))[:20]
                        identifiers.append(clean_value)
                        break
            if identifiers:
                break
        
        # If no identifier found, use index
        if not identifiers:
            identifiers.append(f"row_{index + 1}")
        
        # Create filename
        filename = "_".join(identifiers) + "_template.xlsx"
        
        # Ensure it's not too long
        if len(filename) > 100:
            filename = filename[:90] + "_template.xlsx"
        
        return filename
        
    except Exception as e:
        print(f"Error generating filename: {e}")
        return f"template_row_{index + 1}.xlsx"


def filter_images_for_row(extracted_images, row, columns):
    """Filter extracted images for a specific row - simplified version"""
    try:
        if not extracted_images or 'all_sheets' not in extracted_images:
            return {}
        
        # For now, return all extracted images
        # In a production version, you'd implement filtering by part number/description
        return extracted_images
        
    except Exception as e:
        print(f"Error filtering images: {e}")
        return {}


def main():
    """Main application entry point with error handling"""
    try:
        # Initialize session state if needed
        if 'authenticated' not in st.session_state:
            st.session_state.authenticated = False
        
        # CRITICAL: Check for enhanced_mapper initialization
        if 'enhanced_mapper' not in st.session_state:
            st.error("🚨 CRITICAL ERROR: Enhanced mapper not found in session state!")
            st.error("This is the root cause of the 'Template filler returned no workbook' error.")
            
            st.markdown("""
            ### 🔧 **IMMEDIATE SOLUTION REQUIRED**
            
            Before running this application, you must initialize the enhanced_mapper:
            
            ```python
            # Add this to your main initialization code:
            if 'enhanced_mapper' not in st.session_state:
                st.session_state.enhanced_mapper = YourEnhancedMapperClass()
            ```
            
            **What you need to do:**
            1. Import your enhanced mapper class
            2. Initialize it and store in session state
            3. Ensure it has these required methods:
               - `find_template_fields_with_context_and_images()`
               - `map_data_with_section_context()`
               - `fill_template_with_data_and_images()`
            
            **Example initialization:**
            ```python
            from your_module import EnhancedTemplateMapper
            
            # Initialize the mapper
            if 'enhanced_mapper' not in st.session_state:
                st.session_state.enhanced_mapper = EnhancedTemplateMapper()
            ```
            """)
            
            st.stop()  # Stop execution until this is fixed
        
        if not st.session_state.authenticated:
            show_login()
        else:
            show_main_app()
            
    except Exception as e:
        st.error(f"Application Error: {e}")
        st.exception(e)


# Additional diagnostic function
def diagnose_enhanced_mapper():
    """Comprehensive diagnostic function for enhanced mapper"""
    st.subheader("🔍 Enhanced Mapper Diagnostics")
    
    if 'enhanced_mapper' not in st.session_state:
        st.error("❌ Enhanced mapper not found in session state")
        return False
    
    mapper = st.session_state.enhanced_mapper
    
    # Check basic properties
    st.write(f"**Mapper Type:** {type(mapper)}")
    st.write(f"**Mapper ID:** {id(mapper)}")
    
    # Check required methods
    required_methods = [
        'find_template_fields_with_context_and_images',
        'map_data_with_section_context', 
        'fill_template_with_data_and_images'
    ]
    
    method_status = {}
    for method in required_methods:
        has_method = hasattr(mapper, method)
        method_status[method] = has_method
        
        if has_method:
            st.success(f"✅ {method}")
            
            # Try to get method signature if possible
            try:
                import inspect
                sig = inspect.signature(getattr(mapper, method))
                st.write(f"   Signature: {sig}")
            except:
                pass
        else:
            st.error(f"❌ {method}")
    
    # Check additional useful properties
    optional_properties = ['similarity_threshold', 'image_extractor']
    for prop in optional_properties:
        if hasattr(mapper, prop):
            st.info(f"ℹ️ Has {prop}: {getattr(mapper, prop)}")
    
    # Overall health check
    all_required_present = all(method_status.values())
    
    if all_required_present:
        st.success("🎉 Enhanced mapper appears to be properly initialized!")
        return True
    else:
        missing = [method for method, present in method_status.items() if not present]
        st.error(f"❌ Missing required methods: {missing}")
        return False


if __name__ == "__main__":
    main()
