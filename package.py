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
        """Add uploaded images to template - COMPLETELY REWRITTEN FOR RELIABILITY"""
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
                    # Move to next position for row 42 (your spacing calculation)
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
        """Place a single image at the specified cell position"""
        try:
            print(f"  Placing {img_key} at {cell_position} ({width_cm}x{height_cm}cm)")
            
            # Create temporary image file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                image_bytes = base64.b64decode(img_data['data'])
                tmp_img.write(image_bytes)
                tmp_img_path = tmp_img.name
            
            # Create openpyxl image object
            img = OpenpyxlImage(tmp_img_path)
            
            # Set image size (converting cm to pixels: 1cm ≈ 37.8 pixels)
            img.width = int(width_cm * 37.8)
            img.height = int(height_cm * 37.8)
            
            # Set position using simple anchor
            img.anchor = cell_position
            
            # Add image to worksheet
            worksheet.add_image(img)
            
            # Track temporary file for cleanup
            temp_image_paths.append(tmp_img_path)
            
            print(f"    ✅ Successfully placed {img_key} at {cell_position}")
            return True
            
        except Exception as e:
            print(f"    ❌ Failed to place {img_key} at {cell_position}: {e}")
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
        """Fill template with mapped data, images, and procedure steps"""
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
        
            filled_count = 0
            images_added = 0
            procedure_steps_added = 0
            temp_image_paths = []
        
            # Create data dictionary for procedure step replacement
            data_dict = {}
            if len(data_df) > 0:
                for col in data_df.columns:
                    try:
                        data_dict[col] = data_df.iloc[0][col]
                    except:
                        data_dict[col] = 'XXX'
        
            # Fill data fields
            for coord, mapping in mapping_results.items():
                try:
                    if mapping['data_column'] is not None and mapping['is_mappable']:
                        field_info = mapping['field_info']
                    
                        target_cell = self.find_data_cell_for_label(worksheet, field_info)
                    
                        if target_cell and len(data_df) > 0:
                            data_value = data_df.iloc[0][mapping['data_column']]
                        
                            cell_obj = worksheet[target_cell]
                            if hasattr(cell_obj, '__class__') and cell_obj.__class__.__name__ == 'MergedCell':
                                for merged_range in worksheet.merged_cells.ranges:
                                    if target_cell in merged_range:
                                        anchor_cell = merged_range.start_cell
                                        anchor_cell.value = str(data_value) if not pd.isna(data_value) else ""
                                        break
                            else:
                                cell_obj.value = str(data_value) if not pd.isna(data_value) else ""
                            filled_count += 1
                        
                except Exception as e:
                    st.error(f"Error filling mapping {coord}: {e}")
                    continue
        
            # Add images if provided
            if uploaded_images:
                # First, identify image upload areas
                _, image_areas = self.find_template_fields_with_context_and_images(template_file)
                images_added, temp_image_paths = self.image_extractor.add_images_to_template(worksheet, uploaded_images, image_areas)
        
            # Write procedure steps if packaging type is provided
            if packaging_type and packaging_type != "Select Packaging Procedure":
                try:
                    procedure_steps_added = self.write_procedure_steps_to_template(worksheet, packaging_type, data_dict)
                    print(f"Added {procedure_steps_added} procedure steps for packaging type: {packaging_type}")
                except Exception as e:
                    st.error(f"Error adding procedure steps: {e}")
                    print(f"Error adding procedure steps: {e}")
                    procedure_steps_added = 0
                    if packaging_type and packaging_type != "Select Packaging Procedure":
                        try:
                            # Create data dictionary for procedure step replacement
                            data_dict = {}
                            if len(data_df) > 0:
                                for col in data_df.columns:
                                    try:
                                        data_dict[col] = data_df.iloc[0][col]
                                    except:
                                        data_dict[col] = 'XXX'
                
                            procedure_steps_added = self.write_procedure_steps_to_template(worksheet, packaging_type, data_dict)
                            print(f"Added {procedure_steps_added} procedure steps for packaging type: {packaging_type}")
                        except Exception as e:
                            st.error(f"Error adding procedure steps: {e}")
                            print(f"Error adding procedure steps: {e}")
                            procedure_steps_added = 0
            
            return workbook, filled_count, images_added, temp_image_paths, procedure_steps_added
        
        except Exception as e:
            st.error(f"Error filling template: {e}")
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

def show_main_app():
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
            help="Upload the data file to map to template (images will be extracted from Excel files)"
        )
        
        # 🆕 NEW: Bulk Image Upload Section
        st.subheader("🖼️ Bulk Image Upload (Same for All)")
        st.info("Upload 4 images that will be applied to ALL templates")
        
        # Individual image uploaders for each type
        current_img = st.file_uploader(
            "Current Packaging Image",
            type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
            help="Image for current packaging (Position: T3)"
        )
        
        primary_img = st.file_uploader(
            "Primary Packaging Image", 
            type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
            help="Image for primary packaging (Position: Row 42, Col A)"
        )
        
        secondary_img = st.file_uploader(
            "Secondary Packaging Image",
            type=['png', 'jpg', 'jpeg', 'gif', 'bmp'], 
            help="Image for secondary packaging (Position: Row 42, next column)"
        )
        
        label_img = st.file_uploader(
            "Label Image",
            type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
            help="Image for label (Position: Row 42, next column)"
        )
        
        # Settings
        st.subheader("⚙️ Settings")
        similarity_threshold = st.slider(
            "Similarity Threshold",
            min_value=0.1,
            max_value=1.0,
            value=0.3,
            step=0.1,
            help="Minimum similarity score for field matching"
        )
        
        st.session_state.enhanced_mapper.similarity_threshold = similarity_threshold
    
    if template_file and data_file:
        extracted_images = {}
        data_df = pd.DataFrame()
        template_path = None
        bulk_images = {}

        # Process bulk uploaded images
        bulk_upload_images = {
            'current': current_img,
            'primary': primary_img, 
            'secondary': secondary_img,
            'label': label_img
        }
        
        # Convert uploaded images to the format expected by the system
        for img_type, uploaded_file in bulk_upload_images.items():
            if uploaded_file is not None:
                try:
                    # Read the uploaded file
                    image_bytes = uploaded_file.read()
                    
                    # Create PIL Image
                    pil_image = Image.open(io.BytesIO(image_bytes))
                    
                    # Convert to base64
                    buffered = io.BytesIO()
                    pil_image.save(buffered, format="PNG")
                    img_str = base64.b64encode(buffered.getvalue()).decode()
                    
                    # Create image data structure
                    image_key = f"bulk_{img_type}_0"
                    bulk_images[image_key] = {
                        'data': img_str,
                        'format': 'PNG',
                        'size': pil_image.size,
                        'position': f'BULK_{img_type.upper()}',
                        'sheet': 'BULK_UPLOAD',
                        'index': 0,
                        'type': img_type,
                        'hash': hashlib.md5(image_bytes).hexdigest()
                    }
                    
                    st.success(f"✅ {img_type.title()} image loaded: {uploaded_file.name}")
                    
                except Exception as e:
                    st.error(f"❌ Error processing {img_type} image: {e}")

        # Show bulk uploaded images preview
        if bulk_images:
            st.subheader("🖼️ Bulk Upload Preview")
            st.info(f"These {len(bulk_images)} images will be applied to ALL templates")
            
            cols = st.columns(min(4, len(bulk_images)))
            for idx, (img_key, img_data) in enumerate(bulk_images.items()):
                with cols[idx % 4]:
                    try:
                        img_bytes = base64.b64decode(img_data['data'])
                        st.image(img_bytes, width=150, caption=f"{img_data['type'].title()}")
                        st.write(f"Size: {img_data['size']}")
                    except Exception as e:
                        st.error(f"Error displaying {img_key}: {e}")

        # ✅ 1. Extract images from data file (Excel only)
        if data_file.name.endswith(('.xlsx', '.xls')):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_data:
                    tmp_data.write(data_file.getvalue())
                    data_path = tmp_data.name

                st.info("🔍 Extracting images from data file...")

                try:
                    with st.spinner("Extracting images from Excel file..."):
                        # Store Excel path for image classification
                        st.session_state.enhanced_mapper.image_extractor.current_excel_path = data_path

                        extracted_images = st.session_state.enhanced_mapper.image_extractor.extract_images_from_excel(data_path)
                        st.success(f"✅ Extracted {sum(len(sheet_images) for sheet_images in extracted_images.values())} images.")

                except Exception as extract_err:
                    st.error(f"❌ Error during image extraction: {extract_err}")
                    st.code(traceback.format_exc())

                # Clean up temp data file
                try:
                    os.unlink(data_path)
                except Exception as cleanup_err:
                    print(f"⚠️ Could not delete temp file: {cleanup_err}")

            except Exception as e:
                st.error(f"❌ Unexpected error while saving data file: {e}")
                st.code(traceback.format_exc())

        # ✅ 2. Read data file into DataFrame
        try:
            if data_file.name.endswith('.csv'):
                data_df = pd.read_csv(data_file)
            else:
                data_df = pd.read_excel(data_file)

            st.info(f"📊 Data file contains {len(data_df)} rows of data")

        except Exception as read_err:
            st.error(f"❌ Failed to read data file: {read_err}")
            st.code(traceback.format_exc())
            data_df = pd.DataFrame()

        # ✅ 3. Save template file and process it
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_template:
                tmp_template.write(template_file.getvalue())
                template_path = tmp_template.name
            
            # ✅ TEMPLATE PROCESSING
            st.subheader("📋 Template Analysis")
        
            with st.spinner("Analyzing template fields and image areas..."):
                template_fields, image_areas = st.session_state.enhanced_mapper.find_template_fields_with_context_and_images(template_path)
        
            if template_fields:
                st.success(f"Found {len(template_fields)} mappable fields")
                
                # Show template fields
                with st.expander("Template Fields Details", expanded=False):
                    fields_df = pd.DataFrame([
                        {
                            'Position': coord,
                            'Field': field['value'],
                            'Section': field.get('section_context', 'Unknown'),
                            'Row': field['row'],
                            'Column': field['column']
                        }
                        for coord, field in template_fields.items()
                    ])
                    st.dataframe(fields_df, use_container_width=True)
                
                # Show image areas
                if image_areas:
                    st.info(f"Found {len(image_areas)} image upload areas in template")
                    with st.expander("Image Upload Areas", expanded=False):
                        image_df = pd.DataFrame(image_areas)
                        st.dataframe(image_df, use_container_width=True)
                
                # 🆕 UPDATED: Show combined image usage
                total_extracted = 0
                if extracted_images:
                    total_extracted = sum(len(sheet_images) for sheet_images in extracted_images.values())
                
                if bulk_images and extracted_images and total_extracted > 0:
                    st.success(f"🖼️ Using COMBINED images: {len(bulk_images)} bulk uploaded + {total_extracted} extracted from data file")
                    st.info("📸 **HYBRID MODE**: Bulk images + Row-specific extracted images will be combined")
                    
                    # Show breakdown
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**Bulk Images (Same for All):**")
                        for img_type in ['current', 'primary', 'secondary', 'label']:
                            if any(img_type in key for key in bulk_images.keys()):
                                st.write(f"• {img_type.title()} packaging")
                    
                    with col2:
                        st.write("**Extracted Images (Row-Specific):**")
                        st.write(f"• {total_extracted} images from data file")
                        st.write("• Filtered by part number/description")
                
                elif bulk_images:
                    st.success(f"🖼️ Using {len(bulk_images)} bulk uploaded images for ALL templates")
                    st.info("🎯 **BULK MODE**: All templates will use the same uploaded images")
                
                elif extracted_images and total_extracted > 0:
                    st.success(f"🖼️ Extracted {total_extracted} images from data file")
                    st.info("🎯 **EXTRACTION MODE**: Images filtered by part number/description for each row")
                    
                    with st.expander("Extracted Images from Data File", expanded=True):
                        for sheet_name, sheet_images in extracted_images.items():
                            if sheet_images:
                                st.write(f"**Sheet: {sheet_name}**")
                                cols = st.columns(min(3, len(sheet_images)))
                                
                                for idx, (position, img_data) in enumerate(sheet_images.items()):
                                    with cols[idx % 3]:
                                        st.write(f"Position: {position}")
                                        try:
                                            img_bytes = base64.b64decode(img_data['data'])
                                            st.image(img_bytes, width=150)
                                            st.write(f"Size: {img_data['size']}")
                                            st.write(f"Type: {img_data.get('type', 'Unknown')}")
                                        except Exception as img_err:
                                            st.error(f"Error displaying image: {img_err}")
                else:
                    if data_file.name.endswith(('.xlsx', '.xls')):
                        st.info("No images found in the data file and no bulk images uploaded")
                    else:
                        st.info("CSV files don't contain images. Use Excel files to include images or upload bulk images above.")
                
                # Data mapping - using first row to establish mapping
                st.subheader("🔗 Field Mapping")
                
                with st.spinner("Mapping template fields to data columns..."):
                    mapping_results = st.session_state.enhanced_mapper.map_data_with_section_context(
                        template_fields, data_df
                    )
                
                if mapping_results:
                    # Show mapping results
                    mapping_df = pd.DataFrame([
                        {
                            'Template Field': mapping['template_field'],
                            'Data Column': mapping['data_column'] if mapping['data_column'] else 'No Match',
                            'Similarity': f"{mapping['similarity']:.2f}" if mapping['similarity'] > 0 else "0.00",
                            'Section': mapping.get('section_context', 'Unknown'),
                            'Status': '✅ Mapped' if mapping['is_mappable'] else '❌ No Match'
                        }
                        for mapping in mapping_results.values()
                    ])
                    
                    st.dataframe(mapping_df, use_container_width=True)
                    
                    # ✨ ENHANCED PACKAGING PROCEDURE SECTION
                    st.subheader("📋 Packaging Procedure Configuration")

                    # Create two columns for better layout
                    col1, col2 = st.columns([1, 2])

                    with col1:
                        st.write("**Select Packaging Type:**")
                        procedure_type = st.selectbox(
                            "Packaging Procedure Type",
                            ["Select Packaging Procedure"] + list(st.session_state.enhanced_mapper.packaging_procedures.keys()),
                            help="Select a packaging type to auto-populate procedure steps"
                        )
                        
                        # Add option to preview steps without adding to data
                        preview_only = st.checkbox(
                            "Preview Only", 
                            value=False, 
                            help="Check to preview steps without adding them to the template"
                        )

                    with col2:
                        if procedure_type and procedure_type != "Select Packaging Procedure":
                            st.info(f"**Selected:** {procedure_type}")
                            
                            # Get procedure steps with data substitution (using first row as example)
                            try:
                                data_dict = data_df.iloc[0].to_dict() if len(data_df) > 0 else {}
                                procedures = st.session_state.enhanced_mapper.get_procedure_steps(procedure_type, data_dict)
                                
                                st.write("**Procedure Steps Preview (using first row data):**")
                                
                                # Display steps in a more organized way
                                steps_container = st.container()
                                with steps_container:
                                    for i, step in enumerate(procedures, 1):
                                        if step.strip():
                                            # Color-code different types of steps
                                            if any(keyword in step.lower() for keyword in ['pick up', 'apply', 'put']):
                                                st.markdown(f"🟢 **{i}.** {step}")
                                            elif any(keyword in step.lower() for keyword in ['seal', 'load', 'attach']):
                                                st.markdown(f"🔵 **{i}.** {step}")
                                            elif any(keyword in step.lower() for keyword in ['ensure', 'prepare']):
                                                st.markdown(f"🟡 **{i}.** {step}")
                                            else:
                                                st.markdown(f"**{i}.** {step}")
                                
                                # Show statistics
                                non_empty_steps = [step for step in procedures if step.strip()]
                                st.write(f"**Total Steps:** {len(non_empty_steps)}")
                                
                            except Exception as e:
                                st.error(f"Error generating procedure steps: {e}")
        
                    # Fill templates for ALL rows
                    st.subheader("📝 Generate Multiple Filled Templates")
                    
                    # Show what will be included in the templates
                    st.write("**Each template will include:**")
                    include_items = []
                    
                    # Count mapped fields
                    mapped_count = sum(1 for mapping in mapping_results.values() if mapping['is_mappable'])
                    if mapped_count > 0:
                        include_items.append(f"📊 {mapped_count} mapped data fields")
                    
                    # Count images
                    if bulk_images and extracted_images and total_extracted > 0:
                        include_items.append(f"🖼️ COMBINED: {len(bulk_images)} bulk + extracted images per row")
                        st.info("🔄 **HYBRID MODE**: Each template gets both bulk and row-specific images")
                    elif bulk_images:
                        include_items.append(f"🖼️ {len(bulk_images)} bulk uploaded images (SAME for all templates)")
                        st.info("🎯 **BULK MODE**: All templates will use the same uploaded images")
                    elif extracted_images and total_extracted > 0:
                        include_items.append(f"🖼️ Images matching each row's part number/description")
                        st.info("🎯 **EXTRACTION MODE**: Row-specific images from data file")
                    
                    # Count procedure steps
                    if procedure_type and procedure_type != "Select Packaging Procedure" and not preview_only:
                        try:
                            steps = st.session_state.enhanced_mapper.get_procedure_steps(procedure_type, data_df.iloc[0].to_dict())
                            step_count = len([s for s in steps if s.strip()])
                            if step_count > 0:
                                include_items.append(f"📋 {step_count} packaging procedure steps")
                        except:
                            pass
                    
                    if include_items:
                        for item in include_items:
                            st.write(f"• {item}")
                    else:
                        st.warning("No items will be added to the templates")
                    
                    # Show file generation info
                    st.info(f"🎯 Will generate {len(data_df)} separate template files (one for each data row)")
                    
                    if st.button("Generate All Filled Templates", type="primary", use_container_width=True):
                        with st.spinner(f"Generating {len(data_df)} filled templates..."):
                            try:
                                # Create a zip file to contain all templates
                                zip_buffer = io.BytesIO()
                                
                                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                    successful_templates = 0
                                    failed_templates = []
                                    
                                    # Progress bar
                                    progress_bar = st.progress(0)
                                    status_placeholder = st.empty()
                                    
                                    for index, row in data_df.iterrows():
                                        try:
                                            # Update progress
                                            progress = (index + 1) / len(data_df)
                                            progress_bar.progress(progress)
                                            status_placeholder.text(f"Processing row {index + 1} of {len(data_df)}...")
                                            
                                            # Create a single-row dataframe for this iteration
                                            single_row_df = pd.DataFrame([row])
                                            
                                            # Add procedure steps to this row if selected
                                            if procedure_type and procedure_type != "Select Packaging Procedure" and not preview_only:
                                                try:
                                                    data_dict = row.to_dict()
                                                    procedure_steps = st.session_state.enhanced_mapper.get_procedure_steps(procedure_type, data_dict)
                                                    
                                                    # Add procedure steps to the single row dataframe
                                                    for i, step in enumerate(procedure_steps, 1):
                                                        if step.strip():  # Only add non-empty steps
                                                            single_row_df.loc[0, f'Procedure Step {i}'] = step
                                                except Exception as e:
                                                    st.warning(f"Failed to add procedure steps for row {index + 1}: {e}")
                                            
                                            # 🆕 UPDATED: COMBINE BOTH IMAGE SOURCES
                                            combined_images = {}
                                            
                                            # First, add extracted images (row-specific)
                                            if extracted_images:
                                                row_images = filter_images_for_row(extracted_images, row, data_df.columns)
                                                if row_images and 'all_sheets' in row_images:
                                                    combined_images.update(row_images['all_sheets'])
                                                    st.write(f"Added {len(row_images.get('all_sheets', {}))} extracted images for row {index + 1}")
                                            
                                            # Then, add bulk images (these will be the same for all rows)
                                            if bulk_images:
                                                combined_images.update(bulk_images)
                                                st.write(f"Added {len(bulk_images)} bulk images for row {index + 1}")
                                            
                                            # Prepare final image structure
                                            if combined_images:
                                                images_to_use = {'all_sheets': combined_images}
                                                st.write(f"Total images for template {index + 1}: {len(combined_images)}")
                                            else:
                                                images_to_use = {}
                                            
                                            # Pass the packaging type to the fill function
                                            selected_packaging_type = procedure_type if (procedure_type and procedure_type != "Select Packaging Procedure" and not preview_only) else None
                                            
                                            # Fill template for this specific row with appropriate images
                                            result = st.session_state.enhanced_mapper.fill_template_with_data_and_images(
                                                template_path, mapping_results, single_row_df, images_to_use, selected_packaging_type
                                            )
                                            
                                            workbook, filled_count, images_added, temp_image_paths, procedure_steps_added = result
                                            
                                            if workbook:
                                                # 🎯 ENHANCED FILENAME GENERATION
                                                filename = generate_enhanced_filename(row, data_df.columns, index)
                                                
                                                # Save workbook to memory
                                                template_buffer = io.BytesIO()
                                                workbook.save(template_buffer)
                                                template_buffer.seek(0)
                                                
                                                # Add to zip file
                                                zip_file.writestr(filename, template_buffer.getvalue())
                                                
                                                # Clean up temporary image files
                                                for path in temp_image_paths:
                                                    try:
                                                        os.unlink(path)
                                                    except Exception as e:
                                                        pass
                                                
                                                workbook.close()
                                                successful_templates += 1
                                                
                                            else:
                                                failed_templates.append(index + 1)
                                                
                                        except Exception as e:
                                            failed_templates.append(index + 1)
                                            st.warning(f"Failed to process row {index + 1}: {e}")
                                    
                                    # Clear progress indicators
                                    progress_bar.empty()
                                    status_placeholder.empty()
                                
                                zip_buffer.seek(0)
                                
                                # Show results
                                if successful_templates > 0:
                                    st.success(f"🎉 Successfully generated {successful_templates} template files!")
                                    
                                    # Show stats
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("Successful Templates", successful_templates)
                                    with col2:
                                        st.metric("Failed Templates", len(failed_templates))
                                    with col3:
                                        st.metric("Success Rate", f"{(successful_templates/len(data_df)*100):.1f}%")
                                    
                                    if failed_templates:
                                        st.warning(f"⚠️ Failed to generate templates for rows: {', '.join(map(str, failed_templates))}")
                                    
                                    # 🆕 UPDATED: SHOW COMBINED IMAGE USAGE SUMMARY
                                    if bulk_images and extracted_images:
                                        bulk_count = len(bulk_images)
                                        extracted_count = sum(len(sheet_images) for sheet_images in extracted_images.values())
                                        st.info(f"📸 All {successful_templates} templates used COMBINED images: {bulk_count} bulk + {extracted_count} extracted (per row)")
                                    elif bulk_images:
                                        st.info(f"📸 All {successful_templates} templates used the same {len(bulk_images)} bulk uploaded images")
                                    elif extracted_images:
                                        total_extracted = sum(len(sheet_images) for sheet_images in extracted_images.values())
                                        st.info(f"📸 Templates used row-specific images from {total_extracted} extracted images")
                                    
                                    # Download button for zip file
                                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                    zip_filename = f"filled_templates_{timestamp}.zip"
                                    
                                    st.download_button(
                                        label=f"📥 Download All Templates ({successful_templates} files)",
                                        data=zip_buffer.getvalue(),
                                        file_name=zip_filename,
                                        mime="application/zip",
                                        use_container_width=True
                                    )
                                    
                                else:
                                    st.error("❌ Failed to generate any templates")
                                    
                            except Exception as e:
                                st.error(f"Error generating templates: {e}")
                                st.exception(e)
                
                else:
                    st.warning("No mapping results generated")
            
            else:
                st.warning("No mappable fields found in template")
            
            # Clean up temporary template file
            try:
                if template_path:
                    os.unlink(template_path)
            except Exception as cleanup_err:
                print(f"Could not cleanup template file: {cleanup_err}")
                
        except Exception as template_err:
            st.error(f"❌ Failed to process template file: {template_err}")
            st.code(traceback.format_exc())
            # Clean up on error too
            try:
                if template_path:
                    os.unlink(template_path)
            except:
                pass
                
    else:
        st.info("👆 Please upload both an Excel template and a data file to begin")
        
        # Show demo information with updated features
        st.markdown("### 🎯 Features")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **Template Processing:**
            - 📋 Smart field detection
            - 🎯 Section-aware mapping
            - 🔄 Merged cell handling
            - 📏 Packaging-specific patterns
            - 🗂️ Multi-template generation
            - 🎯 **Row-specific image filtering**
            - 🆕 **BULK image upload for all templates**
            """)
            
        with col2:
            st.markdown("""
            **Image Processing:**
            - 🖼️ Auto image extraction from Excel data files
            - 📍 Smart image placement in templates
            - 🎨 Format conversion and optimization
            - 📦 Packaging image area detection
            - 🎯 **Images filtered by part number/description**
            - 🆕 **BULK: Upload images that apply to all templates**
            - 🔄 **HYBRID: Combine bulk + extracted images**
            """)
        
        st.markdown("""
        ### 🖼️ Enhanced Image Processing Modes
        - **AUTO-EXTRACTION**: Images are automatically extracted from Excel data files
        - **BULK UPLOAD**: Upload images that will be applied to ALL templates
        - **HYBRID MODE (NEW)**: Combine both bulk uploaded AND extracted images
        - **Smart Filtering**: Extracted images are filtered by part number/description for each row
        - **Non-Override**: Bulk images complement extracted images instead of replacing them
        - **Fixed Positions**: Images are placed at predefined positions in templates
        - Supports multiple image formats (PNG, JPG, GIF, BMP)
        
        ### 🔄 HYBRID Mode Benefits
        - **Best of Both**: Use consistent bulk images + specific extracted images
        - **Flexible Workflow**: Upload common images once, extract specific ones per row
        - **No Data Loss**: Both image sources are preserved and combined
        - **Smart Combination**: Each template gets both bulk and row-specific images
        """)


def filter_images_for_row(extracted_images, row, columns):
    """
    Filter extracted images to only include those that match the current row's 
    part number and description.
    
    Args:
        extracted_images: Dictionary of all extracted images
        row: Current data row (pandas Series)
        columns: List of column names from the dataframe
        
    Returns:
        Dictionary of filtered images for this specific row
    """
    if not extracted_images or 'all_sheets' not in extracted_images:
        return {}
    
    try:
        # Get part number and description from current row
        part_no = get_field_value(row, columns, ['part_no', 'partno', 'part_number', 'partnumber', 'part no', 'part number'])
        part_desc = get_field_value(row, columns, ['part_description', 'partdescription', 'description', 'part_desc', 'partdesc', 'part description', 'part desc'])
        
        if not part_no and not part_desc:
            print("⚠️ No part number or description found for filtering images")
            return extracted_images  # Return all images if we can't identify the row
        
        print(f"🎯 Filtering images for: Part No='{part_no}', Description='{part_desc}'")
        
        filtered_images = {}
        all_images = extracted_images['all_sheets']
        
        # Create search terms for matching
        search_terms = []
        if part_no:
            search_terms.append(str(part_no).lower().strip())
        if part_desc:
            search_terms.append(str(part_desc).lower().strip())
        
        # Check each image to see if it matches this row
        for img_key, img_data in all_images.items():
            should_include = False
            
            # Method 1: Check if image is from a sheet that matches the part info
            sheet_name = img_data.get('sheet', '').lower()
            position = img_data.get('position', '').lower()
            
            # Look for part number or description in sheet name or position
            for term in search_terms:
                if term and (term in sheet_name or term in position):
                    should_include = True
                    print(f"✅ Including image {img_key} - found '{term}' in sheet/position")
                    break
            
            # Method 2: If we don't have specific matching, include all images from the first sheet
            # (This is a fallback when images aren't clearly labeled)
            if not should_include and not any(search_terms):
                should_include = True
                print(f"✅ Including image {img_key} - fallback (no specific identifiers)")
            
            # Method 3: If this is the only row or images aren't clearly separated, include all
            if not should_include and len(search_terms) == 0:
                should_include = True
                print(f"✅ Including image {img_key} - no filtering criteria")
            
            if should_include:
                filtered_images[img_key] = img_data
        
        print(f"🎯 Filtered {len(filtered_images)} images from {len(all_images)} total images")
        return {'all_sheets': filtered_images}
        
    except Exception as e:
        print(f"❌ Error filtering images for row: {e}")
        return extracted_images  # Return all images on error


def get_field_value(row, columns, field_names):
    """
    Get value from row using multiple possible field names.
    
    Args:
        row: pandas Series (data row)
        columns: List of column names
        field_names: List of possible field names to search for
        
    Returns:
        Field value or None if not found
    """
    try:
        # Normalize field names for comparison
        normalized_field_names = [name.lower().replace(' ', '').replace('_', '') for name in field_names]
        
        # Check each column in the dataframe
        for col in columns:
            normalized_col = col.lower().replace(' ', '').replace('_', '')
            
            # Check if this column matches any of our target field names
            for target_field in normalized_field_names:
                if target_field in normalized_col or normalized_col in target_field:
                    value = row.get(col)
                    if pd.notna(value) and str(value).strip():
                        return str(value).strip()
        
        return None
        
    except Exception as e:
        print(f"Error getting field value: {e}")
        return None


def generate_enhanced_filename(row, columns, index):
    """
    Generate an enhanced filename based on row data.
    
    Args:
        row: pandas Series (data row)
        columns: List of column names
        index: Row index
        
    Returns:
        Enhanced filename string
    """
    try:
        # Get key identifiers from the row
        vendor_code = get_field_value(row, columns, [
            'vendor_code', 'vendorcode', 'vendor', 'supplier_code', 'supplier'
        ])
        
        part_no = get_field_value(row, columns, [
            'part_no', 'partno', 'part_number', 'partnumber', 'part no', 'part number'
        ])
        
        part_desc = get_field_value(row, columns, [
            'part_description', 'partdescription', 'description', 'part_desc', 'partdesc'
        ])
        
        # Clean up values for filename (remove invalid characters)
        def clean_for_filename(value, max_length=30):
            if not value:
                return ""
            # Remove invalid characters and limit length
            cleaned = re.sub(r'[<>:"/\\|?*]', '_', str(value))
            cleaned = re.sub(r'\s+', '_', cleaned)  # Replace spaces with underscores
            return cleaned[:max_length]
        
        # Build filename components
        filename_parts = []
        
        if vendor_code:
            filename_parts.append(clean_for_filename(vendor_code, 20))
        
        if part_no:
            filename_parts.append(clean_for_filename(part_no, 25))
        
        if part_desc:
            filename_parts.append(clean_for_filename(part_desc, 40))
        
        # If we don't have enough info, use row index
        if not filename_parts:
            filename_parts.append(f"row_{index + 1}")
        
        # Join parts and add template suffix
        filename = "_".join(filename_parts) + "_template.xlsx"
        
        # Ensure filename isn't too long (Windows has 255 char limit)
        if len(filename) > 200:
            filename = filename[:190] + "_template.xlsx"
        
        return filename
        
    except Exception as e:
        print(f"Error generating filename: {e}")
        return f"template_row_{index + 1}.xlsx"

def main():
    try:
        if not st.session_state.authenticated:
            show_login()
        else:
            show_main_app()
    except Exception as e:
        st.error(f"Application Error: {e}")
        st.exception(e)

if __name__ == "__main__":
    main()
