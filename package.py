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
    """Handles image extraction from Excel files with improved duplicate handling"""
    
    def __init__(self):
        self.supported_formats = ['.png', '.jpg', '.jpeg', '.gif', '.bmp']
        self._placement_counters = defaultdict(int)
    
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
        """Extract unique images from Excel file with better type classification"""
        try:
            images = {}
            workbook = openpyxl.load_workbook(excel_file_path)
            image_hashes = set()
            
            print("=== Extracting images from Excel ===")
            
            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
                print(f"Processing sheet: {sheet_name}")

                if hasattr(worksheet, '_images') and worksheet._images:
                    print(f"Found {len(worksheet._images)} images in {sheet_name}")

                    for idx, img in enumerate(worksheet._images):
                        try:
                            # Get image data
                            image_data = img._data()

                            # Create hash of image data to detect duplicates
                            image_hash = hashlib.md5(image_data).hexdigest()

                            if image_hash in image_hashes:
                                print(f"Skipping duplicate image in {sheet_name}")
                                continue
                            image_hashes.add(image_hash)

                            # Create PIL Image
                            pil_image = Image.open(io.BytesIO(image_data))

                            # Get image position
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

                            # Improved image type classification
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

                            print(f"Extracted image: {image_key} (type: {image_type}) at position {position}")
                            
                        except Exception as e:
                            print(f"Error extracting image {idx} from {sheet_name}: {e}")
                            continue
                else:
                    print(f"No images found in sheet: {sheet_name}")
                    
            workbook.close()
            print(f"Total unique extracted images: {len(images)}")
            return {'all_sheets': images}
            
        except Exception as e:
            st.error(f"Error extracting images: {e}")
            print(f"Error in extract_images_from_excel: {e}")
            return {}

    def _classify_image_type(self, sheet_name, position, index):
        """Classify image type based on position and index, not sheet name"""
        # Ignore sheet name completely - classify based on image order/position
        # First image = current, then cycle through primary, secondary, label
        
        print(f"Classifying image {index} from sheet '{sheet_name}' at position '{position}'")
        
        # Simple classification based on image index only
        if index == 0:
            image_type = 'current'  # First image is always current packaging
        elif index == 1:
            image_type = 'primary'  # Second image is primary
        elif index == 2:
            image_type = 'secondary'  # Third image is secondary
        elif index == 3:
            image_type = 'label'  # Fourth image is label
        else:
            # For additional images, cycle through types
            type_cycle = ['primary', 'secondary', 'label']
            image_type = type_cycle[(index - 1) % len(type_cycle)]
        
        print(f"-> Classified as: {image_type}")
        return image_type

    def add_images_to_template(self, worksheet, uploaded_images, image_areas):
        """Add uploaded images to template with fixed positions"""
        try:
            added_images = 0
            temp_image_paths = []
            used_images = set()
            
            print("=== Adding images to template ===")
            print(f"Available images: {len(uploaded_images)}")
            
            # Initialize counter for row 42 images (primary, secondary, label)
            row_42_counter = 0
            
            # Process images in order: current, primary, secondary, label
            for image_type in ['current', 'primary', 'secondary', 'label']:
                type_images = {
                    k: v for k, v in uploaded_images.items()
                    if v.get('type', '').lower() == image_type and k not in used_images
                }
                if not type_images:
                    continue
                    
                print(f"\n--- Processing {image_type} images ---")
                print(f"Found {len(type_images)} images of type '{image_type}'")
                
                for idx, (img_key, img_data) in enumerate(type_images.items()):
                    print(f"Processing image {idx + 1}/{len(type_images)}: {img_key}")
                    
                    added_images += self._place_single_image(
                        worksheet, img_key, img_data, image_type, idx, row_42_counter,
                        temp_image_paths, used_images
                    )
                    
                    # Increment row 42 counter for non-current images
                    if image_type != 'current':
                        row_42_counter += 1
                    
            print(f"\n✅ Total images added: {added_images}")
            return added_images, temp_image_paths
            
        except Exception as e:
            st.error(f"Error adding images to template: {e}")
            print(f"Error in add_images_to_template: {e}")
            return 0, []
                    
    def _place_single_image(self, worksheet, img_key, img_data, image_type, image_index, row_42_counter, temp_image_paths, used_images):
        """Place image at fixed positions: current at row 2 col 20, others at row 42 with spacing"""
        if not hasattr(self, '_global_image_counter'):
            self._global_image_counter = 0
        try:
            # Create temporary image file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                image_bytes = base64.b64decode(img_data['data'])
                tmp_img.write(image_bytes)
                tmp_img_path = tmp_img.name
            # Create openpyxl image object
            img = OpenpyxlImage(tmp_img_path)
        
            # Initialize cell_coord variable
            cell_coord = None
        
            # FIXED PLACEMENT LOGIC - CORRECTED POSITIONING
            if image_type == 'current':
                # CURRENT PACKAGING: Always at row 3, column 20 (T3)
                target_row = 3
                target_col = 20
                # Current images are larger (8.3cm x 8.3cm)
                img.width = int(8.3 * 37.8)
                img.height = int(8.3 * 37.8)
                cell_coord = f"{get_column_letter(target_col)}{target_row}"
                print(f"🎯 CURRENT IMAGE: Placing at row={target_row}, col={target_col} (8.3x8.3cm) -> {cell_coord}")
            else:
                # 🟢 Sequential horizontal placement for other images on row 42 with proper spacing
                target_row = 42  # Changed from 41 to 42 as per your requirement
            
                # Use proper spacing calculations for row 42
                image_width_cols = int(4.3 * 1.162)  # ≈ 5 columns for regular images
                gap_cols = int(1.162 * 1.162)         # ≈ 3 columns gap
                total_spacing = image_width_cols + gap_cols
            
                # Start at column 1 (A), then shift right for each non-current image
                target_col = 1 + (self._global_image_counter * total_spacing)
            
                # Increment counter for next non-current image
                self._global_image_counter += 1
            
                # Other images are smaller (4.3cm x 4.3cm)
                img.width = int(4.3 * 37.8)
                img.height = int(4.3 * 37.8)
            
                cell_coord = f"{get_column_letter(target_col)}{target_row}"
                print(f"📍 {image_type.upper()} IMAGE: Placing at row 42, position: {cell_coord}")
                print(f"   Image key: {img_key}")
                print(f"   Image type: {image_type}")
                print(f"   Global counter: {self._global_image_counter}")
                print(f"   Spacing calculation: width_cols={image_width_cols}, gap_cols={gap_cols}, total={total_spacing}")
            
            # Ensure cell_coord was set
            if cell_coord is None:
                raise ValueError(f"Could not determine cell coordinate for image type: {image_type}")
            
            # Set image position and add to worksheet
            img.anchor = cell_coord
            worksheet.add_image(img)
        
            # Track temporary files and used images
            temp_image_paths.append(tmp_img_path)
            used_images.add(img_key)
            print(f"✅ Successfully added {image_type} image '{img_key}' at {cell_coord}")
            return 1
            
        except Exception as e:
            print(f"❌ Could not add image {img_key}: {e}")
            st.warning(f"Could not add image {img_key}: {e}")
            return 0

    def _create_additional_placement_area(self, area_type, index, existing_areas):
        """Create additional placement area when no predefined area exists"""
        # Define column positions for different image types
        type_columns = {
            'primary': 2,    # Column B
            'secondary': 6,  # Column F
            'current': 3,    # Column C (special case)
            'label': 11      # Column K
        }
        
        if area_type == 'current':
            # Current images should go to a specific location
            target_column = 3  # Column C
            target_row = 6     # Row 6
        else:
            # Other images go to row 42 with spacing (changed from 41)
            target_column = type_columns.get(area_type, 2)
            target_row = 42 + (index * 12)  # Vertical spacing for multiple images of same type
        
        # Create a virtual area for additional placement
        return {
            'position': f"{get_column_letter(target_column)}{target_row}",
            'row': target_row,
            'column': target_column,
            'text': f"Additional {area_type}",
            'type': area_type,
            'header_text': area_type,
            'matched_keyword': area_type,
            'match_score': 1
        }

    def _place_remaining_images(self, worksheet, remaining_images, image_type, temp_image_paths, used_images):
        """Place remaining images in available columns with proper spacing"""
        try:
            # Define column positions for different image types
            type_columns = {
                'primary': 2,    # Column B
                'secondary': 6,  # Column F
                'current': 3,    # Column C (special case)
                'label': 11      # Column K
            }
            
            if image_type == 'current':
                # Current images should go to specific location
                target_col = 3  # Column C
                start_row = 6   # Row 6
            else:
                target_col = type_columns.get(image_type, 2)
                start_row = 42  # Changed from 41 to 42
            
            for idx, (img_key, img_data) in enumerate(list(remaining_images.items())):
                if image_type == 'current':
                    target_row = start_row + (idx * 10)  # Vertical spacing for multiple current images
                else:
                    target_row = start_row + (idx * 12)  # Vertical spacing for other types
                
                try:
                    # Create and place image
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                        image_bytes = base64.b64decode(img_data['data'])
                        tmp_img.write(image_bytes)
                        tmp_img_path = tmp_img.name
                    
                    img = OpenpyxlImage(tmp_img_path)
                    
                    # Set size based on image type
                    if image_type == 'current':
                        img.width = int(8.3 * 37.8)   # 8.3cm width for current
                        img.height = int(8.3 * 37.8)  # 8.3cm height for current
                    else:
                        img.width = int(4.3 * 37.8)   # 4.3cm width for others
                        img.height = int(4.3 * 37.8)  # 4.3cm height for others
                    
                    cell_coord = f"{get_column_letter(target_col)}{target_row}"
                    img.anchor = cell_coord
                    worksheet.add_image(img)
                    
                    temp_image_paths.append(tmp_img_path)
                    used_images.add(img_key)
                    
                    print(f"✅ Placed remaining {image_type} image at {cell_coord}")
                    
                except Exception as e:
                    print(f"❌ Error placing remaining image: {e}")
                    continue
                    
        except Exception as e:
            print(f"Error in _place_remaining_images: {e}")

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
                    'part information', 'part', 'component', 'item'
                ],
                'field_mappings': {
                    'l': 'Part L',
                    'length': 'Part L',
                    'w': 'Part W',
                    'width': 'Part W',
                    'h': 'Part H',
                    'height': 'Part H',
                    'part no': 'Part No',
                    'part number': 'Part No',
                    'description': 'Part Description',
                    'unit weight': 'Part Unit Weight'
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
            for search_row in range(max(1, row - max_search_rows), row + 1):
                for search_col in range(max(1, col - 10), min(worksheet.max_column + 1, col + 10)):
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
                                    
                                    # Check for key words within the text
                                    if section_name == 'primary_packaging' and ('primary' in cell_text and 'packaging' in cell_text):
                                        return section_name
                                    elif section_name == 'secondary_packaging' and ('secondary' in cell_text and 'packaging' in cell_text):
                                        return section_name
                                    elif section_name == 'part_information' and ('part' in cell_text and 'information' in cell_text):
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
            
            # Define mappable field patterns for packaging templates
            mappable_patterns = [
                r'l[-\s]*mm', r'w[-\s]*mm', r'h[-\s]*mm',  # Dimension fields
                r'l\b', r'w\b', r'h\b',  # Single letter dimensions
                r'packaging\s+type', r'qty[/\s]*pack',      # Packaging fields
                r'part\s+[lwh]', r'component\s+[lwh]',      # Part dimension fields
                r'length', r'width', r'height',             # Basic dimensions
                r'quantity', r'pack\s+weight', r'total',    # Quantity fields
                r'empty\s+weight', r'weight', r'unit\s+weight',  # Weight fields
                r'code', r'name', r'description',           # Basic info fields
                r'vendor', r'supplier', r'customer',        # Entity fields
                r'date', r'revision', r'reference',         # Document fields
                r'part\s+no', r'part\s+number'              # Part identification
            ]
            
            for pattern in mappable_patterns:
                if re.search(pattern, text):
                    return True
            
            # Check if it ends with colon (label pattern)
            if text.endswith(':'):
                return True
                
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
        
    def debug_worksheet_procedure_area(self, worksheet):
        """Debug function to inspect the procedure area of the worksheet"""
        try:
            print(f"\n=== DEBUGGING PROCEDURE AREA (Rows 25-40, Cols A-Q) ===")
        
            # Check for merged cells in the procedure area
            procedure_merged_cells = []
            for merged_range in worksheet.merged_cells.ranges:
                if (merged_range.min_row <= 40 and merged_range.max_row >= 25 and merged_range.min_col <= 17 and merged_range.max_col >= 1):
                    procedure_merged_cells.append(merged_range)
        
            print(f"Merged cells in procedure area: {len(procedure_merged_cells)}")
            for merged_range in procedure_merged_cells:
                print(f"  - {merged_range}")
        
            # Check specific procedure area (rows 28-38, columns B and P)
            print(f"\nProcedure step target area inspection:")
            for row in range(28, 39):  # rows 28-38
                for col in [2, 16]:  # columns B(2) and P(16)
                    cell = worksheet.cell(row=row, column=col)
                    col_letter = 'B' if col == 2 else 'P'
                
                    # Check if cell is in a merged range
                    is_merged = False
                    merged_info = ""
                    for merged_range in worksheet.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            is_merged = True
                            merged_info = f" [MERGED: {merged_range}]"
                            break
                
                    cell_value = cell.value if cell.value else "<empty>"
                    print(f"  {col_letter}{row}: {cell_value}{merged_info}")
        
            # Check if cells are protected
            print(f"\nWorksheet protection status:")
            print(f"  Sheet protected: {worksheet.protection.sheet}")
            print(f"  Password: {worksheet.protection.password}")
        
            # Check specific cells for protection
            sample_cells = [worksheet.cell(row=28, column=2), worksheet.cell(row=28, column=16)]
            for cell in sample_cells:
                if hasattr(cell, 'protection'):
                    print(f"  {cell.coordinate} locked: {cell.protection.locked}")
        
            print("=== END PROCEDURE AREA DEBUG ===\n")
        
        except Exception as e:
            print(f"Error in debug_worksheet_procedure_area: {e}")
            import traceback
            traceback.print_exc()
    
    def write_procedure_steps_to_template(self, worksheet, packaging_type, data_dict=None):
        """Write packaging procedure steps in Column B starting from Row 28 (Step numbers already exist in Column A)"""
        try:
            from openpyxl.cell import MergedCell
            from openpyxl.styles import Font, Alignment
            import traceback

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

                # Check for merged cell and unmerge if necessary
                for merged_range in worksheet.merged_cells.ranges:
                    if target_cell.coordinate in merged_range:
                        if not (step_row == merged_range.min_row and target_col == merged_range.min_col):
                            print(f"🔧 Unmerging range {merged_range}")
                            worksheet.unmerge_cells(str(merged_range))
                            target_cell = worksheet.cell(row=step_row, column=target_col)
                        break

                # Write step content only (no step number)
                target_cell.value = step_text

                # Apply formatting
                target_cell.font = Font(name='Calibri', size=10)
                target_cell.alignment = Alignment(wrap_text=True, vertical='top')

                steps_written += 1

            print(f"\n✅ PROCEDURE STEPS COMPLETED")
            print(f"   Total steps written: {steps_written}")
            print(f"   Location: Column B, starting from Row 28")

            return steps_written

        except Exception as e:
            print(f"💥 Critical error in write_procedure_steps_to_template: {e}")
            import traceback
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
        
        try:
            data_columns = data_df.columns.tolist()
            
            for coord, field in template_fields.items():
                try:
                    best_match = None
                    best_score = 0.0
                    field_value = field['value'].lower().strip()
                    section_context = field.get('section_context')
                    
                    # Try section-based mapping first
                    if section_context and section_context in self.section_mappings:
                        section_mappings = self.section_mappings[section_context]['field_mappings']
                        
                        # Look for direct field matches within section
                        for template_field_key, data_column_pattern in section_mappings.items():
                            if template_field_key in field_value or field_value in template_field_key:
                                # Look for exact match first
                                for data_col in data_columns:
                                    if data_column_pattern.lower() == data_col.lower():
                                        best_match = data_col
                                        best_score = 1.0
                                        break
                                
                                # If no exact match, try similarity matching
                                if not best_match:
                                    for data_col in data_columns:
                                        similarity = self.calculate_similarity(data_column_pattern, data_col)
                                        if similarity > best_score and similarity >= self.similarity_threshold:
                                            best_score = similarity
                                            best_match = data_col
                                break
                    
                    # Fallback to general similarity matching
                    if not best_match:
                        for data_col in data_columns:
                            similarity = self.calculate_similarity(field_value, data_col)
                            if similarity > best_score and similarity >= self.similarity_threshold:
                                best_score = similarity
                                best_match = data_col
                    
                    mapping_results[coord] = {
                        'template_field': field['value'],
                        'data_column': best_match,
                        'similarity': best_score,
                        'field_info': field,
                        'section_context': section_context,
                        'is_mappable': best_match is not None
                    }
                        
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
        """Add uploaded images to template in designated areas"""
        try:
            added_images = 0
            temp_image_paths = []
            used_images = set()
            for area in image_areas:
                area_type = area['type']
                label_text = area.get('text', '').lower()  # ✅ Define label_text here

                matching_image = None

                for label, img_data in uploaded_images.items():
                    if label in used_images:
                        continue
                    label_lower = label.lower()
                    if (
                        area_type in label_lower
                        or area_type.replace('_', ' ') in label_lower
                        or 'primary' in label_lower and area_type == 'primary_packaging'
                        or 'secondary' in label_lower and area_type == 'secondary_packaging'
                        or 'current' in label_lower and area_type == 'current_packaging'
                        or label_lower in label_text
                        or label_text in label_lower
                    ):
                        matching_image = img_data
                        used_images.add(label)
                        break
                        
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
    
            print(f"Starting template fill process...")
            print(f"Data dictionary keys: {list(data_dict.keys())}")
            print(f"Packaging type: {packaging_type}")
    
            # Fill data fields
            print("=== Filling data fields ===")
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
                            print(f"Filled {target_cell} with: {data_value}")
                    
                except Exception as e:
                    print(f"Error filling mapping {coord}: {e}")
                    continue
    
            print(f"=== Data fields filled: {filled_count} ===")
    
            # Add images if provided
            if uploaded_images:
                print("=== Adding images ===")
                try:
                    # First, identify image upload areas
                    _, image_areas = self.find_template_fields_with_context_and_images(template_file)
                    images_added, temp_image_paths = self.image_extractor.add_images_to_template(worksheet, uploaded_images, image_areas)
                    print(f"Images added: {images_added}")
                except Exception as e:
                    print(f"Error adding images: {e}")
                    images_added = 0
    
            # Write procedure steps if packaging type is provided
            if packaging_type and packaging_type != "Select Packaging Procedure":
                print(f"=== Writing procedure steps for {packaging_type} ===")
                try:
                    # Debug worksheet area before writing procedure steps
                    self.debug_worksheet_procedure_area(worksheet)
                
                    # Write procedure steps
                    procedure_steps_added = self.write_procedure_steps_to_template(worksheet, packaging_type, data_dict)
                    print(f"✅ Successfully added {procedure_steps_added} procedure steps")
                
                    # Debug worksheet area after writing procedure steps
                    print("After writing procedure steps:")
                    self.debug_worksheet_procedure_area(worksheet)
                
                except Exception as e:
                    print(f"❌ Error adding procedure steps: {e}")
                    import traceback
                    traceback.print_exc()
                    procedure_steps_added = 0
            else:
                print("No packaging type provided or invalid packaging type")
        
            print(f"=== Template fill completed ===")
            print(f"Final results - Data fields: {filled_count}, Images: {images_added}, Procedure steps: {procedure_steps_added}")
        
            return workbook, filled_count, images_added, temp_image_paths, procedure_steps_added
    
        except Exception as e:
            print(f"Critical error in fill_template_with_data_and_images: {e}")
            import traceback
            traceback.print_exc()
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
    
    # Main content area
    if template_file and data_file:
        try:
            # Extract images from data file (only if it's Excel)
            extracted_images = {}
            if data_file.name.endswith(('.xlsx', '.xls')):
                # Create temporary file for data file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_data:
                    tmp_data.write(data_file.getvalue())
                    data_path = tmp_data.name
                
                st.info("🔍 Extracting images from data file...")
                with st.spinner("Extracting images from Excel file..."):
                    extracted_images = st.session_state.enhanced_mapper.image_extractor.extract_images_from_excel(data_path)
                
                # Clean up data file copy
                try:
                    os.unlink(data_path)
                except:
                    pass
            
            # Read data file
            if data_file.name.endswith('.csv'):
                data_df = pd.read_csv(data_file)
            else:
                data_df = pd.read_excel(data_file)
            
            # Create temporary file for template
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_template:
                tmp_template.write(template_file.getvalue())
                template_path = tmp_template.name
            
            # Process template and find fields
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
                
                # Show extracted images from data file
                if extracted_images:
                    total_images = sum(len(sheet_images) for sheet_images in extracted_images.values())
                    st.success(f"🖼️ Extracted {total_images} images from data file")
                    
                    with st.expander("Extracted Images from Data File", expanded=True):
                        for sheet_name, sheet_images in extracted_images.items():
                            if sheet_images:
                                st.write(f"**Sheet: {sheet_name}**")
                                cols = st.columns(min(3, len(sheet_images)))
                                
                                for idx, (position, img_data) in enumerate(sheet_images.items()):
                                    with cols[idx % 3]:
                                        st.write(f"Position: {position}")
                                        # Display image thumbnail
                                        img_bytes = base64.b64decode(img_data['data'])
                                        st.image(img_bytes, width=150)
                                        st.write(f"Size: {img_data['size']}")
                                        st.write(f"Type: {img_data.get('type', 'Unknown')}")
                else:
                    if data_file.name.endswith(('.xlsx', '.xls')):
                        st.info("No images found in the data file")
                    else:
                        st.info("CSV files don't contain images. Use Excel files to include images.")
                
                # Data mapping
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
                            
                            # Get procedure steps with data substitution
                            try:
                                data_dict = data_df.iloc[0].to_dict() if len(data_df) > 0 else {}
                                procedures = st.session_state.enhanced_mapper.get_procedure_steps(procedure_type, data_dict)
                                
                                st.write("**Procedure Steps Preview:**")
                                
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
                    
                    # Only inject steps into data_df if not preview only
                    if procedure_type and procedure_type != "Select Packaging Procedure" and not preview_only:
                        try:
                            # Get the first row of data for procedure step generation
                            data_dict = data_df.iloc[0].to_dict() if len(data_df) > 0 else {}
                            procedure_steps = st.session_state.enhanced_mapper.get_procedure_steps(procedure_type, data_dict)
                            
                            # Add procedure steps to the dataframe
                            for i, step in enumerate(procedure_steps, 1):
                                if step.strip():  # Only add non-empty steps
                                    data_df.loc[0, f'Procedure Step {i}'] = step
                            
                            # Add packaging type to data
                            data_df.loc[0, 'Primary Packaging Type'] = procedure_type
                            
                            st.success(f"✅ Added {len([s for s in procedure_steps if s.strip()])} packaging procedure steps to template data")
                            
                        except Exception as e:
                            st.error(f"Error adding procedure steps to data: {e}")
		
                    # Fill template
                    st.subheader("📝 Generate Filled Template")
                    
                    # Show what will be included in the template
                    st.write("**Template will include:**")
                    include_items = []
                    
                    # Count mapped fields
                    mapped_count = sum(1 for mapping in mapping_results.values() if mapping['is_mappable'])
                    if mapped_count > 0:
                        include_items.append(f"📊 {mapped_count} mapped data fields")
                    
                    # Count images
                    if extracted_images:
                        total_images = sum(len(sheet_images) for sheet_images in extracted_images.values())
                        if total_images > 0:
                            include_items.append(f"🖼️ {total_images} extracted images")
                    
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
                        st.warning("No items will be added to the template")
                    
                    if st.button("Generate Filled Template", type="primary", use_container_width=True):
                        with st.spinner("Filling template with data, images, and procedure steps..."):
                            try:
                                # Convert extracted images to the format expected by fill_template_with_data_and_images
                                processed_images = {}
                                if extracted_images:
                                    for sheet_name, sheet_images in extracted_images.items():
                                        for position, img_data in sheet_images.items():
                                            # Create a unique key for each image
                                            image_key = f"{sheet_name}_{position}"
                                            processed_images[image_key] = img_data
                                
                                # Pass the packaging type to the fill function
                                selected_packaging_type = procedure_type if (procedure_type and procedure_type != "Select Packaging Procedure" and not preview_only) else None
                                
                                # Updated function call with procedure steps
                                result = st.session_state.enhanced_mapper.fill_template_with_data_and_images(
                                    template_path, mapping_results, data_df, processed_images, selected_packaging_type
                                )
                                
                                workbook, filled_count, images_added, temp_image_paths, procedure_steps_added = result
                                
                                if workbook:
                                    # Save filled template
                                    output_buffer = io.BytesIO()
                                    workbook.save(output_buffer)
                                    
                                    # Clean up temporary image files
                                    for path in temp_image_paths:
                                        try:
                                            os.unlink(path)
                                        except Exception as e:
                                            st.warning(f"Failed to delete temp file {path}: {e}")
                                    
                                    output_buffer.seek(0)
                                    
                                    # Enhanced success message
                                    st.success("🎉 Template filled successfully!")
                                    
                                    # Show detailed completion stats
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("Data Fields", filled_count)
                                    with col2:
                                        st.metric("Images Added", images_added)
                                    with col3:
                                        st.metric("Procedure Steps", procedure_steps_added)
                                    
                                    # Provide additional feedback
                                    if images_added > 0:
                                        st.info(f"🖼️ Successfully placed {images_added} images from data file")
                                    elif processed_images:
                                        st.warning("⚠️ Images were found but could not be placed in template areas")
                                    
                                    if procedure_steps_added > 0:
                                        st.info(f"📋 Added {procedure_steps_added} packaging procedure steps")
                                    
                                    # Download button
                                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                    filename = f"filled_template_{timestamp}.xlsx"
                                    
                                    st.download_button(
                                        label="📥 Download Filled Template",
                                        data=output_buffer.getvalue(),
                                        file_name=filename,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True
                                    )
                                    
                                    workbook.close()
                                else:
                                    st.error("❌ Failed to fill template")
                                    
                            except Exception as e:
                                st.error(f"Error filling template: {e}")
                                st.exception(e)
                
                else:
                    st.warning("No mapping results generated")
            
            else:
                st.warning("No mappable fields found in template")
            
            # Clean up temporary file
            try:
                os.unlink(template_path)
            except:
                pass
                
        except Exception as e:
            st.error(f"Error processing files: {e}")
            st.exception(e)
    
    else:
        st.info("👆 Please upload both an Excel template and a data file to begin")
        
        # Show demo information
        st.markdown("### 🎯 Features")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **Template Processing:**
            - 📋 Smart field detection
            - 🎯 Section-aware mapping
            - 🔄 Merged cell handling
            - 📏 Packaging-specific patterns
            """)
            
        with col2:
            st.markdown("""
            **Image Processing:**
            - 🖼️ Auto image extraction from Excel data files
            - 📍 Smart image placement in templates
            - 🎨 Format conversion and optimization
            - 📦 Packaging image area detection
            """)
        
        st.markdown("""
        ### 📚 Supported Sections
        - **Primary Packaging**: Internal packaging dimensions and specifications
        - **Secondary Packaging**: Outer packaging details
        - **Part Information**: Component specifications and measurements
        
        ### 🖼️ Image Processing
        - Images are automatically extracted from Excel data files
        - Supports multiple image formats (PNG, JPG, GIF, BMP)
        - Images are intelligently placed in designated template areas
        - No manual image upload required - everything is automated!
        
        ### 📋 Packaging Procedures
        - **11+ Predefined Procedures**: Complete packaging workflows for different product types
        - **Smart Substitution**: Automatically replaces placeholders with actual data values
        - **Preview Mode**: Review steps before adding to template
        - **Integrated Workflow**: Seamlessly adds procedure steps to your filled templates
        """)

# Main application logic
def main():
    if not st.session_state.authenticated:
        show_login()
    else:
        show_main_app()

if __name__ == "__main__":
    main()
