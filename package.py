import streamlit as st
import pandas as pd
import numpy as np
import os
import json
import hashlib
import tempfile
import shutil
from pathlib import Path
from collections import defaultdict
import zipfile
from PIL import Image
import base64
import traceback
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.drawing.image import Image as OpenpyxlImage
import re
from datetime import datetime
from difflib import SequenceMatcher

def navigate_to_step(step_number):
    """Helper function to navigate between steps"""
    if 1 <= step_number <= 6:
        st.session_state.current_step = step_number
        st.rerun()

# Configure Streamlit page
st.set_page_config(
    page_title="AI Packaging Template Mapper",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)
if 'mapping_completed' not in st.session_state:
    st.session_state.mapping_completed = False
if 'auto_fill_started' not in st.session_state:
    st.session_state.auto_fill_started = False
# Initialize session state
if 'current_step' not in st.session_state:
    st.session_state.current_step = 1
if 'selected_packaging_type' not in st.session_state:
    st.session_state.selected_packaging_type = ''
if 'template_file' not in st.session_state:
    st.session_state.template_file = None
if 'data_file' not in st.session_state:
    st.session_state.data_file = None
if 'mapped_data' not in st.session_state:
    st.session_state.mapped_data = None
if 'image_option' not in st.session_state:
    st.session_state.image_option = ''
if 'uploaded_images' not in st.session_state:
    st.session_state.uploaded_images = {}
if 'extracted_excel_images' not in st.session_state:
    st.session_state.extracted_excel_images = {}

class ImageExtractor:
    """Handles image extraction from Excel files with improved duplicate handling"""
    
    def __init__(self):
        self.supported_formats = ['.png', '.jpg', '.jpeg', '.gif', '.bmp']
        self._placement_counters = defaultdict(int)
        self.current_excel_path = None
    
    def extract_images_from_excel(self, excel_file_path):
        """Extract images from Excel file using multiple methods"""
        try:
            self.current_excel_path = excel_file_path
            images = {}
            
            st.write("🔍 Extracting images from Excel file...")
            
            # METHOD 1: Standard openpyxl extraction
            try:
                result1 = self._extract_with_openpyxl(excel_file_path)
                images.update(result1)
                st.write(f"✅ Standard extraction found {len(result1)} images")
            except Exception as e:
                st.write(f"⚠️ Standard extraction failed: {e}")
            
            # METHOD 2: ZIP-based extraction (Excel files are ZIP archives)
            if not images:
                try:
                    result2 = self._extract_with_zipfile(excel_file_path)
                    images.update(result2)
                    st.write(f"✅ ZIP extraction found {len(result2)} images")
                except Exception as e:
                    st.write(f"⚠️ ZIP extraction failed: {e}")
            
            if not images:
                st.warning("⚠️ No images found in Excel file. Please ensure images are embedded in the Excel file.")
            else:
                st.success(f"🎯 Total images extracted: {len(images)}")
            
            return {'all_sheets': images}
            
        except Exception as e:
            st.error(f"❌ Error extracting images: {e}")
            return {}

    def _extract_with_openpyxl(self, excel_file_path):
        """Standard openpyxl image extraction"""
        images = {}
        
        try:
            workbook = openpyxl.load_workbook(excel_file_path, data_only=False)
            
            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
                
                if hasattr(worksheet, '_images') and worksheet._images:
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
                            
                            # Classify image type based on position
                            image_type = self._classify_image_type(idx)
                            
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
                            
                        except Exception as e:
                            st.write(f"❌ Failed to extract image {idx} from sheet {sheet_name}: {e}")
            
            workbook.close()
            
        except Exception as e:
            st.error(f"❌ Error in openpyxl extraction: {e}")
            raise
        
        return images

    def _extract_with_zipfile(self, excel_file_path):
        """Extract images by treating Excel file as ZIP archive"""
        images = {}
        
        try:
            with zipfile.ZipFile(excel_file_path, 'r') as zip_ref:
                # List all files in the archive
                file_list = zip_ref.namelist()
                
                # Look for media files
                media_files = [f for f in file_list if '/media/' in f.lower()]
                image_files = [f for f in file_list if any(f.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp'])]
                
                # Extract images from media folder
                for idx, media_file in enumerate(media_files):
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
                            image_type = self._classify_image_type(idx)
                            image_key = f"{image_type}_{filename}_{idx}"
                            
                            images[image_key] = {
                                'data': img_str,
                                'format': 'PNG',
                                'size': pil_image.size,
                                'position': f"ZIP_{idx}",
                                'sheet': 'ZIP_EXTRACTED',
                                'index': idx,
                                'type': image_type,
                                'hash': image_hash,
                                'source_path': media_file
                            }
                            
                    except Exception as e:
                        st.write(f"❌ Failed to extract {media_file}: {e}")
        
        except Exception as e:
            st.error(f"❌ Error in ZIP extraction: {e}")
            raise
        
        return images

    def _classify_image_type(self, index):
        """Classify image type based on index"""
        types = ['current', 'primary', 'secondary', 'label']
        return types[index % len(types)]

    def add_images_to_template(self, worksheet, uploaded_images):
        """Add uploaded images to template at specific positions"""
        try:
            added_images = 0
            temp_image_paths = []
            
            # Fixed positions for different image types
            positions = {
                'current': 'T3',  # Current packaging at T3
                'primary': 'A42',  # Primary packaging at A42
                'secondary': 'F42',  # Secondary packaging at F42 (next column set)
                'label': 'K42'  # Label at K42 (next column set)
            }
            
            for img_key, img_data in uploaded_images.items():
                img_type = img_data.get('type', 'current')
                if img_type in positions:
                    position = positions[img_type]
                    success = self._place_image_at_position(
                        worksheet, img_key, img_data, position,
                        width_cm=4.3 if img_type != 'current' else 8.3,
                        height_cm=4.3 if img_type != 'current' else 8.3,
                        temp_image_paths=temp_image_paths
                    )
                    if success:
                        added_images += 1
            
            return added_images, temp_image_paths
            
        except Exception as e:
            st.error(f"Error adding images to template: {e}")
            return 0, []

    def _place_image_at_position(self, worksheet, img_key, img_data, cell_position, width_cm, height_cm, temp_image_paths):
        """Place a single image at the specified cell position"""
        try:
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
            
            return True
            
        except Exception as e:
            st.write(f"❌ Failed to place {img_key} at {cell_position}: {e}")
            return False
    
    def extract_images_by_part_info(self, excel_file_path, part_no=None, description=None):
        """Extract images based on part number and description from Excel sheets"""
        try:
            workbook = openpyxl.load_workbook(excel_file_path, data_only=False)
            matched_images = {}
        
            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
            
                # Check if sheet contains matching part info
                sheet_matches = False
                if part_no or description:
                    for row in worksheet.iter_rows():
                        for cell in row:
                            if cell.value:
                                cell_text = str(cell.value).lower()
                                if ((part_no and part_no.lower() in cell_text) or (description and description.lower() in cell_text)):
                                    sheet_matches = True
                                    break
                        if sheet_matches:
                            break
            
                # Extract images from matching sheets
                if sheet_matches and hasattr(worksheet, '_images') and worksheet._images:
                    for idx, img in enumerate(worksheet._images):
                        # ... (use existing image extraction logic)
                        # Add to matched_images instead of all images
                        pass
        
            workbook.close()
            return matched_images
        
        except Exception as e:
            st.error(f"Error extracting images by part info: {e}")
            return {}

class EnhancedTemplateMapperWithImages:
    def __init__(self):
        self.image_extractor = ImageExtractor()
        self.similarity_threshold = 0.3
        
        # Enhanced section-based mapping rules (from your working code)
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
                    'type': 'Secondary Packaging Type',
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

    def preprocess_text(self, text):
        """Preprocess text for better matching"""
        try:
            if pd.isna(text) or text is None:
                return ""
            
            text = str(text).lower()
            text = re.sub(r'[()[\]{}]', ' ', text)
            text = re.sub(r'[^\w\s/-]', ' ', text)
            text = re.sub(r'\s+', ' ', text).strip()
            
            return text
        except Exception as e:
            st.error(f"Error in preprocess_text: {e}")
            return ""

    def is_mappable_field(self, text):
        """Enhanced field detection for packaging templates"""
        try:
            if not text or pd.isna(text):
                return False
            
            text = str(text).lower().strip()
            if not text:
                return False
        
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
                r'packaging\s+type', r'\btype\b',
                r'\bl[-\s]*mm\b', r'\bw[-\s]*mm\b', r'\bh[-\s]*mm\b',
                r'\bl\b', r'\bw\b', r'\bh\b',
                r'part\s+l\b', r'part\s+w\b', r'part\s+h\b',
                r'\blength\b', r'\bwidth\b', r'\bheight\b',
                r'qty[/\s]*pack', r'quantity\b', r'weight\b', r'empty\s+weight',
                r'\bcode\b', r'\bname\b', r'\bdescription\b', r'\blocation\b',
                r'part\s+no\b', r'part\s+number\b'
            ]
        
            for pattern in mappable_patterns:
                if re.search(pattern, text):
                    print(f"DEBUG: '{text}' matches pattern '{pattern}'")
                    return True
        
            if text.endswith(':'):
                print(f"DEBUG: '{text}' ends with colon")
                return True
                
            print(f"DEBUG: '{text}' is NOT mappable")
            return False
        except Exception as e:
            st.error(f"Error in is_mappable_field: {e}")
            return False

    def identify_section_context(self, worksheet, row, col, max_search_rows=15):
        """Enhanced section identification with better pattern matching"""
        try:
            section_context = None
        
            for search_row in range(max(1, row - max_search_rows), row + 2):
                for search_col in range(max(1, col - 15), min(worksheet.max_column + 1, col + 15)):
                    try:
                        cell = worksheet.cell(row=search_row, column=search_col)
                        if cell.value:
                            cell_text = self.preprocess_text(str(cell.value))
                        
                            for section_name, section_info in self.section_mappings.items():
                                for keyword in section_info['section_keywords']:
                                    keyword_processed = self.preprocess_text(keyword)
                                
                                    if keyword_processed == cell_text or keyword_processed in cell_text or cell_text in keyword_processed:
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
            
            sequence_sim = SequenceMatcher(None, text1, text2).ratio()
            return sequence_sim
        except Exception as e:
            st.error(f"Error in calculate_similarity: {e}")
            return 0.0

    def find_template_fields_with_context_and_images(self, template_file):
        """Find template fields and image upload areas"""
        fields = {}
        image_areas = []
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
        
            merged_ranges = worksheet.merged_cells.ranges
        
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
                            
                                section_context = self.identify_section_context(
                                    worksheet, cell.row, cell.column
                                )
                            
                                print(f"DEBUG: Found field '{cell_value}' at {cell_coord}")
                                print(f"DEBUG: Section context: {section_context}")
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
        
            workbook.close()
        
        except Exception as e:
            st.error(f"Error reading template: {e}")
    
        return fields, image_areas

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

    def map_data_with_section_context(self, template_fields, data_df):
        """Enhanced mapping with better section-aware logic"""
        mapping_results = {}
        used_columns = set()

        try:
            data_columns = data_df.columns.tolist()
            print(f"DEBUG: Available data columns: {data_columns}")

            for coord, field in template_fields.items():
                try:
                    best_match = None
                    best_score = 0.0
                    field_value = field['value']
                    section_context = field.get('section_context')

                    print(f"DEBUG: Mapping field '{field_value}' with section '{section_context}'")

                    # If section context exists, use its field mappings
                    if section_context and section_context in self.section_mappings:
                        section_mappings = self.section_mappings[section_context]['field_mappings']
                        print(f"DEBUG: Section mappings: {section_mappings}")

                        for template_field_key, data_column_pattern in section_mappings.items():
                            normalized_field_value = self.preprocess_text(field_value)
                            normalized_template_key = self.preprocess_text(template_field_key)

                            print(f"DEBUG: Comparing '{normalized_field_value}' with '{normalized_template_key}'")

                            if normalized_field_value == normalized_template_key:
                                section_prefix = section_context.split('_')[0].capitalize()
                                expected_column = f"{section_prefix} {data_column_pattern}".strip()
                            
                                print(f"DEBUG: Looking for expected column: '{expected_column}'")

                                for data_col in data_columns:
                                    if data_col in used_columns:
                                        continue
                                    if self.preprocess_text(data_col) == self.preprocess_text(expected_column):
                                        best_match = data_col
                                        best_score = 1.0
                                        print(f"DEBUG: EXACT MATCH FOUND: {data_col}")
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
                                            print(f"DEBUG: SIMILARITY MATCH: {data_col} (score: {similarity})")
                                break

                    # Fallback 1: If 'type' and no section, assume secondary packaging
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

                    # Fallback 2: If 'L', 'W', 'H', etc. and no section, assume part_information
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

                    print(f"DEBUG: Final mapping result - Field: '{field_value}' -> Column: '{best_match}' (Score: {best_score})")
                    print("=" * 50)

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

    def clean_data_value(self, value):
        """Clean data value to handle NaN, None, and empty values"""
        if pd.isna(value) or value is None:
            return ""
        
        # Convert to string and strip whitespace
        str_value = str(value).strip()
        
        # Handle common representations of empty/null values
        if str_value.lower() in ['nan', 'none', 'null', 'n/a', '#n/a', '']:
            return ""
            
        return str_value

    def map_template_with_data(self, template_path, data_path):
        """Enhanced mapping with section-based approach and procedure steps integration"""
        try:
            # Read data from Excel with proper NaN handling
            data_df = pd.read_excel(data_path)
        
            # Replace NaN values with empty strings in the entire dataframe
            data_df = data_df.fillna("")
        
            st.write(f"📊 Loaded data with {len(data_df)} rows and {len(data_df.columns)} columns")
        
            # Load template
            workbook = openpyxl.load_workbook(template_path)
            worksheet = workbook.active
        
            st.write(f"📋 Template has {worksheet.max_row} rows and {worksheet.max_column} columns")
        
            # Find template fields with section context
            template_fields, _ = self.find_template_fields_with_context_and_images(template_path)
            st.write(f"🗺️ Found {len(template_fields)} template fields")
        
            # Map data with section context
            mapping_results = self.map_data_with_section_context(template_fields, data_df)
        
            # Apply mappings to template
            mapping_count = 0
            successful_mappings = []
            failed_mappings = []
            data_dict = {}  # Store mapped data for procedure generation
        
            for coord, mapping in mapping_results.items():
                if mapping['is_mappable'] and mapping['data_column']:
                    try:
                        # Get data value with proper NaN handling
                        data_col = mapping['data_column']
                    
                        data_col = mapping['data_column']
                        if not data_df[data_col].empty and len(data_df[data_col]) > 0:
                            # For multiple rows, you might want to concatenate or use specific logic
                            # Option 1: Use all non-empty values
                            all_values = [self.clean_data_value(val) for val in data_df[data_col] if self.clean_data_value(val)]
                            if all_values:
                                data_value = all_values[0]  # or join them: " | ".join(all_values)
                        else:
                            data_value = ""
                    
                        # Store in data_dict for procedure generation
                        data_dict[mapping['template_field']] = data_value
                    
                        # Find target cell for writing
                        target_cell_coord = self.find_data_cell_for_label(worksheet, mapping['field_info'])
                    
                        if target_cell_coord:
                            target_cell = worksheet[target_cell_coord]
                        
                            # Only write non-empty values to avoid cluttering template with empty strings
                            if data_value:  # Only write if there's actual data
                                target_cell.value = data_value
                                mapping_count += 1
                                successful_mappings.append(f"{mapping['template_field']} -> {data_col} -> {target_cell_coord}")
                                st.write(f"✅ Mapped '{mapping['template_field']}' = '{data_value}' to cell {target_cell_coord}")
                            else:
                                # Log empty values but don't write them
                                st.write(f"ℹ️ Skipped empty value for '{mapping['template_field']}' from column '{data_col}'")
                        else:
                            failed_mappings.append(mapping['template_field'])
                            st.write(f"❌ Could not find target cell for '{mapping['template_field']}'")
                        
                    except Exception as e:
                        failed_mappings.append(mapping['template_field'])
                        st.write(f"⚠️ Error writing '{mapping['template_field']}': {e}")
                else:
                    failed_mappings.append(mapping['template_field'])
        
            # ===== ADD PROCEDURE STEPS INTEGRATION HERE =====
            st.write(f"\n🔄 Adding procedure steps for packaging type...")
        
            # Get packaging type from session state (assuming it's stored there)
            if hasattr(st.session_state, 'selected_packaging_type') and st.session_state.selected_packaging_type:
                packaging_type = st.session_state.selected_packaging_type
                st.write(f"📦 Packaging type: {packaging_type}")
            
                # Write procedure steps to template
                steps_written = self.write_procedure_steps_to_template(worksheet, packaging_type, data_dict)
                st.write(f"✅ Added {steps_written} procedure steps to template")
            else:
                st.warning("⚠️ No packaging type selected - skipping procedure steps")
        
            # Summary
            st.success(f"🎉 Successfully mapped {mapping_count}/{len(mapping_results)} fields!")
        
            if successful_mappings:
                st.write("✅ Successful mappings:")
                for mapping in successful_mappings[:10]:
                    st.write(f"  - {mapping}")
                
            if failed_mappings:
                st.write("❌ Failed mappings:")
                for field in failed_mappings[:5]:
                    st.write(f"  - {field}")
        
            return workbook, mapping_results
        
        except Exception as e:
            st.error(f"❌ Error mapping template: {e}")
            st.write("📋 Traceback:", traceback.format_exc())
            return None, {}
    
    # Keep your packaging procedure methods
    def get_procedure_steps(self, packaging_type, data_dict=None):
        """Get procedure steps with data substitution"""
        # Use the PACKAGING_PROCEDURES from your constants
        procedures = PACKAGING_PROCEDURES.get(packaging_type, [""] * 11)
    
        if data_dict:
            filled_procedures = []
            for procedure in procedures:
                filled_procedure = procedure
            
                # Define all possible replacements
                replacements = {
                    '{x No. of Parts}': self.clean_data_value(data_dict.get('x No. of Parts', data_dict.get('Qty/Veh', data_dict.get('Quantity', 'XXX')))),
                    '{Inner L}': self.clean_data_value(data_dict.get('Inner L', data_dict.get('Inner Length', 'XXX'))),
                    '{Inner W}': self.clean_data_value(data_dict.get('Inner W', data_dict.get('Inner Width', 'XXX'))),
                    '{Inner H}': self.clean_data_value(data_dict.get('Inner H', data_dict.get('Inner Height', 'XXX'))),
                    '{Inner Qty/Pack}': self.clean_data_value(data_dict.get('Inner Qty/Pack', 'XXX')),
                    '{Outer L}': self.clean_data_value(data_dict.get('Outer L', data_dict.get('Outer Length', 'XXX'))),
                    '{Outer W}': self.clean_data_value(data_dict.get('Outer W', data_dict.get('Outer Width', 'XXX'))),
                    '{Outer H}': self.clean_data_value(data_dict.get('Outer H', data_dict.get('Outer Height', 'XXX'))),
                    '{Primary Qty/Pack}': self.clean_data_value(data_dict.get('Primary Qty/Pack', data_dict.get('Qty/Pack', 'XXX'))),
                    '{Layer}': self.clean_data_value(data_dict.get('Layer', 'XXX')),
                    '{Level}': self.clean_data_value(data_dict.get('Level', 'XXX')),
                    '{Qty/Pack}': self.clean_data_value(data_dict.get('Qty/Pack', 'XXX')),
                    '{Qty/Veh}': self.clean_data_value(data_dict.get('Qty/Veh', 'XXX')),
                }
            
                # Apply replacements
                for placeholder, value in replacements.items():
                    # If value is empty after cleaning, keep XXX as placeholder
                    if not value or value == '' or value == 'nan':
                        value = 'XXX'
                    filled_procedure = filled_procedure.replace(placeholder, str(value))
            
                filled_procedures.append(filled_procedure)
            return filled_procedures
        else:
            return procedures

    def write_procedure_steps_to_template(self, worksheet, packaging_type, data_dict=None):
        """Write packaging procedure steps in merged cells B to P starting from Row 28"""
        try:
            from openpyxl.cell import MergedCell
            from openpyxl.styles import Font, Alignment

            print(f"\n=== WRITING PROCEDURE STEPS FOR {packaging_type} ===")
            st.write(f"🔄 Processing procedure steps for: {packaging_type}")

            # Get procedure steps with data substitution
            steps = self.get_procedure_steps(packaging_type, data_dict)
            if not steps:
                print(f"❌ No procedure steps found for packaging type: {packaging_type}")
                st.error(f"No procedure steps found for packaging type: {packaging_type}")
                return 0

            print(f"📋 Retrieved {len(steps)} procedure steps")
            st.write(f"📋 Retrieved {len(steps)} procedure steps")

            start_row = 28
            target_col = 2  # Column B
            end_col = 16    # Column P

            # Filter out empty steps
            non_empty_steps = [step for step in steps if step and step.strip()]
            steps_to_write = non_empty_steps

            print(f"✏️  Will write {len(steps_to_write)} non-empty steps")
            st.write(f"✏️ Writing {len(steps_to_write)} non-empty steps to template")

            steps_written = 0

            for i, step in enumerate(steps_to_write):
                step_row = start_row + i
                step_text = step.strip()
            
                # Make sure we don't exceed template boundaries
                if step_row > worksheet.max_row + 20:  # Safety check
                    st.warning(f"⚠️ Stopping at row {step_row} to avoid exceeding template boundaries")
                    break
            
                try:
                    # Define the merge range for this row (B to P)
                    merge_range = f"B{step_row}:P{step_row}"
                    target_cell = worksheet.cell(row=step_row, column=target_col)
                
                    print(f"📝 Writing step {i + 1} to {merge_range}: {step_text[:50]}...")
                    st.write(f"📝 Step {i + 1} -> {merge_range}: {step_text[:50]}...")

                    # First, check if this range is already merged and unmerge if necessary
                    existing_merged_ranges = []
                    for merged_range in list(worksheet.merged_cells.ranges):
                        # Check if any part of our target range overlaps with existing merged ranges
                        if (merged_range.min_row <= step_row <= merged_range.max_row and
                            merged_range.min_col <= end_col and merged_range.max_col >= target_col):
                            existing_merged_ranges.append(merged_range)

                    # Unmerge overlapping ranges
                    for merged_range in existing_merged_ranges:
                        try:
                            worksheet.unmerge_cells(str(merged_range))
                            print(f"🔧 Unmerged existing range: {merged_range}")
                        except Exception as unmerge_error:
                            print(f"⚠️ Warning: Could not unmerge {merged_range}: {unmerge_error}")

                    # Clear any existing content in the range
                    for col in range(target_col, end_col + 1):
                        cell = worksheet.cell(row=step_row, column=col)
                        cell.value = None

                    # Write the step text to the first cell (B)
                    target_cell.value = step_text
                    target_cell.font = Font(name='Calibri', size=10)
                    target_cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

                    # Merge the cells B to P for this row
                    try:
                        worksheet.merge_cells(merge_range)
                        print(f"✅ Merged range: {merge_range}")
                        st.write(f"✅ Merged cells: {merge_range}")
                    except Exception as merge_error:
                        print(f"⚠️ Warning: Could not merge {merge_range}: {merge_error}")
                        st.warning(f"Could not merge {merge_range}: {merge_error}")
  
                    # Adjust row height based on text length
                    # Calculate approximate number of lines needed
                    chars_per_line = 120  # Approximate characters per line in merged B:P range
                    num_lines = max(1, len(step_text) // chars_per_line + 1)
                    estimated_height = 15 + (num_lines - 1) * 15
                    worksheet.row_dimensions[step_row].height = estimated_height

                    steps_written += 1
                
                except Exception as step_error:
                    print(f"❌ Error writing step {i + 1}: {step_error}")
                    st.error(f"Error writing step {i + 1}: {step_error}")
                    import traceback
                    traceback.print_exc()
                    continue

            print(f"\n✅ PROCEDURE STEPS COMPLETED")
            print(f"   Total steps written: {steps_written}")
            print(f"   Location: Merged cells B:P, starting from Row 28")
        
            st.success(f"✅ Successfully wrote {steps_written} procedure steps to template")

            return steps_written

        except Exception as e:
            print(f"💥 Critical error in write_procedure_steps_to_template: {e}")
            st.error(f"Critical error writing procedure steps: {e}")
            import traceback
            traceback.print_exc()
            return 0
    
    def write_manual_procedure_steps_to_template(self, worksheet, manual_procedures):
        """Write manually entered procedure steps to template"""
        try:
            from openpyxl.styles import Font, Alignment
        
            print(f"\n=== WRITING MANUAL PROCEDURE STEPS ===")
            st.write(f"🔄 Processing manual procedure steps...")
        
            start_row = 28
            target_col = 2  # Column B
            end_col = 16    # Column P
        
            # Filter out empty steps
            steps_to_write = [step.strip() for step in manual_procedures if step and step.strip()]
        
            print(f"✏️ Will write {len(steps_to_write)} manual procedure steps")
            st.write(f"✏️ Writing {len(steps_to_write)} manual procedure steps to template")
        
            steps_written = 0
        
            for i, step in enumerate(steps_to_write):
                step_row = start_row + i
                step_text = step.strip()
            
                # Make sure we don't exceed template boundaries
                if step_row > worksheet.max_row + 20:  # Safety check
                    st.warning(f"⚠️ Stopping at row {step_row} to avoid exceeding template boundaries")
                    break
            
                try:
                    # Define the merge range for this row (B to P)
                    merge_range = f"B{step_row}:P{step_row}"
                    target_cell = worksheet.cell(row=step_row, column=target_col)
                
                    print(f"📝 Writing manual step {i + 1} to {merge_range}: {step_text[:50]}...")
                    st.write(f"📝 Manual Step {i + 1} -> {merge_range}: {step_text[:50]}...")
                
                    # First, check if this range is already merged and unmerge if necessary
                    existing_merged_ranges = []
                    for merged_range in list(worksheet.merged_cells.ranges):
                        # Check if any part of our target range overlaps with existing merged ranges
                        if (merged_range.min_row <= step_row <= merged_range.max_row and merged_range.min_col <= end_col and merged_range.max_col >= target_col):
                            existing_merged_ranges.append(merged_range)
                
                    # Unmerge overlapping ranges
                    for merged_range in existing_merged_ranges:
                        try:
                            worksheet.unmerge_cells(str(merged_range))
                            print(f"🔧 Unmerged existing range: {merged_range}")
                        except Exception as unmerge_error:
                            print(f"⚠️ Warning: Could not unmerge {merged_range}: {unmerge_error}")
                
                    # Clear any existing content in the range
                    for col in range(target_col, end_col + 1):
                        cell = worksheet.cell(row=step_row, column=col)
                        cell.value = None
                
                    # Write the step text to the first cell (B)
                    target_cell.value = step_text
                    target_cell.font = Font(name='Calibri', size=10)
                    target_cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                
                    # Merge the cells B to P for this row
                    try:
                        worksheet.merge_cells(merge_range)
                        print(f"✅ Merged range: {merge_range}")
                    except Exception as merge_error:
                        print(f"⚠️ Warning: Could not merge {merge_range}: {merge_error}")
                        st.warning(f"Could not merge {merge_range}: {merge_error}")
                
                    # Adjust row height based on text length
                    chars_per_line = 120  # Approximate characters per line in merged B:P range
                    num_lines = max(1, len(step_text) // chars_per_line + 1)
                    estimated_height = 15 + (num_lines - 1) * 15
                    worksheet.row_dimensions[step_row].height = estimated_height
                
                    steps_written += 1
                
                except Exception as step_error:
                    print(f"❌ Error writing manual step {i + 1}: {step_error}")
                    st.error(f"Error writing manual step {i + 1}: {step_error}")
                    continue
        
            print(f"\n✅ MANUAL PROCEDURE STEPS COMPLETED")
            print(f"   Total steps written: {steps_written}")
        
            st.success(f"✅ Successfully wrote {steps_written} manual procedure steps to template")
        
            return steps_written
        
        except Exception as e:
            print(f"💥 Critical error in write_manual_procedure_steps_to_template: {e}")
            st.error(f"Critical error writing manual procedure steps: {e}")
            import traceback
            traceback.print_exc()
            return 0
        
# Packaging types and procedures from reference code
PACKAGING_TYPES = [
    "BOX IN BOX SENSITIVE",
    "BOX IN BOX", 
    "CARTON BOX WITH SEPARATOR FOR ONE PART",
    "INDIVIDUAL NOT SENSITIVE",
    "INDIVIDUAL PROTECTION FOR EACH PART MANY TYPE",
    "INDIVIDUAL PROTECTION FOR EACH PART",
    "INDIVIDUAL SENSITIVE",
    "MANY IN ONE TYPE",
    "SINGLE BOX",
    "ADD NEW TEMPLATE" 
]

PACKAGING_PROCEDURES = {
    "BOX IN BOX SENSITIVE": [
        "Pick up {x No. of Parts} quantity of part and apply bubble wrapping over it",
        "Apply tape and Put 1 such bubble wrapped part into a carton box [L-{Inner L} mm, W-{Inner W} mm, H-{Inner H} mm]",
        "Seal carton box and put {Inner Qty/Pack} such carton boxes into another carton box [L-{Outer L} mm, W-{Outer W} mm, H-{Outer H} mm]",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Primary Qty/Pack})",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",
        "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level",
        "Put corner / edge protector and apply pet strap (2 times -- cross way)",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
    ],
    "BOX IN BOX": [
        "Pick up {x No. of Parts} quantity of part and put it in a polybag",
        "Seal the polybag and put it into a carton box [L-{Inner L} mm, W-{Inner W} mm, H-{Inner H} mm]",
        "Put {Inner Qty/Pack} such carton boxes into another carton box [L-{Outer L} mm, W-{Outer W} mm, H-{Outer H} mm]",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Primary Qty/Pack})",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",
        "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level",
        "Put corner / edge protector and apply pet strap (2 times -- cross way)",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
    ],
    "CARTON BOX WITH SEPARATOR FOR ONE PART": [
        "Pick up {x No. of Parts} parts and apply bubble wrapping over it (individually)",
        "Apply tape and Put bubble wrapped part into a carton box. Apply part separator & filler material between two parts to arrest part movement during handling",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Primary Qty/Pack})",
        "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",
        "Put corner / edge protector and apply pet strap (2 times -- cross way)",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
    ],
    "INDIVIDUAL NOT SENSITIVE": [
        "Pick up {x No. of Parts} part and put it into a polybag",
        "Seal polybag and Put polybag into a carton box",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Primary Qty/Pack})",
        "Load carton boxes on base wooden pallet -- Maximum {Layer} boxes per layer & Maximum {Level} level",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",
        "Put corner / edge protector and apply pet strap (2 times -- cross way)",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
    ],
    "INDIVIDUAL PROTECTION FOR EACH PART MANY TYPE": [
        "Pick up {x No. of Parts} parts and apply bubble wrapping over it (individually)",
        "Apply tape and Put bubble wrapped part into a carton box. Apply part separator & filler material between two parts to arrest part movement during handling",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule (multiple of primary pack quantity -- {Qty/Pack})",
        "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",
        "Put corner / edge protector and apply pet strap (2 times -- cross way)",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
    ],
    "INDIVIDUAL PROTECTION FOR EACH PART": [
        "Pick up {x No. of Parts} parts and apply bubble wrapping over it (individually)",
        "Apply tape and Put bubble wrapped part into a carton box. Apply part separator & filler material between two parts to arrest part movement during handling",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Primary Qty/Pack})",
        "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",
        "Put corner / edge protector and apply pet strap (2 times -- cross way)",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
    ],
    "INDIVIDUAL SENSITIVE": [
        "Pick up {x No. of Parts} part and apply bubble wrapping over it",
        "Apply tape and Put bubble wrapped part into a carton box",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Primary Qty/Pack})",
        "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",
        "Put corner / edge protector and apply pet strap (2 times -- cross way)",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
    ],
    "MANY IN ONE TYPE": [
        "Pick up {x No. of Parts} quantity of part and put it in a polybag",
        "Seal polybag and Put it into a carton box",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Primary Qty/Pack})",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",
        "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level",
        "Put corner / edge protector and apply pet strap (2 times -- cross way)",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
    ],
    "SINGLE BOX": [
        "Pick up {x No. of Parts} quantity of part and put it in a polybag",
        "Put into a carton box",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Primary Qty/Pack})",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",
        "Load carton boxes on base wooden pallet -- {Layer} boxes per layer & max {Level} level",
        "Put corner / edge protector and apply pet strap (2 times -- cross way) and stretch wrap it",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
    ]
}

def main():
    # Header
    st.title("📦 AI Packaging Template Mapper")
    st.markdown("---")
    
    # Progress indicator
    steps = [
        "Select Packaging Type",
        "Upload Template File", 
        "Upload Data File",
        "Auto-Fill Template",
        "Choose Image Option",
        "Generate Final Document"
    ]
    
    # Create progress bar
    progress_cols = st.columns(len(steps))
    for i, (col, step) in enumerate(zip(progress_cols, steps)):
        with col:
            if i + 1 < st.session_state.current_step:
                st.success(f"✅ {i+1}. {step}")
            elif i + 1 == st.session_state.current_step:
                st.info(f"🔄 {i+1}. {step}")
            else:
                st.write(f"⏳ {i+1}. {step}")
    
    st.markdown("---")
    
    # Step 1: Select Packaging Type
    if st.session_state.current_step == 1:
        st.header("📦 Step 1: Select Packaging Type")
    
        # Create columns for packaging types
        cols = st.columns(3)
        for i, packaging_type in enumerate(PACKAGING_TYPES):
            with cols[i % 3]:
                if st.button(packaging_type, key=f"pkg_{i}", use_container_width=True):
                    st.session_state.selected_packaging_type = packaging_type
                
                    # Handle Add New Template option
                    if packaging_type == "ADD NEW TEMPLATE":
                        st.session_state.is_new_template = True
                        navigate_to_step(2)  # Go to template upload
                    else:
                        st.session_state.is_new_template = False
                        navigate_to_step(2)
    
        # Show selected packaging details (skip for ADD NEW TEMPLATE)
        if (st.session_state.selected_packaging_type and 
            st.session_state.selected_packaging_type != "ADD NEW TEMPLATE"):
            st.success(f"Selected: {st.session_state.selected_packaging_type}")
        
            with st.expander("View Packaging Procedure"):
                procedures = PACKAGING_PROCEDURES.get(st.session_state.selected_packaging_type, [])
                for i, step in enumerate(procedures, 1):
                    st.write(f"{i}. {step}")
        elif st.session_state.selected_packaging_type == "ADD NEW TEMPLATE":
            st.info("📝 You selected to add a new template. You'll manually fill 11 procedure steps later.")
    
    # Step 2: Upload Template File
    elif st.session_state.current_step == 2:
        st.header("📄 Step 2: Upload Template File")
        
        st.info(f"Selected Packaging Type: {st.session_state.selected_packaging_type}")
        
        uploaded_template = st.file_uploader(
            "Choose template file (Excel or Word)",
            type=['xlsx', 'xls', 'docx'],
            key="template_upload"
        )
        
        if uploaded_template is not None:
            # Save uploaded file
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_template.name.split('.')[-1]}") as tmp_file:
                tmp_file.write(uploaded_template.getvalue())
                st.session_state.template_file = tmp_file.name
            
            st.success("✅ Template file uploaded successfully!")
            
            if st.button("Continue to Data Upload", key="continue_to_step3"):
                navigate_to_step(3)
        
        # Back navigation
        if st.button("← Go Back", key="back_from_2"):
            navigate_to_step(1)
    
    # Step 3: Upload Data File
    elif st.session_state.current_step == 3:
        st.header("📊 Step 3: Upload Data File (Excel)")
        
        uploaded_data = st.file_uploader(
            "Choose Excel data file",
            type=['xlsx', 'xls'],
            key="data_upload"
        )
        
        if uploaded_data is not None:
            # Save uploaded file
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_data.name.split('.')[-1]}") as tmp_file:
                tmp_file.write(uploaded_data.getvalue())
                st.session_state.data_file = tmp_file.name
            
            st.success("✅ Data file uploaded successfully!")
            
            try:
                df = pd.read_excel(st.session_state.data_file)
                st.write("Data Preview:")
                st.dataframe(df.head())
            except Exception as e:
                st.error(f"Error reading data file: {e}")
            
            if st.button("Continue to Auto-Fill", key="continue_to_step4"):
                navigate_to_step(4)
        
        # Back navigation
        if st.button("← Go Back", key="back_from_3"):
            navigate_to_step(2)
    
    # Step 4: Auto-Fill Template
    elif st.session_state.current_step == 4:
        st.header("🔄 Step 4: Auto-Fill Template")
        
        # Check if mapping is already completed
        if st.session_state.mapping_completed and st.session_state.mapped_data:
            st.success("✅ Template auto-filling completed!")
            
            # Show mapped fields if available
            if hasattr(st.session_state, 'last_mapped_fields') and st.session_state.last_mapped_fields:
                with st.expander("View Mapped Fields"):
                    for field, value in st.session_state.last_mapped_fields.items():
                        st.write(f"**{field}**: {value}")
            
            # Always show the continue button if mapping is completed
            if st.button("Continue to Image Options", key="continue_to_images"):
                navigate_to_step(5)
        
        else:
            # Show the start button if mapping hasn't been completed
            if st.button("Start Auto-Fill Process", key="start_autofill"):
                st.session_state.auto_fill_started = True
                
                with st.spinner("Processing template and data mapping..."):
                    try:
                        mapper = EnhancedTemplateMapperWithImages()
                        
                        # Map template with data
                        workbook, mapped_fields = mapper.map_template_with_data(
                            st.session_state.template_file,
                            st.session_state.data_file
                        )
                        
                        if workbook:
                            # Save the mapped workbook
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                                workbook.save(tmp_file.name)
                                st.session_state.mapped_data = tmp_file.name
                            
                            # Mark as completed and store mapped fields
                            st.session_state.mapping_completed = True
                            st.session_state.last_mapped_fields = mapped_fields
                            
                            st.success(f"✅ Template auto-filled with {len(mapped_fields)} data fields!")
                            st.rerun()  # Refresh to show the continue button
                        else:
                            st.error("Failed to process template mapping")
                            
                    except Exception as e:
                        st.error(f"Error during auto-fill: {e}")
                        st.write("Traceback:", traceback.format_exc())
        
        # Back navigation
        if st.button("← Go Back", key="back_from_4"):
            navigate_to_step(3)
    
    # Step 4.5: Manual Procedure Entry (only for ADD NEW TEMPLATE)
    elif (st.session_state.current_step == 4 and st.session_state.selected_packaging_type == "ADD NEW TEMPLATE" and st.session_state.mapping_completed):
        st.header("✏️ Step 4.5: Add Procedure Steps")
        st.info("Enter 11 procedure steps for your new template:")
    
        # Initialize procedure steps in session state if not exists
        if 'manual_procedures' not in st.session_state:
            st.session_state.manual_procedures = [""] * 11
    
        # Create input fields for 11 steps
        for i in range(11):
            st.session_state.manual_procedures[i] = st.text_area(
                f"Step {i+1}:",
                value=st.session_state.manual_procedures[i],
                key=f"manual_step_{i}",
                height=80
            )
    
        if st.button("Save Procedure Steps"):
            # Validate at least some steps are filled
            filled_steps = [step for step in st.session_state.manual_procedures if step.strip()]
            if len(filled_steps) >= 3:
                st.success(f"✅ Saved {len(filled_steps)} procedure steps!")
                navigate_to_step(5)
            else:
                st.error("Please fill at least 3 procedure steps before continuing.")
                
    # Step 5: Choose Image Option
    elif st.session_state.current_step == 5:
        st.header("🖼️ Step 5: Choose Image Option")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("Extract Images from Data File", use_container_width=True):
                st.session_state.image_option = 'extract'
                
                # Extract images from data file
                with st.spinner("Extracting images from Excel file..."):
                    extractor = ImageExtractor()
                    extracted_images = extractor.extract_images_from_excel(st.session_state.data_file)
                    
                    if extracted_images and 'all_sheets' in extracted_images:
                        st.session_state.extracted_excel_images = extracted_images['all_sheets']
                        st.success(f"✅ Extracted {len(st.session_state.extracted_excel_images)} images!")
                        
                        # Preview extracted images
                        st.write("**Extracted Images Preview:**")
                        for img_key, img_data in st.session_state.extracted_excel_images.items():
                            with st.expander(f"Image: {img_key}"):
                                st.image(f"data:image/png;base64,{img_data['data']}", 
                                       caption=f"Size: {img_data['size']}, Type: {img_data['type']}")
                    else:
                        st.warning("No images found in the Excel file")
        
        with col2:
            if st.button("Upload New Images", use_container_width=True):
                st.session_state.image_option = 'upload'
        
        # Handle upload new images option
        if st.session_state.image_option == 'upload':
            st.subheader("Upload Images")
            
            # Image upload for different types
            image_types = ['current', 'primary', 'secondary', 'label']
            
            for img_type in image_types:
                uploaded_img = st.file_uploader(
                    f"Upload {img_type.capitalize()} Packaging Image",
                    type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
                    key=f"img_upload_{img_type}"
                )
                
                if uploaded_img is not None:
                    # Convert to base64
                    img_bytes = uploaded_img.read()
                    img_b64 = base64.b64encode(img_bytes).decode()
                    
                    # Store in session state
                    st.session_state.uploaded_images[f"{img_type}_uploaded"] = {
                        'data': img_b64,
                        'format': uploaded_img.type.split('/')[-1].upper(),
                        'size': len(img_bytes),
                        'type': img_type
                    }
                    
                    # Preview
                    st.image(f"data:image/{uploaded_img.type.split('/')[-1]};base64,{img_b64}", 
                           caption=f"{img_type.capitalize()} Image", width=200)
        
        # Continue button
        if (st.session_state.image_option == 'extract' and st.session_state.extracted_excel_images) or \
           (st.session_state.image_option == 'upload' and st.session_state.uploaded_images):
            if st.button("Continue to Final Generation", key="continue_to_step6"):
                navigate_to_step(6)
        
        # Back navigation
        if st.button("← Go Back", key="back_from_5"):
            navigate_to_step(4)
    
    # Step 6: Generate Final Document
    elif st.session_state.current_step == 6:
        st.header("📋 Step 6: Generate Final Document")
    
        # Show summary of what will be generated
        st.subheader("📊 Generation Summary")
        col1, col2 = st.columns(2)
    
        with col1:
            st.write(f"**Packaging Type**: {st.session_state.selected_packaging_type}")
            st.write(f"**Template File**: Uploaded ✅")
            st.write(f"**Data File**: Uploaded ✅")
            st.write(f"**Auto-Fill**: Completed ✅")
    
        with col2:
            image_count = 0
            if st.session_state.image_option == 'extract' and st.session_state.extracted_excel_images:
                image_count = len(st.session_state.extracted_excel_images)
            elif st.session_state.image_option == 'upload' and st.session_state.uploaded_images:
                image_count = len(st.session_state.uploaded_images)
        
            st.write(f"**Image Option**: {st.session_state.image_option.capitalize()}")
            st.write(f"**Images Ready**: {image_count} images")
        
            # Show procedure steps info
            if st.session_state.selected_packaging_type == "ADD NEW TEMPLATE":
                manual_steps = getattr(st.session_state, 'manual_procedures', [])
                filled_steps = len([step for step in manual_steps if step and step.strip()])
                st.write(f"**Manual Steps**: {filled_steps}/11 filled")
            else:
                procedure_steps = len(PACKAGING_PROCEDURES.get(st.session_state.selected_packaging_type, []))
                st.write(f"**Procedure Steps**: {procedure_steps} auto-generated")
    
        st.markdown("---")
    
        # Main generation button
        if st.button("🚀 Generate Final Template with Images", type="primary", use_container_width=True):
            with st.spinner("🔄 Generating final document..."):
                try:
                    # Load the mapped template
                    workbook = openpyxl.load_workbook(st.session_state.mapped_data)
                    worksheet = workbook.active
                
                    st.write("📋 Loaded mapped template...")
                
                    # Add images based on selected option
                    extractor = ImageExtractor()
                    images_to_add = {}
                    added_count = 0
                    temp_paths = []
                
                    if st.session_state.image_option == 'extract':
                        images_to_add = st.session_state.extracted_excel_images
                        st.write(f"🖼️ Using {len(images_to_add)} extracted images...")
                    elif st.session_state.image_option == 'upload':
                        images_to_add = st.session_state.uploaded_images
                        st.write(f"🖼️ Using {len(images_to_add)} uploaded images...")
                
                    if images_to_add:
                        added_count, temp_paths = extractor.add_images_to_template(
                            worksheet, images_to_add
                        )
                        st.success(f"✅ Added {added_count} images to template!")
                    else:
                        st.info("ℹ️ No images to add - proceeding with text-only template")
                
                    # Handle procedure steps based on template type
                    if st.session_state.selected_packaging_type == "ADD NEW TEMPLATE":
                        # Write manual procedure steps
                        if hasattr(st.session_state, 'manual_procedures'):
                            steps_written = self.write_manual_procedure_steps_to_template(
                                worksheet, st.session_state.manual_procedures
                            )
                            st.write(f"✅ Added {steps_written} manual procedure steps")
                        else:
                            st.warning("⚠️ No manual procedure steps found")
                    else:
                        # Auto-generated procedure steps (existing logic)
                        st.write(f"🔄 Adding auto-generated procedure steps...")
                        # This part remains the same as your existing code
                
                    # Generate dynamic filename based on vendor code, part no, and description
                    try:
                        # Extract vendor code, part no, and description from data
                        data_df = pd.read_excel(st.session_state.data_file)
                    
                        vendor_code = ""
                        part_no = ""
                        description = ""
                    
                        # Look for these fields in the data
                        for col in data_df.columns:
                            col_lower = col.lower()
                            if any(keyword in col_lower for keyword in ['vendor', 'supplier']) and 'code' in col_lower:
                                vendor_code = str(data_df[col].iloc[0]) if not data_df[col].empty else ""
                            elif 'part' in col_lower and ('no' in col_lower or 'number' in col_lower):
                                part_no = str(data_df[col].iloc[0]) if not data_df[col].empty else ""
                            elif any(keyword in col_lower for keyword in ['description', 'desc']):
                                description = str(data_df[col].iloc[0]) if not data_df[col].empty else ""
                    
                        # Clean the values for filename (remove special characters)
                        vendor_code = re.sub(r'[^\w\-_]', '', vendor_code)[:15]
                        part_no = re.sub(r'[^\w\-_]', '', part_no)[:20]
                        description = re.sub(r'[^\w\-_]', '', description)[:25]
                    
                        # Create filename: vendorcode_partno_desc.xlsx
                        if vendor_code and part_no:
                            if description:
                                final_filename = f"{vendor_code}_{part_no}_{description}.xlsx"
                            else:
                                final_filename = f"{vendor_code}_{part_no}.xlsx"
                        else:
                            # Fallback filename if vendor code or part no is missing
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            final_filename = f"Packaging_Template_{timestamp}.xlsx"
                    
                        st.write(f"📄 Generated filename: {final_filename}")
                        
                    except Exception as filename_error:
                        # Fallback filename in case of any error
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        final_filename = f"Packaging_Template_{timestamp}.xlsx"
                        st.warning(f"⚠️ Using fallback filename: {final_filename}")
                
                    #  Save final document
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                        workbook.save(tmp_file.name)
                        st.session_state.final_template_path = tmp_file.name
                    
                        # Read file for download
                        with open(tmp_file.name, 'rb') as f:
                            file_bytes = f.read()
                
                    workbook.close()
                
                    # Cleanup temporary image files
                    for temp_path in temp_paths:
                        try:
                            os.unlink(temp_path)
                        except:
                            pass
                
                    st.success("🎉 Final template generated successfully!")
                
                    # Download section with enhanced UI
                    st.markdown("---")
                    st.subheader("📥 Download Your Template")
                
                    # Create download columns
                    download_col1, download_col2 = st.columns([2, 1])
                
                    with download_col1:
                        # Main download button
                        st.download_button(
                            label=f"📥 Download {final_filename}",
                            data=file_bytes,
                            file_name=final_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                
                    with download_col2:
                        # File info
                        file_size_mb = len(file_bytes) / (1024 * 1024)
                        st.metric("File Size", f"{file_size_mb:.2f} MB")
                
                    # Detailed generation summary
                    with st.expander("📋 Detailed Generation Summary", expanded=True):
                        summary_col1, summary_col2 = st.columns(2)
                    
                        with summary_col1:
                            st.write("**Template Information:**")
                            st.write(f"• Packaging Type: {st.session_state.selected_packaging_type}")
                            st.write(f"• Images Added: {added_count}")
                            st.write(f"• Image Source: {st.session_state.image_option.capitalize()}")
                        
                            if hasattr(st.session_state, 'last_mapped_fields'):
                                mapped_count = len([m for m in st.session_state.last_mapped_fields.values() if m.get('is_mappable')])
                                st.write(f"• Fields Mapped: {mapped_count}")
                    
                        with summary_col2:
                            st.write("**File Information:**")
                            st.write(f"• Filename: {final_filename}")
                            st.write(f"• Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                            st.write(f"• File Size: {file_size_mb:.2f} MB")
                        
                            # Show procedure steps count
                            if st.session_state.selected_packaging_type == "ADD NEW TEMPLATE":
                                manual_steps = getattr(st.session_state, 'manual_procedures', [])
                                filled_manual_steps = len([step for step in manual_steps if step and step.strip()])
                                st.write(f"• Manual Steps: {filled_manual_steps}/11")
                            else:
                                auto_steps = len(PACKAGING_PROCEDURES.get(st.session_state.selected_packaging_type, []))
                                st.write(f"• Auto Steps: {auto_steps}")
                
                    # Show preview of mapped data if available
                    if hasattr(st.session_state, 'last_mapped_fields') and st.session_state.last_mapped_fields:
                        with st.expander("🗺️ View Mapped Fields"):
                            for coord, mapping in st.session_state.last_mapped_fields.items():
                                if mapping.get('is_mappable') and mapping.get('data_column'):
                                    st.write(f"**{mapping['template_field']}** ← {mapping['data_column']}")
                
                    # Additional download options
                    st.markdown("---")
                    st.subheader("🔄 Additional Options")
                
                    option_col1, option_col2, option_col3 = st.columns(3)
                
                    with option_col1:
                        # Option to generate another template with same data
                        if st.button("📋 Generate Another Template", use_container_width=True):
                            # Keep data and template files, go back to step 1
                            st.session_state.current_step = 1
                            # Clear only packaging-specific selections
                            if 'selected_packaging_type' in st.session_state:
                                del st.session_state['selected_packaging_type']
                            if 'manual_procedures' in st.session_state:
                                del st.session_state['manual_procedures']
                            st.rerun()
                
                    with option_col2:
                        # Option to start completely over
                        if st.button("🔄 Start New Project", use_container_width=True):
                            # Clear all session state
                            for key in list(st.session_state.keys()):
                                if key != 'current_step':
                                    del st.session_state[key]
                            st.session_state.current_step = 1
                            st.rerun()
                
                    with option_col3:
                        # Option to go back and modify
                        if st.button("← Modify Images", use_container_width=True):
                            navigate_to_step(5)
                
                    # Success message with tips
                    st.markdown("---")
                    st.success("""
                               🎉 **Template Generated Successfully!**
                               Your packaging template is ready with:
                               ✅ Auto-filled data fields
                               ✅ Packaging procedure steps  
                               ✅ Positioned images
                               ✅ Professional formatting
                               💡 **Tip**: The filename follows the format `vendorcode_partno_description.xlsx` for easy identification.
                    """)
                except Exception as e:
                    st.error(f"❌ Error generating final document: {e}")
                    st.write("**Error Details:**")
                    st.code(traceback.format_exc())
                
                    # Provide fallback download if possible
                    if hasattr(st.session_state, 'mapped_data'):
                        st.warning("🔧 Providing fallback download without images...")
                        try:
                            with open(st.session_state.mapped_data, 'rb') as f:
                                fallback_bytes = f.read()
                        
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            fallback_filename = f"Packaging_Template_Fallback_{timestamp}.xlsx"
                        
                            st.download_button(
                                label=f"📥 Download Fallback Template",
                                data=fallback_bytes,
                                file_name=fallback_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        except Exception as fallback_error:
                            st.error(f"Fallback download also failed: {fallback_error}")

        # Back navigation
        if st.button("← Go Back", key="back_from_6"):
            navigate_to_step(5)
    
    # Sidebar with help and information
    with st.sidebar:
        st.header("ℹ️ Help & Information")
        
        st.subheader("Current Progress")
        st.write(f"**Step**: {st.session_state.current_step}/6")
        if st.session_state.selected_packaging_type:
            st.write(f"**Packaging Type**: {st.session_state.selected_packaging_type}")
        
        st.subheader("Instructions")
        st.write("""
        1. **Select Packaging Type**: Choose from predefined packaging types
        2. **Upload Template**: Upload your Excel template file
        3. **Upload Data**: Upload Excel file with part data
        4. **Auto-Fill**: Let AI map data to template fields
        5. **Add Images**: Extract from Excel or upload new images
        6. **Generate**: Create final template with images
        """)
        
        st.subheader("Supported Formats")
        st.write("**Template Files**: .xlsx, .xls, .docx")
        st.write("**Data Files**: .xlsx, .xls")
        st.write("**Image Files**: .png, .jpg, .jpeg, .gif, .bmp")
        
        # Reset button
        if st.button("🔄 Reset All", type="secondary"):
            for key in list(st.session_state.keys()):
                if key != 'current_step':
                    del st.session_state[key]
            st.session_state.current_step = 1
            st.rerun()

if __name__ == "__main__":
    main()
