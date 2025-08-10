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

class EnhancedTemplateMapperWithImages:
    def __init__(self):
        self.image_extractor = ImageExtractor()
        self.similarity_threshold = 0.3
        
        # Enhanced section-based mapping rules (from first code)
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
            # Remove parentheses and special characters but keep spaces
            text = re.sub(r'[()[\]{}]', ' ', text)
            text = re.sub(r'[^\w\s/-]', ' ', text)
            text = re.sub(r'\s+', ' ', text).strip()
            
            return text
        except Exception as e:
            print(f"Error in preprocess_text: {e}")
            return ""

    def identify_section_context(self, worksheet, row, col, max_search_rows=15):
        """Enhanced section identification with better pattern matching"""
        try:
            # Search upwards and in nearby cells for section headers
            for search_row in range(max(1, row - max_search_rows), row + 2):
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
        
            return None
        
        except Exception as e:
            print(f"Error in identify_section_context: {e}")
            return None

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
            print(f"Error in is_mappable_field: {e}")
            return False

    def find_template_fields_with_context_and_images(self, template_file):
        """Find template fields and image upload areas using the improved logic from first code"""
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
        
            # Find image upload areas
            image_areas = self.image_extractor.identify_image_upload_areas(worksheet)
        
            workbook.close()
        
        except Exception as e:
            st.error(f"Error reading template: {e}")
    
        return fields, image_areas

    def find_data_cell_for_label(self, worksheet, field_info):
        """Find data cell for a label with improved merged cell handling - FIXED VERSION"""
        try:
            row = field_info['row']
            col = field_info['column']
            merged_ranges = list(worksheet.merged_cells.ranges)
        
            def is_suitable_data_cell(target_row, target_col):
                """Check if a cell is suitable for data entry"""
                try:
                    if target_row < 1 or target_col < 1 or target_row > worksheet.max_row or target_col > worksheet.max_column:
                        return False
                        
                    cell = worksheet.cell(row=target_row, column=target_col)
                    
                    # Check if it's a merged cell - if so, skip it for now
                    if hasattr(cell, '__class__') and cell.__class__.__name__ == 'MergedCell':
                        return False
                    
                    # Empty cells are good candidates
                    if cell.value is None or str(cell.value).strip() == "":
                        return True
                        
                    # Check for data placeholder patterns
                    cell_text = str(cell.value).lower().strip()
                    data_patterns = [r'^_+$', r'^\.*$', r'^-+$', r'enter', r'fill', r'data', r'xxx']
                    if any(re.search(pattern, cell_text) for pattern in data_patterns):
                        return True
                        
                    return False
                except Exception as e:
                    print(f"Error checking cell suitability: {e}")
                    return False
            
            # Strategy 1: Look right of label (most common pattern) - ENHANCED
            print(f"Looking for data cell for label '{field_info['value']}' at row {row}, col {col}")
            
            # Try immediate right cells first
            for offset in range(1, 6):
                target_col = col + offset
                if is_suitable_data_cell(row, target_col):
                    cell_coord = worksheet.cell(row=row, column=target_col).coordinate
                    print(f"Found suitable cell to the right: {cell_coord}")
                    return cell_coord
            
            # Strategy 2: Look below label
            for offset in range(1, 4):
                target_row = row + offset
                if is_suitable_data_cell(target_row, col):
                    cell_coord = worksheet.cell(row=target_row, column=col).coordinate
                    print(f"Found suitable cell below: {cell_coord}")
                    return cell_coord
            
            # Strategy 3: Look diagonally (down-right)
            for row_offset in range(1, 3):
                for col_offset in range(1, 4):
                    target_row = row + row_offset
                    target_col = col + col_offset
                    if is_suitable_data_cell(target_row, target_col):
                        cell_coord = worksheet.cell(row=target_row, column=target_col).coordinate
                        print(f"Found suitable cell diagonally: {cell_coord}")
                        return cell_coord
            
            # Strategy 4: Handle merged cells specially
            # Look for merged ranges that might be data input areas
            for merged_range in merged_ranges:
                min_col, min_row, max_col, max_row = merged_range.bounds
                
                # Check if this merged range is in the vicinity of our label
                if (abs(min_row - row) <= 2 and abs(min_col - col) <= 5):
                    # Get the top-left cell of the merged range
                    anchor_cell = worksheet.cell(row=min_row, column=min_col)
                    if anchor_cell.value is None or str(anchor_cell.value).strip() == "":
                        print(f"Found suitable merged cell: {anchor_cell.coordinate}")
                        return anchor_cell.coordinate
            
            print(f"No suitable data cell found for label '{field_info['value']}'")
            return None
            
        except Exception as e:
            print(f"Error in find_data_cell_for_label: {e}")
            return None

    def map_data_with_section_context(self, template_fields, data_df):
        """Enhanced mapping with better section-aware logic - FIXED VERSION"""
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
                        print(f"DEBUG: Section mappings: {list(section_mappings.keys())}")

                        for template_field_key, data_column_pattern in section_mappings.items():
                            normalized_field_value = self.preprocess_text(field_value)
                            normalized_template_key = self.preprocess_text(template_field_key)

                            print(f"DEBUG: Comparing '{normalized_field_value}' with '{normalized_template_key}'")

                            if normalized_field_value == normalized_template_key or normalized_template_key in normalized_field_value:
                                # Look for exact column match first
                                for data_col in data_columns:
                                    if data_col in used_columns:
                                        continue
                                    if self.preprocess_text(data_col) == self.preprocess_text(data_column_pattern):
                                        best_match = data_col
                                        best_score = 1.0
                                        print(f"DEBUG: EXACT MATCH FOUND: {data_col}")
                                        break

                                # If no exact match, try similarity matching
                                if not best_match:
                                    for data_col in data_columns:
                                        if data_col in used_columns:
                                            continue
                                        similarity = self.calculate_similarity(data_column_pattern, data_col)
                                        if similarity > best_score and similarity >= self.similarity_threshold:
                                            best_score = similarity
                                            best_match = data_col
                                            print(f"DEBUG: SIMILARITY MATCH: {data_col} (score: {similarity})")
                                
                                if best_match:
                                    break

                    # Fallback mappings for fields without clear section context
                    if not best_match:
                        # Handle common cases
                        normalized_field = self.preprocess_text(field_value)
                        
                        if normalized_field == 'type' and not section_context:
                            # Assume secondary packaging type
                            for data_col in data_columns:
                                if data_col in used_columns:
                                    continue
                                if 'secondary' in data_col.lower() and 'type' in data_col.lower():
                                    best_match = data_col
                                    best_score = 1.0
                                    break
                        
                        elif normalized_field in ['l', 'w', 'h'] and not section_context:
                            # Assume part dimensions
                            for data_col in data_columns:
                                if data_col in used_columns:
                                    continue
                                if f'part {normalized_field}' == data_col.lower():
                                    best_match = data_col
                                    best_score = 1.0
                                    break

                        # General similarity matching as final fallback
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
                    print(f"Error mapping field {coord}: {e}")
                    continue

        except Exception as e:
            print(f"Error in map_data_with_section_context: {e}")

        return mapping_results

    def calculate_similarity(self, text1, text2):
        """Calculate similarity between two texts"""
        try:
            if not text1 or not text2:
                return 0.0
            
            text1 = self.preprocess_text(text1)
            text2 = self.preprocess_text(text2)
            
            if not text1 or not text2:
                return 0.0
            
            # Use SequenceMatcher for similarity
            from difflib import SequenceMatcher
            return SequenceMatcher(None, text1, text2).ratio()
            
        except Exception as e:
            print(f"Error in calculate_similarity: {e}")
            return 0.0

    def fill_template_with_data_and_images(self, template_file, mapping_results, data_df, uploaded_images=None, packaging_type=None):
        """Fill template with mapped data, images, and procedure steps - FIXED VERSION"""
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
        
            # Fill data fields - ENHANCED LOGIC
            for coord, mapping in mapping_results.items():
                try:
                    if mapping['data_column'] is not None and mapping['is_mappable']:
                        field_info = mapping['field_info']
                    
                        # Find the target cell where data should be written
                        target_cell_coord = self.find_data_cell_for_label(worksheet, field_info)
                    
                        if target_cell_coord and len(data_df) > 0:
                            # Get the data value
                            data_value = data_df.iloc[0][mapping['data_column']]
                            
                            # Write to the target cell
                            cell_obj = worksheet[target_cell_coord]
                            
                            # Handle merged cells properly
                            if hasattr(cell_obj, '__class__') and cell_obj.__class__.__name__ == 'MergedCell':
                                # Find the anchor cell of the merged range
                                for merged_range in worksheet.merged_cells.ranges:
                                    if target_cell_coord in merged_range:
                                        min_col, min_row, max_col, max_row = merged_range.bounds
                                        anchor_cell = worksheet.cell(row=min_row, column=min_col)
                                        anchor_cell.value = str(data_value) if not pd.isna(data_value) else ""
                                        print(f"✅ Filled merged cell {anchor_cell.coordinate} with '{data_value}' for field '{field_info['value']}'")
                                        filled_count += 1
                                        break
                            else:
                                # Regular cell
                                cell_obj.value = str(data_value) if not pd.isna(data_value) else ""
                                print(f"✅ Filled cell {target_cell_coord} with '{data_value}' for field '{field_info['value']}'")
                                filled_count += 1
                        else:
                            print(f"❌ Could not find target cell for field '{field_info['value']}'")
                        
                except Exception as e:
                    print(f"Error filling mapping {coord}: {e}")
                    continue
        
            # Add images if provided
            if uploaded_images:
                # First, identify image upload areas
                _, image_areas = self.find_template_fields_with_context_and_images(template_file)
                images_added, temp_image_paths = self.add_images_to_template(worksheet, uploaded_images, image_areas)
        
            # Write procedure steps if packaging type is provided
            if packaging_type and packaging_type != "Select Packaging Procedure":
                try:
                    procedure_steps_added = self.write_procedure_steps_to_template(worksheet, packaging_type, data_dict)
                    print(f"Added {procedure_steps_added} procedure steps for packaging type: {packaging_type}")
                except Exception as e:
                    print(f"Error adding procedure steps: {e}")
                    procedure_steps_added = 0
            
            return workbook, filled_count, images_added, temp_image_paths, procedure_steps_added
        
        except Exception as e:
            print(f"Error filling template: {e}")
            return None, 0, 0, [], 0

    # Include other necessary methods from the first code...
    def write_procedure_steps_to_template(self, worksheet, packaging_type, data_dict=None):
        """Write packaging procedure steps (implementation from first code)"""
        # Implementation would be the same as in your first code
        pass
    
    def add_images_to_template(self, worksheet, uploaded_images, image_areas):
        """Add images to template (implementation from first code)"""
        # Implementation would be the same as in your first code
        pass

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
    "SINGLE BOX"
]

PACKAGING_PROCEDURES = {
    "BOX IN BOX SENSITIVE": [
        "Pick up 1 quantity of part and apply bubble wrapping over it",
        "Apply tape and Put 1 such bubble wrapped part into a carton box [L-{Inner L} mm, W-{Inner W} mm, H-{Inner H} mm]",
        "Seal carton box and put {Inner Qty/Pack} such carton boxes into another carton box [L-{Outer L} mm, W-{Outer W} mm, H-{Outer H} mm]",
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
        "Put {Inner Qty/Pack} such carton boxes into another carton box [L-{Outer L} mm, W-{Outer W} mm, H-{Outer H} mm]",
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
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only."
    ],
    "INDIVIDUAL NOT SENSITIVE": [
        "Pick up one part and put it into a polybag",
        "Seal polybag and Put polybag into a carton box",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule (multiple of pack quantity -- {Inner Qty/Pack})",
        "Load carton boxes on base wooden pallet -- Maximum 20 boxes per layer & Maximum 5 level (max height including pallet - 1000 mm)",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet.",
        "Put corner / edge protector and apply pet strap (2 times -- cross way)",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only."
    ],
    "INDIVIDUAL PROTECTION FOR EACH PART MANY TYPE": [
        "Pick up {Qty/Veh} parts and apply bubble wrapping over it (individually)",
        "Apply tape and Put bubble wrapped part into a carton box. Apply part separator & filler material between two parts to arrest part movement during handling",
        "Seal carton box and put Traceability label as per PMSPL standard guideline",
        "Prepare additional carton boxes in line with procurement schedule ( multiple of primary pack quantity – {Qty/Pack})",
        "Load carton boxes on base wooden pallet – {Layer} boxes per layer & max {Level} level",
        "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same wooden pallet",
        "Put corner / edge protector and apply pet strap ( 2 times – cross way)",
        "Apply traceability label on complete pack",
        "Attach packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
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
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only."
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
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only."
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
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only."
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
        "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only."
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
                    navigate_to_step(2)
        
        # Show selected packaging details
        if st.session_state.selected_packaging_type:
            st.success(f"Selected: {st.session_state.selected_packaging_type}")
            
            with st.expander("View Packaging Procedure"):
                procedures = PACKAGING_PROCEDURES.get(st.session_state.selected_packaging_type, [])
                for i, step in enumerate(procedures, 1):
                    st.write(f"{i}. {step}")
    
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
        
        if st.button("Generate Final Template with Images"):
            with st.spinner("Generating final document..."):
                try:
                    # Load the mapped template
                    workbook = openpyxl.load_workbook(st.session_state.mapped_data)
                    worksheet = workbook.active
                    
                    # Add images based on selected option
                    extractor = ImageExtractor()
                    images_to_add = {}
                    
                    if st.session_state.image_option == 'extract':
                        images_to_add = st.session_state.extracted_excel_images
                    elif st.session_state.image_option == 'upload':
                        images_to_add = st.session_state.uploaded_images
                    
                    if images_to_add:
                        added_count, temp_paths = extractor.add_images_to_template(
                            worksheet, images_to_add
                        )
                        st.success(f"✅ Added {added_count} images to template!")
                    
                    # Save final document
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    final_filename = f"Packaging_Template_{st.session_state.selected_packaging_type.replace(' ', '_')}_{timestamp}.xlsx"
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                        workbook.save(tmp_file.name)
                        
                        # Read file for download
                        with open(tmp_file.name, 'rb') as f:
                            file_bytes = f.read()
                    
                    # Provide download button
                    st.download_button(
                        label="📥 Download Final Template",
                        data=file_bytes,
                        file_name=final_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("🎉 Final template generated successfully!")
                    
                    # Show summary
                    with st.expander("Generation Summary"):
                        st.write(f"**Packaging Type**: {st.session_state.selected_packaging_type}")
                        st.write(f"**Images Added**: {added_count if 'added_count' in locals() else 0}")
                        st.write(f"**Template File**: {final_filename}")
                        st.write(f"**Generated On**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    
                    # Option to start over
                    if st.button("🔄 Start New Template"):
                        # Clear session state
                        for key in list(st.session_state.keys()):
                            if key.startswith(('current_step', 'selected_', 'template_', 'data_', 'mapped_', 'image_', 'uploaded_', 'extracted_')):
                                del st.session_state[key]
                        st.session_state.current_step = 1
                        st.rerun()
                        
                except Exception as e:
                    st.error(f"Error generating final document: {e}")
                    st.write("Traceback:", traceback.format_exc())
        
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
