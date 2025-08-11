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
    
    def extract_images_for_part(self, excel_file_path, part_no, description, vendor_code=None):
        """Extract images specific to a part number, description, or vendor code"""
        try:
            all_images = self.extract_images_from_excel(excel_file_path)
            if not all_images or 'all_sheets' not in all_images:
                return {}

            search_terms = [
                str(term).lower().strip()
                for term in [vendor_code, part_no, description]
                if term and str(term).strip()
            ]

            part_specific_images = {
                key: data
                for key, data in all_images['all_sheets'].items()
                if any(term in key.lower() for term in search_terms)
            }

            return part_specific_images if part_specific_images else {}
        
        except Exception as e:
            st.error(f"Error extracting images for part {part_no}: {e}")
            return {}

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
        """Enhanced mapping with section-based approach and multiple row processing"""
        try:
            # Read data from Excel with proper NaN handling
            data_df = pd.read_excel(data_path)
            data_df = data_df.fillna("")
        
            st.write(f"📊 Loaded data with {len(data_df)} rows and {len(data_df.columns)} columns")
        
            # Store all row data for multi-template generation
            st.session_state.all_row_data = []
        
            # Process each row
            for row_idx in range(len(data_df)):
                st.write(f"🔄 Processing row {row_idx + 1}/{len(data_df)}")
            
                # Load fresh template for each row
                workbook = openpyxl.load_workbook(template_path)
                worksheet = workbook.active
            
                # Find template fields with section context
                template_fields, _ = self.find_template_fields_with_context_and_images(template_path)
            
                # Map data with section context for current row
                mapping_results = self.map_data_with_section_context_for_row(template_fields, data_df, row_idx)
            
                # Apply mappings to template
                mapping_count = 0
                data_dict = {}  # Store mapped data for procedure generation
                filename_parts = {}  # Store parts for filename
            
                for coord, mapping in mapping_results.items():
                    if mapping['is_mappable'] and mapping['data_column']:
                        try:
                            data_col = mapping['data_column']
                            raw_value = data_df[data_col].iloc[row_idx]  # Use current row
                            data_value = self.clean_data_value(raw_value)
                        
                            # Store in data_dict for procedure generation
                            data_dict[mapping['template_field']] = data_value
                        
                            # Store filename components
                            field_name_lower = mapping['template_field'].lower()
                            if any(term in field_name_lower for term in ['vendor code', 'supplier code', 'code']):
                                filename_parts['vendor_code'] = data_value
                            elif 'part' in field_name_lower and ('no' in field_name_lower or 'number' in field_name_lower):
                                filename_parts['part_no'] = data_value
                            elif 'description' in field_name_lower or 'desc' in field_name_lower:
                                filename_parts['description'] = data_value
                        
                            # Find target cell and write data
                            target_cell_coord = self.find_data_cell_for_label(worksheet, mapping['field_info'])
                        
                            if target_cell_coord and data_value:
                                target_cell = worksheet[target_cell_coord]
                                target_cell.value = data_value
                                mapping_count += 1
                            
                        except Exception as e:
                            st.write(f"⚠️ Error processing row {row_idx + 1}, field '{mapping['template_field']}': {e}")
            
                # Add procedure steps
                if hasattr(st.session_state, 'selected_packaging_type') and st.session_state.selected_packaging_type:
                    steps_written = self.write_procedure_steps_to_template(worksheet, st.session_state.selected_packaging_type, data_dict)
            
                # Generate filename
                vendor_code = filename_parts.get('vendor_code', 'Unknown')
                part_no = filename_parts.get('part_no', 'Unknown') 
                description = filename_parts.get('description', 'Unknown')
            
                # Clean filename parts
                vendor_code = re.sub(r'[^\w\-_]', '', str(vendor_code))[:10]
                part_no = re.sub(r'[^\w\-_]', '', str(part_no))[:15]
                description = re.sub(r'[^\w\-_]', '', str(description))[:20]
            
                filename = f"{vendor_code}_{part_no}_{description}.xlsx"
            
                # Save workbook to temporary file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    workbook.save(tmp_file.name)
                
                    # Store row data
                    row_data = {
                        'row_index': row_idx,
                        'filename': filename,
                        'file_path': tmp_file.name,
                        'data_dict': data_dict,
                        'mapping_count': mapping_count,
                        'vendor_code': vendor_code,
                        'part_no': part_no,
                        'description': description
                    }
                    st.session_state.all_row_data.append(row_data)
            
                workbook.close()
                st.write(f"✅ Row {row_idx + 1} processed: {mapping_count} fields mapped -> {filename}")
        
            st.success(f"🎉 Successfully processed {len(data_df)} rows!")
            return True, st.session_state.all_row_data
        
        except Exception as e:
            st.error(f"❌ Error mapping template: {e}")
            st.write("📋 Traceback:", traceback.format_exc())
            return False, []
    def map_data_with_section_context_for_row(self, template_fields, data_df, row_idx):
        """Map data for specific row"""
        mapping_results = {}
        used_columns = set()

        try:
            data_columns = data_df.columns.tolist()

            for coord, field in template_fields.items():
                try:
                    best_match = None
                    best_score = 0.0
                    field_value = field['value']
                    section_context = field.get('section_context')

                    # Use existing mapping logic but for specific row
                    if section_context and section_context in self.section_mappings:
                        section_mappings = self.section_mappings[section_context]['field_mappings']

                        for template_field_key, data_column_pattern in section_mappings.items():
                            normalized_field_value = self.preprocess_text(field_value)
                            normalized_template_key = self.preprocess_text(template_field_key)

                            if normalized_field_value == normalized_template_key:
                                section_prefix = section_context.split('_')[0].capitalize()
                                expected_column = f"{section_prefix} {data_column_pattern}".strip()

                                for data_col in data_columns:
                                    if data_col in used_columns:
                                        continue
                                    if self.preprocess_text(data_col) == self.preprocess_text(expected_column):
                                        best_match = data_col
                                        best_score = 1.0
                                        break

                                if not best_match:
                                    for data_col in data_columns:
                                        if data_col in used_columns:
                                            continue
                                        similarity = self.calculate_similarity(expected_column, data_col)
                                        if similarity > best_score and similarity >= self.similarity_threshold:
                                            best_score = similarity
                                            best_match = data_col
                                break

                    # Fallback logic (same as original)
                    if not best_match:
                        for data_col in data_columns:
                            if data_col in used_columns:
                                continue
                            similarity = self.calculate_similarity(field_value, data_col)
                            if similarity > best_score and similarity >= self.similarity_threshold:
                                best_score = similarity
                                best_match = data_col

                    mapping_results[coord] = {
                        'template_field': field_value,
                        'data_column': best_match,
                        'similarity': best_score,
                        'field_info': field,
                        'section_context': section_context,
                        'is_mappable': best_match is not None
                    }

                    if best_match:
                        used_columns.add(best_match)

                except Exception as e:
                    st.error(f"Error mapping field {coord}: {e}")
                    continue

        except Exception as e:
            st.error(f"Error in map_data_with_section_context_for_row: {e}")

        return mapping_results
    
    # Keep your packaging procedure methods
    def get_procedure_steps(self, packaging_type, data_dict=None):
        """Get procedure steps with data substitution"""
        procedures = PACKAGING_PROCEDURES.get(packaging_type, [""] * 11)
        if not data_dict:
            return procedures

        filled_procedures = []
        for procedure in procedures:
            filled_procedure = procedure

            # Direct mapping: most placeholders match exactly to data_dict keys
            replacements = {
                '{x No. of Parts}': data_dict.get('x No. of Parts') or data_dict.get('Qty/Veh') or data_dict.get('Quantity'),
                '{Inner L}': data_dict.get('Inner L'),
                '{Inner W}': data_dict.get('Inner W'),
                '{Inner H}': data_dict.get('Inner H'),
                '{Inner Qty/Pack}': data_dict.get('Inner Qty/Pack'),
                '{Outer L}': data_dict.get('Outer L'),
                '{Outer W}': data_dict.get('Outer W'),
                '{Outer H}': data_dict.get('Outer H'),
                '{Primary Qty/Pack}': data_dict.get('Primary Qty/Pack'),  # ✅ Only this comes from primary
                '{Layer}': data_dict.get('Layer'),
                '{Level}': data_dict.get('Level'),
                '{Qty/Pack}': data_dict.get('Qty/Pack'),
                '{Qty/Veh}': data_dict.get('Qty/Veh'),
            }

            # Clean and replace
            for placeholder, raw_value in replacements.items():
                value = self.clean_data_value(raw_value)
                if not value:
                    value = 'XXX'
                filled_procedure = filled_procedure.replace(placeholder, str(value))

            filled_procedures.append(filled_procedure)

        return filled_procedures

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
    
        if st.session_state.mapping_completed and hasattr(st.session_state, 'all_row_data'):
            st.success(f"✅ Template auto-filling completed for {len(st.session_state.all_row_data)} rows!")
        
            # Show summary of processed rows
            with st.expander("View Processed Rows Summary"):
                for i, row_data in enumerate(st.session_state.all_row_data):
                    st.write(f"**Row {i+1}**: {row_data['filename']} ({row_data['mapping_count']} fields mapped)")
        
            if st.button("Continue to Image Options", key="continue_to_images"):
                navigate_to_step(5)
    
        else:
            if st.button("Start Auto-Fill Process", key="start_autofill"):
                with st.spinner("Processing template and data mapping for all rows..."):
                    try:
                        mapper = EnhancedTemplateMapperWithImages()
                    
                        success, all_row_data = mapper.map_template_with_data(
                            st.session_state.template_file,
                            st.session_state.data_file
                        )
                    
                        if success and all_row_data:
                            st.session_state.mapping_completed = True
                            st.session_state.all_row_data = all_row_data
                            st.rerun()
                        else:
                            st.error("Failed to process template mapping")
                        
                    except Exception as e:
                        st.error(f"Error during auto-fill: {e}")
                        st.write("Traceback:", traceback.format_exc())
    
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
        st.header("📋 Step 6: Generate Final Documents")
    
        if st.button("Generate All Templates with Images"):
            with st.spinner("Generating all documents..."):
                try:
                    extractor = ImageExtractor()
                    generated_files = []
                
                    for row_data in st.session_state.all_row_data:
                        # Load the mapped template for this row
                        workbook = openpyxl.load_workbook(row_data['file_path'])
                        worksheet = workbook.active
                    
                        # Determine images to add
                        images_to_add = {}
                    
                        if st.session_state.image_option == 'extract':
                            # Extract images specific to this part
                            images_to_add = extractor.extract_images_for_part(
                                st.session_state.data_file,
                                row_data['part_no'],
                                row_data['description'],
                                row_data['vendor_code']
                            )
                        elif st.session_state.image_option == 'upload':
                            # Use same uploaded images for all templates
                            images_to_add = st.session_state.uploaded_images
                    
                        # Add images to template
                        if images_to_add:
                            added_count, temp_paths = extractor.add_images_to_template(
                                worksheet, images_to_add
                            )
                            st.write(f"✅ Added {added_count} images to {row_data['filename']}")
                    
                        # Save final document
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        final_filename = f"{row_data['filename'].replace('.xlsx', '')}_{timestamp}.xlsx"
                    
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                            workbook.save(tmp_file.name)
                        
                            with open(tmp_file.name, 'rb') as f:
                                file_bytes = f.read()
                        
                            generated_files.append({
                                'filename': final_filename,
                                'data': file_bytes,
                                'row_info': row_data
                            })
                    
                        workbook.close()
                
                    st.success(f"🎉 Generated {len(generated_files)} final templates!")
                
                    # Create download buttons for each file
                    st.subheader("📥 Download Generated Templates")
                    for file_info in generated_files:
                        col1, col2 = st.columns([3, 1])
                        with col1:
                            st.write(f"**{file_info['filename']}**")
                            st.caption(f"Vendor: {file_info['row_info']['vendor_code']} | Part: {file_info['row_info']['part_no']}")
                        with col2:
                            st.download_button(
                                label="📥 Download",
                                data=file_info['data'],
                                file_name=file_info['filename'],
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"download_{file_info['filename']}"
                            )
                
                    # Option to download all as ZIP
                    if len(generated_files) > 1:
                        if st.button("📦 Download All as ZIP"):
                            import zipfile
                            zip_buffer = io.BytesIO()
                        
                            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                for file_info in generated_files:
                                    zip_file.writestr(file_info['filename'], file_info['data'])
                        
                            zip_buffer.seek(0)
                            st.download_button(
                                label="📥 Download ZIP File",
                                data=zip_buffer.getvalue(),
                                file_name=f"All_Templates_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                                mime="application/zip"
                            )
                
                except Exception as e:
                    st.error(f"Error generating final documents: {e}")
                    st.write("Traceback:", traceback.format_exc())
    
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
