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
        
        # Enhanced field mappings based on the template structure
        self.field_mappings = {
            # Vendor Information mappings
            'vendor code': ['Code', 'vendor code'],
            'vendor name': ['Name', 'vendor name'],  
            'vendor location': ['Location', 'vendor location'],
            
            # Part Information mappings
            'part no': ['Part No.', 'part number', 'part no'],
            'part number': ['Part No.', 'part number', 'part no'],
            'description': ['Description', 'part description'],
            'unit weight': ['Unit Weight', 'weight'],
            'part l': ['L', 'length', 'part l'],
            'part w': ['W', 'width', 'part w'], 
            'part h': ['H', 'height', 'part h'],
            
            # Primary Packaging mappings
            'primary packaging type': ['Packaging Type', 'primary type'],
            'primary l-mm': ['L-mm', 'primary length', 'length mm'],
            'primary w-mm': ['W-mm', 'primary width', 'width mm'],
            'primary h-mm': ['H-mm', 'primary height', 'height mm'],
            'primary qty/pack': ['Qty/Pack', 'primary quantity'],
            'primary empty weight': ['Empty Weight', 'primary empty weight'],
            'primary pack weight': ['Pack Weight', 'primary pack weight'],
            
            # Secondary Packaging mappings
            'secondary packaging type': ['Packaging Type', 'secondary type'],
            'secondary l-mm': ['L-mm', 'secondary length'],
            'secondary w-mm': ['W-mm', 'secondary width'],
            'secondary h-mm': ['H-mm', 'secondary height'],
            'secondary qty/pack': ['Qty/Pack', 'secondary quantity'],
            'secondary empty weight': ['Empty Weight', 'secondary empty weight'],
            'secondary pack weight': ['Pack Weight', 'secondary pack weight'],
            
            # Common dimension mappings
            'length': ['L-mm', 'L', 'length'],
            'width': ['W-mm', 'W', 'width'],
            'height': ['H-mm', 'H', 'height'],
            'l': ['L-mm', 'L'],
            'w': ['W-mm', 'W'],
            'h': ['H-mm', 'H']
        }

    def is_merged_cell(self, worksheet, row, col):
        """Check if a cell is part of a merged range"""
        cell = worksheet.cell(row=row, column=col)
        for merged_range in worksheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return True, merged_range
        return False, None

    def get_writable_cell_in_merged_range(self, worksheet, merged_range):
        """Get the top-left cell of a merged range (the writable one)"""
        min_col, min_row, max_col, max_row = merged_range.bounds
        return worksheet.cell(row=min_row, column=min_col)

    def find_template_sections(self, worksheet):
        """Identify different sections in the template"""
        sections = {
            'vendor_info': [],
            'part_info': [],
            'primary_packaging': [],
            'secondary_packaging': [],
            'packaging_procedure': []
        }
        
        section_keywords = {
            'vendor_info': ['vendor information', 'code', 'name', 'location'],
            'part_info': ['part information', 'part no', 'description', 'unit weight'],
            'primary_packaging': ['primary packaging instruction', 'primary', 'internal'],
            'secondary_packaging': ['secondary packaging instruction', 'secondary', 'outer', 'external'],
            'packaging_procedure': ['packaging procedure', 'procedure']
        }
        
        for row_num in range(1, worksheet.max_row + 1):
            for col_num in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                if cell.value and isinstance(cell.value, str):
                    cell_text = str(cell.value).lower().strip()
                    
                    for section_name, keywords in section_keywords.items():
                        if any(keyword.lower() in cell_text for keyword in keywords):
                            sections[section_name].append((row_num, col_num, cell_text))
        
        return sections

    def find_field_cells(self, worksheet, search_terms, start_row=1, end_row=None, start_col=1, end_col=None):
        """Find cells containing specific field names"""
        if end_row is None:
            end_row = worksheet.max_row
        if end_col is None:
            end_col = worksheet.max_column
            
        found_cells = []
        
        for row_num in range(start_row, end_row + 1):
            for col_num in range(start_col, end_col + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                if cell.value and isinstance(cell.value, str):
                    cell_text = str(cell.value).lower().strip()
                    
                    for term in search_terms:
                        if term.lower() in cell_text or cell_text in term.lower():
                            found_cells.append((row_num, col_num, cell_text))
                            break
        
        return found_cells

    def find_adjacent_writable_cell(self, worksheet, row, col, max_distance=3):
        """Find the nearest writable cell to place data"""
        directions = [
            (0, 1),   # Right
            (0, 2),   # Right +2
            (1, 0),   # Down
            (1, 1),   # Down-Right
            (0, -1),  # Left
            (-1, 0),  # Up
        ]
        
        for distance in range(1, max_distance + 1):
            for dr, dc in directions:
                try:
                    target_row = row + (dr * distance)
                    target_col = col + (dc * distance)
                    
                    if target_row < 1 or target_col < 1:
                        continue
                    
                    target_cell = worksheet.cell(row=target_row, column=target_col)
                    
                    # Check if it's a merged cell
                    is_merged, merged_range = self.is_merged_cell(worksheet, target_row, target_col)
                    
                    if is_merged:
                        # Get the writable cell in merged range
                        writable_cell = self.get_writable_cell_in_merged_range(worksheet, merged_range)
                        if not writable_cell.value:  # Only use if empty
                            return writable_cell
                    else:
                        # Regular cell
                        if not target_cell.value:  # Only use if empty
                            return target_cell
                            
                except Exception:
                    continue
        
        return None

    def map_template_with_data(self, template_path, data_path):
        """Enhanced mapping with dynamic field detection and smart target cell finding"""
        try:
            # Read data from Excel
            data_df = pd.read_excel(data_path)
            st.write(f"📊 Loaded data with {len(data_df)} rows and {len(data_df.columns)} columns")
        
            # Load template
            workbook = openpyxl.load_workbook(template_path)
            worksheet = workbook.active
        
            st.write(f"📋 Template has {worksheet.max_row} rows and {worksheet.max_column} columns")
            st.write(f"🔗 Found {len(worksheet.merged_cells.ranges)} merged cell ranges")
        
            # Create data mapping
            mapped_fields = {}
            data_map = {}
         
            # Process each row in data
            for index, row in data_df.iterrows():
                for col_name, value in row.items():
                    if pd.notna(value) and col_name:
                        clean_col = str(col_name).lower().strip()
                        clean_value = str(value).strip()
                        data_map[clean_col] = clean_value
                        mapped_fields[clean_col] = clean_value
        
            st.write(f"📝 Created data map with {len(data_map)} fields")
        
            # Build template field map with precise target cells
            template_field_map = self.build_template_field_map(worksheet)
            st.write(f"🗺️ Found {len(template_field_map)} template fields")
        
            # Apply mappings using smart matching
            mapping_count = 0
        
            for data_field, data_value in data_map.items():
                # Try direct mapping first
                target_cell = self.find_target_cell_for_field(worksheet, data_field, template_field_map)
            
                if target_cell:
                    success = self.write_to_target_cell(worksheet, target_cell, data_value, data_field)
                    if success:
                        mapping_count += 1
                else:
                    # Try enhanced field mapping
                    mapped_target = self.try_enhanced_field_mapping(worksheet, data_field, data_value)
                    if mapped_target:
                        mapping_count += 1
        
            st.success(f"🎉 Successfully mapped {mapping_count} fields to template!")
        
            return workbook, mapped_fields
        
        except Exception as e:
            st.error(f"❌ Error mapping template: {e}")
            st.write("📋 Traceback:", traceback.format_exc())
            return None, {}

    def build_template_field_map(self, worksheet):
        """Build a comprehensive map of template fields and their target cells"""
        template_fields = {}
    
        # Scan all cells to find field labels and their corresponding input cells
        for row_num in range(1, worksheet.max_row + 1):
            for col_num in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
            
                if cell.value and isinstance(cell.value, str):
                    cell_text = str(cell.value).lower().strip()
                
                    # Skip cells that look like they contain data (not labels)
                    if self.looks_like_data_cell(cell_text):
                        continue
                
                    # Find potential input cells for this label
                    input_cells = self.find_input_cells_for_label(worksheet, row_num, col_num, cell_text)
                
                    if input_cells:
                        # Store all variations of the field name
                        field_variations = self.generate_field_variations(cell_text)
                        for variation in field_variations:
                            if variation not in template_fields:
                                template_fields[variation] = input_cells
    
        return template_fields

    def looks_like_data_cell(self, text):
        """Check if text looks like data rather than a field label"""
        # Skip very long texts, numbers, dates, etc.
        if len(text) > 50:
            return True
        if text.replace('.', '').replace('-', '').isdigit():
            return True
        if any(word in text for word in ['mm', 'kg', 'gm', 'cm', 'inch']):
            return True
        return False

    def generate_field_variations(self, field_text):
        """Generate variations of field names for better matching"""
        variations = set()
    
        # Original text
        variations.add(field_text)
    
        # Remove common suffixes/prefixes
        clean_text = field_text.replace(':', '').replace('*', '').replace('(', '').replace(')', '')
        variations.add(clean_text)
    
        # Add variations with common words
        if 'packaging' in clean_text:
            variations.add(clean_text.replace('packaging', 'pack'))
        if 'instruction' in clean_text:
            variations.add(clean_text.replace(' instruction', ''))
        if 'information' in clean_text:
            variations.add(clean_text.replace(' information', ''))
    
        # Add short versions
        words = clean_text.split()
        if len(words) > 1:
            variations.add(words[0])  # First word only
            variations.add(words[-1])  # Last word only
    
        return variations

    def find_input_cells_for_label(self, worksheet, label_row, label_col, label_text):
        """Find the most likely input cells for a given label"""
        input_cells = []
    
        # Define search patterns - cells where users typically input data
        search_patterns = [
            # Right side patterns (most common)
            (0, 1), (0, 2), (0, 3),  # Same row, 1-3 columns right
            # Below patterns  
            (1, 0), (1, 1),          # One row down, same or next column
            # Diagonal patterns
            (1, -1), (0, -1),        # Down-left, left (for right-aligned labels)
        ]
    
        for row_offset, col_offset in search_patterns:
            target_row = label_row + row_offset
            target_col = label_col + col_offset
        
            if target_row < 1 or target_col < 1:
                continue
            
            try:
                target_cell = worksheet.cell(row=target_row, column=target_col)
            
                # Check if this looks like an input cell
                if self.is_likely_input_cell(worksheet, target_cell, target_row, target_col):
                    input_cells.append(target_cell)
                
                    # For primary match, break after finding first good candidate
                    if row_offset == 0 and col_offset in [1, 2]:
                        break
                    
            except Exception:
                continue
    
        return input_cells

    def is_likely_input_cell(self, worksheet, cell, row, col):
        """Determine if a cell is likely meant for user input"""
        # Empty cells are good candidates
        if not cell.value:
            return True
    
        # Cells with placeholder text
        if cell.value and isinstance(cell.value, str):
            placeholder_indicators = ['xxx', '___', 'tbd', 'enter', 'input', '?']
            if any(indicator in str(cell.value).lower() for indicator in placeholder_indicators):
                return True
    
        # Check if it's in a merged range (often used for input fields)
        is_merged, merged_range = self.is_merged_cell(worksheet, row, col)
        if is_merged:
            # If it's the top-left cell of a merged range, it's likely an input cell
            min_col, min_row, max_col, max_row = merged_range.bounds
            if row == min_row and col == min_col:
                return True
    
        return False

    def find_target_cell_for_field(self, worksheet, data_field, template_field_map):
        """Find the best target cell for a data field"""
        data_field_clean = data_field.lower().strip()
    
        # Direct match
        if data_field_clean in template_field_map:
            return template_field_map[data_field_clean][0]  # Return first match
    
        # Try field mappings
        if data_field_clean in self.field_mappings:
            template_terms = self.field_mappings[data_field_clean]
        
            for term in template_terms:
                term_clean = term.lower().strip()
                if term_clean in template_field_map:
                    return template_field_map[term_clean][0]
    
        # Fuzzy matching with template fields
        best_match_cell = None
        best_similarity = 0
    
        for template_field, input_cells in template_field_map.items():
            similarity = SequenceMatcher(None, data_field_clean, template_field).ratio()
        
            if similarity > best_similarity and similarity > self.similarity_threshold:
                best_similarity = similarity
                best_match_cell = input_cells[0]
    
        return best_match_cell

    def try_enhanced_field_mapping(self, worksheet, data_field, data_value):
        """Enhanced mapping for fields not found in template map"""
        # Use the original fuzzy search as fallback
        best_match = self.find_best_fuzzy_match(worksheet, data_field)
    
        if best_match:
            row_num, col_num, match_text, similarity = best_match
            if similarity > self.similarity_threshold:
                writable_cell = self.find_adjacent_writable_cell(worksheet, row_num, col_num)
            
                if writable_cell:
                    return self.write_to_target_cell(worksheet, writable_cell, data_value, data_field)
    
        return False

    def write_to_target_cell(self, worksheet, target_cell, data_value, data_field):
        """Write data to target cell with proper error handling"""
        try:
            # Check if target cell is part of merged range
            is_merged, merged_range = self.is_merged_cell(worksheet, target_cell.row, target_cell.column)
        
            if is_merged:
                # Get the writable cell in merged range
                writable_cell = self.get_writable_cell_in_merged_range(worksheet, merged_range)
                if writable_cell:
                    writable_cell.value = data_value
                    st.write(f"✅ Mapped '{data_field}' = '{data_value}' to merged cell {writable_cell.coordinate}")
                    return True
            else:
                # Regular cell
                target_cell.value = data_value
                st.write(f"✅ Mapped '{data_field}' = '{data_value}' to cell {target_cell.coordinate}")
                return True
            
        except Exception as e:
            st.write(f"⚠️ Failed to write '{data_field}' to {target_cell.coordinate}: {e}")
        
        return False

    def find_best_fuzzy_match(self, worksheet, search_term):
        """Find the best fuzzy match for a search term"""
        best_match = None
        best_similarity = 0
        
        for row_num in range(1, worksheet.max_row + 1):
            for col_num in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                if cell.value and isinstance(cell.value, str):
                    cell_text = str(cell.value).lower().strip()
                    
                    # Calculate similarity
                    similarity = SequenceMatcher(None, search_term.lower(), cell_text).ratio()
                    
                    if similarity > best_similarity and similarity > 0.3:
                        best_similarity = similarity
                        best_match = (row_num, col_num, cell_text, similarity)
        
        return best_match

    def add_packaging_procedure(self, worksheet, packaging_type, data_map):
        """Add packaging procedure text to the template"""
        try:
            if packaging_type in PACKAGING_PROCEDURES:
                procedures = PACKAGING_PROCEDURES[packaging_type]
                
                # Find procedure section (around rows with "Packaging Procedure")
                procedure_start_row = None
                for row_num in range(1, worksheet.max_row + 1):
                    for col_num in range(1, worksheet.max_column + 1):
                        cell = worksheet.cell(row=row_num, column=col_num)
                        if cell.value and isinstance(cell.value, str):
                            if 'packaging procedure' in str(cell.value).lower():
                                procedure_start_row = row_num + 1
                                break
                    if procedure_start_row:
                        break
                
                if procedure_start_row:
                    # Add procedures starting from the identified row
                    for i, procedure in enumerate(procedures[:10], 1):  # Limit to first 10 steps
                        try:
                            procedure_row = procedure_start_row + i - 1
                            if procedure_row <= worksheet.max_row:
                                # Find the first cell in the procedure row (usually column A or B)
                                step_cell = worksheet.cell(row=procedure_row, column=1)
                                desc_cell = worksheet.cell(row=procedure_row, column=2)
                                
                                # Format procedure text with data substitution
                                formatted_procedure = self.format_procedure_text(procedure, data_map)
                                
                                # Write step number and description
                                if not step_cell.value:
                                    step_cell.value = str(i)
                                if not desc_cell.value:
                                    desc_cell.value = formatted_procedure
                                    
                        except Exception as e:
                            st.write(f"⚠️ Failed to add procedure step {i}: {e}")
                    
                    st.success(f"✅ Added {len(procedures)} packaging procedure steps")
                    
        except Exception as e:
            st.write(f"⚠️ Error adding packaging procedure: {e}")

    def format_procedure_text(self, procedure_text, data_map):
        """Format procedure text by replacing placeholders with actual data"""
        formatted_text = procedure_text
        
        # Common replacements
        replacements = {
            '{Inner L}': data_map.get('primary l-mm', data_map.get('length', 'XXX')),
            '{Inner W}': data_map.get('primary w-mm', data_map.get('width', 'XXX')),
            '{Inner H}': data_map.get('primary h-mm', data_map.get('height', 'XXX')),
            '{Inner Qty/Pack}': data_map.get('primary qty/pack', data_map.get('quantity', 'XXX')),
            '{Qty/Veh}': data_map.get('qty/veh', data_map.get('quantity per vehicle', '1')),
            '{Layer}': data_map.get('layer', '4'),
            '{Level}': data_map.get('level', '5'),
            '{Qty/Pack}': data_map.get('qty/pack', data_map.get('quantity', 'XXX'))
        }
        
        for placeholder, value in replacements.items():
            formatted_text = formatted_text.replace(placeholder, str(value))
        
        return formatted_text

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
