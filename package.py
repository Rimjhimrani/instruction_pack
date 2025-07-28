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
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
import re
import io
import tempfile
import shutil
from pathlib import Path
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
    """Handles image extraction from Excel files"""
    
    def __init__(self):
        self.supported_formats = ['.png', '.jpg', '.jpeg', '.gif', '.bmp']
    
    def extract_images_from_excel(self, excel_file_path):
        """Extract all images from Excel file"""
        try:
            images = {}
            workbook = openpyxl.load_workbook(excel_file_path)
            
            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
                sheet_images = {}
                
                # Extract images from worksheet
                if hasattr(worksheet, '_images'):
                    for idx, img in enumerate(worksheet._images):
                        try:
                            # Get image data
                            image_data = img._data()
                            
                            # Create PIL Image
                            pil_image = Image.open(io.BytesIO(image_data))
                            
                            # Get image position (approximate)
                            anchor = img.anchor
                            if hasattr(anchor, '_from'):
                                col = anchor._from.col
                                row = anchor._from.row
                                position = f"{get_column_letter(col + 1)}{row + 1}"
                            else:
                                position = f"Image_{idx + 1}"
                            
                            # Convert to base64 for storage
                            buffered = io.BytesIO()
                            pil_image.save(buffered, format="PNG")
                            img_str = base64.b64encode(buffered.getvalue()).decode()
                            
                            sheet_images[position] = {
                                'data': img_str,
                                'format': 'PNG',
                                'size': pil_image.size,
                                'position': position
                            }
                            
                        except Exception as e:
                            print(f"Error extracting image {idx}: {e}")
                            continue
                
                if sheet_images:
                    images[sheet_name] = sheet_images
            
            workbook.close()
            return images
            
        except Exception as e:
            st.error(f"Error extracting images: {e}")
            return {}
    
    def identify_image_upload_areas_precise(self, worksheet):
        """Identify precise image upload positions based on column headers"""
        upload_areas = []
        
        try:
            # Look for specific column headers and their positions
            header_patterns = {
                'primary_packaging': ['primary packaging', 'primary', 'internal packaging'],
                'secondary_packaging': ['secondary packaging', 'secondary', 'outer packaging', 'external packaging'],
                'current_packaging': ['current packaging', 'existing packaging', 'current']
            }
            
            # Search for headers in first few rows
            for row_idx in range(1, min(10, worksheet.max_row + 1)):
                for col_idx in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    if cell.value:
                        cell_text = str(cell.value).lower().strip()
                        
                        for area_type, patterns in header_patterns.items():
                            for pattern in patterns:
                                if pattern in cell_text:
                                    # Found a header, place image below it
                                    image_row = row_idx + 1
                                    upload_areas.append({
                                        'type': area_type,
                                        'position': f"{get_column_letter(col_idx)}{image_row}",
                                        'row': image_row,
                                        'column': col_idx,
                                        'header_text': cell.value,
                                        'header_position': cell.coordinate
                                    })
                                    break
            
            return upload_areas
            
        except Exception as e:
            st.error(f"Error identifying image upload areas: {e}")
            return []

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
            
            # Find precise image upload areas
            image_areas = self.image_extractor.identify_image_upload_areas_precise(worksheet)
            
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
    
    def add_images_to_template_precise(self, worksheet, uploaded_images, image_areas):
        """Add uploaded images to template with precise positioning and sizing"""
        try:
            added_images = 0
            temp_image_paths = []
            used_images = set()
            
            # Process each image area
            for area in image_areas:
                area_type = area['type']
                position = area['position']
                
                # Find matching image for this area
                matching_image = None
                matching_label = None
                
                # First, try to find exact type match
                for label, img_data in uploaded_images.items():
                    if label in used_images:
                        continue
                    label_lower = label.lower()
                    
                    # Check for exact type matches
                    if area_type == 'primary_packaging' and 'primary' in label_lower:
                        matching_image = img_data
                        matching_label = label
                        break
                    elif area_type == 'secondary_packaging' and 'secondary' in label_lower:
                        matching_image = img_data
                        matching_label = label
                        break
                    elif area_type == 'current_packaging' and 'current' in label_lower:
                        matching_image = img_data
                        matching_label = label
                        break
                
                # If no exact match, use any available image
                if not matching_image:
                    for label, img_data in uploaded_images.items():
                        if label not in used_images:
                            matching_image = img_data
                            matching_label = label
                            break
                
                # Add image if found
                if matching_image:
                    try:
                        # Create temporary image file
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                            image_bytes = base64.b64decode(matching_image['data'])
                            tmp_img.write(image_bytes)
                            tmp_img_path = tmp_img.name
                        
                        # Create openpyxl image object
                        img = OpenpyxlImage(tmp_img_path)
                        
                        # Set precise dimensions based on area type
                        if area_type == 'current_packaging':
                            # Current packaging: 8.3 cm x 8.3 cm
                            img.width = int(8.3 * 28.35)  # Convert cm to points (1 cm ≈ 28.35 points)
                            img.height = int(8.3 * 28.35)
                        else:
                            # Primary and Secondary packaging: 4.3 cm x 4.3 cm
                            img.width = int(4.3 * 28.35)
                            img.height = int(4.3 * 28.35)
                        
                        # Add image to the precise position
                        worksheet.add_image(img, position)
                        
                        temp_image_paths.append(tmp_img_path)
                        used_images.add(matching_label)
                        added_images += 1
                        
                        st.success(f"✅ Added {area_type} image at {position}")
                        
                    except Exception as e:
                        st.warning(f"❌ Could not add {area_type} image to {position}: {e}")
                        continue
                else:
                    st.info(f"ℹ️ No matching image found for {area_type} at {position}")
            
            return added_images, temp_image_paths
            
        except Exception as e:
            st.error(f"Error adding images to template: {e}")
            return 0, []
    
    def fill_template_with_data_and_images(self, template_file, mapping_results, data_df, uploaded_images=None):
        """Fill template with mapped data and images"""
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
            
            filled_count = 0
            images_added = 0
            temp_image_paths = []
            
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
                # Find precise image upload areas
                _, image_areas = self.find_template_fields_with_context_and_images(template_file)
                images_added, temp_image_paths = self.add_images_to_template_precise(worksheet, uploaded_images, image_areas)
                
            return workbook, filled_count, images_added, temp_image_paths
            
        except Exception as e:
            st.error(f"Error filling template: {e}")
            return None, 0, 0, []

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'user_role' not in st.session_state:
    st.session_state.user_role = None
if 'templates' not in st.session_state:
    st.session_state.templates = {}
if 'enhanced_mapper' not in st.session_state:
    st.session_state.enhanced_mapper = EnhancedTemplateMapperWithImages()
if 'saved_template_path' not in st.session_state:
    st.session_state.saved_template_path = None

# User management functions
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(password, hashed):
    return hash_password(password) == hashed

DEFAULT_USERS = {
    "admin": {
        "password": hash_password("admin123"),
        "role": "admin"
    },
    "user": {
        "password": hash_password("user123"),
        "role": "user"
    }
}

def authenticate_user(username, password):
    if username in DEFAULT_USERS:
        if verify_password(password, DEFAULT_USERS[username]["password"]):
            return DEFAULT_USERS[username]["role"]
    return None

def login_interface():
    st.title("🔐 AI Template Mapper Login")
    
    with st.form("login_form"):
        username = st.text_input("Username", placeholder="Enter username")
        password = st.text_input("Password", type="password", placeholder="Enter password")
        submit_button = st.form_submit_button("Login")
        
        if submit_button:
            if username and password:
                role = authenticate_user(username, password)
                if role:
                    st.session_state.authenticated = True
                    st.session_state.user_role = role
                    st.session_state.username = username
                    st.success(f"Welcome {username}! Logged in as {role}")
                    st.rerun()
                else:
                    st.error("Invalid username or password")
            else:
                st.warning("Please enter both username and password")
    
    with st.expander("Demo Credentials"):
        st.info("**Admin:** username: `admin`, password: `admin123`")
        st.info("**User:** username: `user`, password: `user123`")

def logout():
    st.session_state.authenticated = False
    st.session_state.user_role = None
    st.session_state.username = None
    st.rerun()

def main_interface():
    # Header with logout
    col1, col2 = st.columns([6, 1])
    with col1:
        st.title("🤖 AI Template Mapper - Enhanced with Images")
        st.markdown("**Intelligent Excel template mapping with image support**")
    with col2:
        if st.button("Logout", type="secondary"):
            logout()
    
    # Sidebar
    with st.sidebar:
        st.header("📊 Dashboard")
        st.info(f"👤 Logged in as: **{st.session_state.username}** ({st.session_state.user_role})")
        
        st.header("⚙️ Settings")
        similarity_threshold = st.slider(
            "Similarity Threshold", 
            min_value=0.1, 
            max_value=0.9, 
            value=0.3, 
            step=0.05,
            help="Minimum similarity score for field matching"
        )
        st.session_state.enhanced_mapper.similarity_threshold = similarity_threshold
        
        st.header("📈 Statistics")
        if 'last_mapping_stats' in st.session_state:
            stats = st.session_state.last_mapping_stats
            st.metric("Fields Mapped", stats.get('mapped_fields', 0))
            st.metric("Images Added", stats.get('images_added', 0))
            st.metric("Success Rate", f"{stats.get('success_rate', 0):.1%}")
    
    # Main tabs
    tab1, tab2, tab3, tab4 = st.tabs(["🔧 Template Mapping", "📋 Template Analysis", "🖼️ Image Management", "📊 Batch Processing"])
    
    with tab1:
        template_mapping_interface()
    
    with tab2:
        template_analysis_interface()
    
    with tab3:
        image_management_interface()
    
    with tab4:
        if st.session_state.user_role == "admin":
            batch_processing_interface()
        else:
            st.warning("🔒 Admin access required for batch processing")

def template_mapping_interface():
    st.header("🔧 Template Mapping")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Upload Template")
        template_file = st.file_uploader(
            "Choose Excel template file",
            type=['xlsx', 'xls'],
            help="Upload your Excel template with fields to be mapped"
        )
        
        if template_file:
            st.success("✅ Template uploaded successfully!")
            
            # Preview template structure
            with st.expander("🔍 Template Preview"):
                try:
                    preview_df = pd.read_excel(template_file, nrows=10)
                    st.dataframe(preview_df, use_container_width=True)
                except Exception as e:
                    st.error(f"Error previewing template: {e}")
    
    with col2:
        st.subheader("📊 Upload Data")
        data_file = st.file_uploader(
            "Choose data file",
            type=['xlsx', 'xls', 'csv'],
            help="Upload your data file with values to map"
        )
        
        if data_file:
            st.success("✅ Data file uploaded successfully!")
            
            # Preview data
            with st.expander("🔍 Data Preview"):
                try:
                    if data_file.name.endswith('.csv'):
                        preview_df = pd.read_csv(data_file, nrows=10)
                    else:
                        preview_df = pd.read_excel(data_file, nrows=10)
                    st.dataframe(preview_df, use_container_width=True)
                except Exception as e:
                    st.error(f"Error previewing data: {e}")
    
    # Image upload section
    st.subheader("🖼️ Upload Images (Optional)")
    uploaded_images = {}
    
    col_img1, col_img2, col_img3 = st.columns(3)
    
    with col_img1:
        st.write("**Primary Packaging Image**")
        primary_img = st.file_uploader(
            "Primary packaging",
            type=['png', 'jpg', 'jpeg'],
            key="primary_img"
        )
        if primary_img:
            # Convert to base64 for storage
            img_bytes = primary_img.read()
            img_b64 = base64.b64encode(img_bytes).decode()
            uploaded_images["primary_packaging"] = {
                'data': img_b64,
                'format': primary_img.type.split('/')[-1].upper()
            }
            st.image(primary_img, caption="Primary Packaging", width=150)
    
    with col_img2:
        st.write("**Secondary Packaging Image**")
        secondary_img = st.file_uploader(
            "Secondary packaging",
            type=['png', 'jpg', 'jpeg'],
            key="secondary_img"
        )
        if secondary_img:
            img_bytes = secondary_img.read()
            img_b64 = base64.b64encode(img_bytes).decode()
            uploaded_images["secondary_packaging"] = {
                'data': img_b64,
                'format': secondary_img.type.split('/')[-1].upper()
            }
            st.image(secondary_img, caption="Secondary Packaging", width=150)
    
    with col_img3:
        st.write("**Current Packaging Image**")
        current_img = st.file_uploader(
            "Current packaging",
            type=['png', 'jpg', 'jpeg'],
            key="current_img"
        )
        if current_img:
            img_bytes = current_img.read()
            img_b64 = base64.b64encode(img_bytes).decode()
            uploaded_images["current_packaging"] = {
                'data': img_b64,
                'format': current_img.type.split('/')[-1].upper()
            }
            st.image(current_img, caption="Current Packaging", width=150)
    
    # Process mapping
    if template_file and data_file:
        if st.button("🚀 Process Mapping", type="primary"):
            with st.spinner("Processing template mapping..."):
                try:
                    # Save uploaded files temporarily
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_template:
                        tmp_template.write(template_file.read())
                        tmp_template_path = tmp_template.name
                    
                    # Read data
                    if data_file.name.endswith('.csv'):
                        data_df = pd.read_csv(data_file)
                    else:
                        data_df = pd.read_excel(data_file)
                    
                    # Find template fields and image areas
                    template_fields, image_areas = st.session_state.enhanced_mapper.find_template_fields_with_context_and_images(tmp_template_path)
                    
                    # Map data
                    mapping_results = st.session_state.enhanced_mapper.map_data_with_section_context(template_fields, data_df)
                    
                    # Display mapping results
                    st.subheader("📋 Mapping Results")
                    
                    mapping_df = []
                    mapped_count = 0
                    for coord, result in mapping_results.items():
                        status = "✅ Mapped" if result['is_mappable'] else "❌ No Match"
                        if result['is_mappable']:
                            mapped_count += 1
                        
                        mapping_df.append({
                            'Position': coord,
                            'Template Field': result['template_field'],
                            'Data Column': result['data_column'] or 'No match',
                            'Similarity': f"{result['similarity']:.2f}" if result['similarity'] > 0 else "0.00",
                            'Section': result['section_context'] or 'General',
                            'Status': status
                        })
                    
                    mapping_display_df = pd.DataFrame(mapping_df)
                    st.dataframe(mapping_display_df, use_container_width=True)
                    
                    # Statistics
                    total_fields = len(mapping_results)
                    success_rate = mapped_count / total_fields if total_fields > 0 else 0
                    
                    col_stat1, col_stat2, col_stat3 = st.columns(3)
                    with col_stat1:
                        st.metric("Total Fields", total_fields)
                    with col_stat2:
                        st.metric("Mapped Fields", mapped_count)
                    with col_stat3:
                        st.metric("Success Rate", f"{success_rate:.1%}")
                    
                    # Generate filled template
                    if st.button("📄 Generate Filled Template", type="secondary"):
                        with st.spinner("Generating filled template..."):
                            try:
                                filled_workbook, filled_count, images_added, temp_paths = st.session_state.enhanced_mapper.fill_template_with_data_and_images(
                                    tmp_template_path, mapping_results, data_df, uploaded_images if uploaded_images else None
                                )
                                
                                if filled_workbook:
                                    # Save to temporary file
                                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_output:
                                        filled_workbook.save(tmp_output.name)
                                        st.session_state.saved_template_path = tmp_output.name
                                    
                                    # Store statistics
                                    st.session_state.last_mapping_stats = {
                                        'mapped_fields': filled_count,
                                        'images_added': images_added,
                                        'success_rate': success_rate
                                    }
                                    
                                    st.success(f"✅ Template filled successfully! {filled_count} fields filled, {images_added} images added.")
                                    
                                    # Download button
                                    with open(st.session_state.saved_template_path, 'rb') as file:
                                        st.download_button(
                                            label="📥 Download Filled Template",
                                            data=file.read(),
                                            file_name=f"filled_template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        )
                                    
                                    # Clean up temporary image files
                                    for temp_path in temp_paths:
                                        try:
                                            os.unlink(temp_path)
                                        except:
                                            pass
                                
                            except Exception as e:
                                st.error(f"Error generating filled template: {e}")
                    
                    # Clean up temporary template file
                    try:
                        os.unlink(tmp_template_path)
                    except:
                        pass
                        
                except Exception as e:
                    st.error(f"Error processing mapping: {e}")

def template_analysis_interface():
    st.header("📋 Template Analysis")
    
    uploaded_file = st.file_uploader(
        "Upload Excel template for analysis",
        type=['xlsx', 'xls'],
        key="analysis_file"
    )
    
    if uploaded_file:
        try:
            # Save temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(uploaded_file.read())
                tmp_file_path = tmp_file.name
            
            # Analyze template
            template_fields, image_areas = st.session_state.enhanced_mapper.find_template_fields_with_context_and_images(tmp_file_path)
            
            # Display analysis results
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🔍 Detected Fields")
                
                if template_fields:
                    fields_df = []
                    for coord, field in template_fields.items():
                        fields_df.append({
                            'Position': coord,
                            'Field Name': field['value'],
                            'Section': field.get('section_context', 'General'),
                            'Row': field['row'],
                            'Column': field['column'],
                            'Merged': 'Yes' if field['merged_range'] else 'No'
                        })
                    
                    fields_display_df = pd.DataFrame(fields_df)
                    st.dataframe(fields_display_df, use_container_width=True)
                    
                    # Section breakdown
                    st.subheader("📊 Section Breakdown")
                    section_counts = fields_display_df['Section'].value_counts()
                    st.bar_chart(section_counts)
                else:
                    st.warning("No mappable fields detected in template")
            
            with col2:
                st.subheader("🖼️ Image Upload Areas")
                
                if image_areas:
                    image_df = []
                    for area in image_areas:
                        image_df.append({
                            'Type': area['type'].replace('_', ' ').title(),
                            'Position': area['position'],
                            'Header Text': area.get('header_text', ''),
                            'Header Position': area.get('header_position', '')
                        })
                    
                    image_display_df = pd.DataFrame(image_df)
                    st.dataframe(image_display_df, use_container_width=True)
                    
                    st.info(f"Found {len(image_areas)} image upload areas")
                else:
                    st.warning("No image upload areas detected")
                
                # Extract existing images
                st.subheader("📷 Existing Images")
                existing_images = st.session_state.enhanced_mapper.image_extractor.extract_images_from_excel(tmp_file_path)
                
                if existing_images:
                    for sheet_name, sheet_images in existing_images.items():
                        st.write(f"**Sheet: {sheet_name}**")
                        for pos, img_info in sheet_images.items():
                            col_img1, col_img2 = st.columns([1, 2])
                            with col_img1:
                                # Display image
                                img_data = base64.b64decode(img_info['data'])
                                st.image(img_data, caption=f"Position: {pos}", width=100)
                            with col_img2:
                                st.write(f"**Position:** {pos}")
                                st.write(f"**Size:** {img_info['size'][0]}x{img_info['size'][1]}")
                                st.write(f"**Format:** {img_info['format']}")
                else:
                    st.info("No existing images found in template")
            
            # Clean up
            try:
                os.unlink(tmp_file_path)
            except:
                pass
                
        except Exception as e:
            st.error(f"Error analyzing template: {e}")

def image_management_interface():
    st.header("🖼️ Image Management")
    
    st.subheader("📤 Bulk Image Upload")
    
    # Multiple image upload
    uploaded_files = st.file_uploader(
        "Upload multiple images",
        type=['png', 'jpg', 'jpeg'],
        accept_multiple_files=True,
        help="Upload multiple images for batch processing"
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} images uploaded successfully!")
        
        # Display uploaded images
        cols = st.columns(min(len(uploaded_files), 4))
        for idx, uploaded_file in enumerate(uploaded_files):
            with cols[idx % 4]:
                st.image(uploaded_file, caption=uploaded_file.name, width=150)
                
                # Image info
                img = Image.open(uploaded_file)
                st.write(f"**Size:** {img.size[0]}x{img.size[1]}")
                st.write(f"**Format:** {img.format}")
                st.write(f"**Mode:** {img.mode}")
        
        # Image processing options
        st.subheader("⚙️ Processing Options")
        
        col_proc1, col_proc2 = st.columns(2)
        
        with col_proc1:
            resize_images = st.checkbox("Resize images", value=True)
            if resize_images:
                resize_width = st.number_input("Width (px)", value=400, min_value=50, max_value=2000)
                resize_height = st.number_input("Height (px)", value=400, min_value=50, max_value=2000)
        
        with col_proc2:
            convert_format = st.selectbox("Convert to format", ["Keep original", "PNG", "JPEG"])
            compress_quality = st.slider("JPEG Quality", 10, 100, 85) if convert_format == "JPEG" else None
        
        # Process images
        if st.button("🔄 Process Images", type="primary"):
            processed_images = {}
            
            for uploaded_file in uploaded_files:
                try:
                    img = Image.open(uploaded_file)
                    
                    # Resize if requested
                    if resize_images:
                        img = img.resize((resize_width, resize_height), Image.Resampling.LANCZOS)
                    
                    # Convert format if requested
                    if convert_format != "Keep original":
                        if convert_format == "JPEG" and img.mode in ("RGBA", "P"):
                            img = img.convert("RGB")
                    
                    # Save to base64
                    buffered = io.BytesIO()
                    save_format = convert_format if convert_format != "Keep original" else img.format
                    save_kwargs = {}
                    
                    if save_format == "JPEG" and compress_quality:
                        save_kwargs['quality'] = compress_quality
                    
                    img.save(buffered, format=save_format, **save_kwargs)
                    img_b64 = base64.b64encode(buffered.getvalue()).decode()
                    
                    processed_images[uploaded_file.name] = {
                        'data': img_b64,
                        'format': save_format,
                        'size': img.size
                    }
                    
                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: {e}")
            
            st.success(f"✅ Processed {len(processed_images)} images successfully!")
            
            # Store in session state for use in mapping
            st.session_state.processed_images = processed_images

def batch_processing_interface():
    st.header("📊 Batch Processing")
    st.info("🔒 Admin feature - Process multiple templates at once")
    
    # Template upload
    st.subheader("📄 Upload Templates")
    template_files = st.file_uploader(
        "Upload multiple template files",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        key="batch_templates"
    )
    
    # Data upload
    st.subheader("📊 Upload Data Files")
    data_files = st.file_uploader(
        "Upload corresponding data files",
        type=['xlsx', 'xls', 'csv'],
        accept_multiple_files=True,
        key="batch_data"
    )
    
    if template_files and data_files:
        if len(template_files) != len(data_files):
            st.warning("⚠️ Number of template files must match number of data files")
        else:
            st.success(f"✅ Ready to process {len(template_files)} template-data pairs")
            
            # Processing options
            st.subheader("⚙️ Batch Processing Options")
            
            col_batch1, col_batch2 = st.columns(2)
            
            with col_batch1:
                output_format = st.selectbox("Output format", ["Individual Files", "ZIP Archive"])
                include_summary = st.checkbox("Include processing summary", value=True)
            
            with col_batch2:
                auto_download = st.checkbox("Auto-download results", value=True)
                parallel_processing = st.checkbox("Parallel processing", value=True, help="Process multiple files simultaneously")
            
            # Process batch
            if st.button("🚀 Start Batch Processing", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                results = []
                temp_files = []
                
                for idx, (template_file, data_file) in enumerate(zip(template_files, data_files)):
                    try:
                        status_text.text(f"Processing {template_file.name}...")
                        progress_bar.progress((idx + 1) / len(template_files))
                        
                        # Save template temporarily
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_template:
                            tmp_template.write(template_file.read())
                            tmp_template_path = tmp_template.name
                        
                        # Read data
                        if data_file.name.endswith('.csv'):
                            data_df = pd.read_csv(data_file)
                        else:
                            data_df = pd.read_excel(data_file)
                        
                        # Process mapping
                        template_fields, _ = st.session_state.enhanced_mapper.find_template_fields_with_context_and_images(tmp_template_path)
                        mapping_results = st.session_state.enhanced_mapper.map_data_with_section_context(template_fields, data_df)
                        
                        # Fill template
                        filled_workbook, filled_count, images_added, temp_paths = st.session_state.enhanced_mapper.fill_template_with_data_and_images(
                            tmp_template_path, mapping_results, data_df
                        )
                        
                        if filled_workbook:
                            # Save result
                            output_filename = f"filled_{template_file.name}"
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_output:
                                filled_workbook.save(tmp_output.name)
                                temp_files.append((tmp_output.name, output_filename))
                            
                            # Record results
                            total_fields = len(template_fields)
                            success_rate = filled_count / total_fields if total_fields > 0 else 0
                            
                            results.append({
                                'Template': template_file.name,
                                'Data File': data_file.name,
                                'Fields Mapped': filled_count,
                                'Total Fields': total_fields,
                                'Success Rate': f"{success_rate:.1%}",
                                'Images Added': images_added,
                                'Status': 'Success'
                            })
                        
                        # Clean up template temp file
                        os.unlink(tmp_template_path)
                        
                        # Clean up image temp files
                        for temp_path in temp_paths:
                            try:
                                os.unlink(temp_path)
                            except:
                                pass
                    
                    except Exception as e:
                        results.append({
                            'Template': template_file.name,
                            'Data File': data_file.name,
                            'Fields Mapped': 0,
                            'Total Fields': 0,
                            'Success Rate': '0%',
                            'Images Added': 0,
                            'Status': f'Error: {str(e)[:50]}...'
                        })
                
                progress_bar.progress(1.0)
                status_text.text("✅ Batch processing completed!")
                
                # Display results
                st.subheader("📊 Processing Results")
                results_df = pd.DataFrame(results)
                st.dataframe(results_df, use_container_width=True)
                
                # Summary statistics
                if include_summary:
                    successful_files = len([r for r in results if r['Status'] == 'Success'])
                    total_mapped = sum([int(r['Fields Mapped']) for r in results if r['Status'] == 'Success'])
                    total_images = sum([int(r['Images Added']) for r in results if r['Status'] == 'Success'])
                    
                    col_sum1, col_sum2, col_sum3, col_sum4 = st.columns(4)
                    with col_sum1:
                        st.metric("Successful Files", successful_files)
                    with col_sum2:
                        st.metric("Total Fields Mapped", total_mapped)
                    with col_sum3:
                        st.metric("Total Images Added", total_images)
                    with col_sum4:
                        st.metric("Overall Success Rate", f"{successful_files/len(results):.1%}")
                
                # Provide downloads
                if temp_files:
                    if output_format == "ZIP Archive":
                        # Create ZIP file
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as tmp_zip:
                            with zipfile.ZipFile(tmp_zip.name, 'w') as zipf:
                                for temp_path, filename in temp_files:
                                    zipf.write(temp_path, filename)
                            
                            # Download ZIP
                            with open(tmp_zip.name, 'rb') as file:
                                st.download_button(
                                    label="📥 Download All Results (ZIP)",
                                    data=file.read(),
                                    file_name=f"batch_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                                    mime="application/zip"
                                )
                            
                            # Clean up ZIP file
                            os.unlink(tmp_zip.name)
                    
                    else:
                        # Individual downloads
                        st.subheader("📥 Individual Downloads")
                        for temp_path, filename in temp_files:
                            with open(temp_path, 'rb') as file:
                                st.download_button(
                                    label=f"📄 {filename}",
                                    data=file.read(),
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"download_{filename}"
                                )
                    
                    # Clean up temporary files
                    for temp_path, _ in temp_files:
                        try:
                            os.unlink(temp_path)
                        except:
                            pass

# Main application logic
def main():
    try:
        if not st.session_state.authenticated:
            login_interface()
        else:
            main_interface()
    except Exception as e:
        st.error(f"Application error: {e}")
        st.info("Please refresh the page and try again.")

if __name__ == "__main__":
    main()
