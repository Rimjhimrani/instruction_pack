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
    
    def identify_image_upload_areas(self, worksheet):
        """Identify areas in template designated for image uploads"""
        upload_areas = []
        
        try:
            # Look for cells with image-related text
            image_keywords = [
                'upload image', 'image', 'photo', 'picture', 'upload',
                'attach image', 'insert image', 'current packaging'
            ]
            
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value:
                        cell_text = str(cell.value).lower().strip()
                        
                        for keyword in image_keywords:
                            if keyword in cell_text:
                                upload_areas.append({
                                    'position': cell.coordinate,
                                    'row': cell.row,
                                    'column': cell.column,
                                    'text': cell.value,
                                    'type': self.classify_image_area(cell_text)
                                })
                                break
            
            return upload_areas
            
        except Exception as e:
            st.error(f"Error identifying image upload areas: {e}")
            return []
    
    def classify_image_area(self, text):
        """Classify the type of image area based on text"""
        text = text.lower()
        
        if 'current' in text or 'existing' in text:
            return 'current_packaging'
        elif 'primary' in text:
            return 'primary_packaging'
        elif 'secondary' in text:
            return 'secondary_packaging'
        elif 'reference' in text or 'ref' in text:
            return 'reference'
        else:
            return 'general'

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
            
            for area in image_areas:
                area_type = area['type']
                
                # Find matching uploaded image
                matching_image = None
                for img_name, img_data in uploaded_images.items():
                    # Match by area type or filename
                    if (area_type in img_name.lower() or 
                        'current' in img_name.lower() and area_type == 'current_packaging' or
                        'primary' in img_name.lower() and area_type == 'primary_packaging' or
                        'secondary' in img_name.lower() and area_type == 'secondary_packaging'):
                        matching_image = img_data
                        break
                
                # If no specific match, use first available image
                if not matching_image and uploaded_images:
                    matching_image = list(uploaded_images.values())[0]
                
                if matching_image:
                    try:
                        # Create temporary image file
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                            image_bytes = base64.b64decode(matching_image['data'])
                            tmp_img.write(image_bytes)
                            tmp_img_path = tmp_img.name
                        
                        # Create openpyxl Image object
                        img = OpenpyxlImage(tmp_img_path)
                        
                        # Resize image to fit cell area (approximate)
                        img.width = 100
                        img.height = 100
                        
                        # Add image to worksheet
                        cell_coord = f"{get_column_letter(area['column'])}{area['row']}"
                        worksheet.add_image(img, cell_coord)
                        
                        # Clean up temporary file
                        os.unlink(tmp_img_path)
                        added_images += 1
                        
                    except Exception as e:
                        st.warning(f"Could not add image to {area['position']}: {e}")
                        continue
            
            return added_images
            
        except Exception as e:
            st.error(f"Error adding images to template: {e}")
            return 0
    
    def fill_template_with_data_and_images(self, template_file, mapping_results, data_df, uploaded_images=None):
        """Fill template with mapped data and images"""
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
            
            filled_count = 0
            images_added = 0
            
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
                images_added = self.add_images_to_template(worksheet, uploaded_images, image_areas)
            
            return workbook, filled_count, images_added
            
        except Exception as e:
            st.error(f"Error filling template: {e}")
            return None, 0, 0

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
    """Main application interface"""
    # Header
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.title("🤖 Enhanced AI Template Mapper with Images")
    with col2:
        st.write(f"**Welcome, {st.session_state.name}**")
    with col3:
        if st.button("Logout"):
            st.session_state.authenticated = False
            st.session_state.user_role = None
            st.rerun()
    
    st.markdown("---")
    
    # Sidebar for navigation
    with st.sidebar:
        st.header("Navigation")
        tab_choice = st.radio(
            "Choose Operation:",
            ["🎯 Template Mapping", "📊 Analytics", "⚙️ Settings"] if st.session_state.user_role == "admin" 
            else ["🎯 Template Mapping", "📊 Analytics"]
        )
        
        st.markdown("---")
        st.header("Quick Stats")
        if st.session_state.templates:
            st.metric("Templates Processed", len(st.session_state.templates))
        else:
            st.metric("Templates Processed", 0)
    
    # Main content area
    if tab_choice == "🎯 Template Mapping":
        show_template_mapping()
    elif tab_choice == "📊 Analytics":
        show_analytics()
    elif tab_choice == "⚙️ Settings" and st.session_state.user_role == "admin":
        show_settings()

def show_template_mapping():
    """Template mapping interface"""
    st.header("Template Mapping with Image Support")
    
    # File upload section
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Upload Template")
        template_file = st.file_uploader(
            "Choose Excel template file",
            type=['xlsx', 'xls'],
            help="Upload your packaging template Excel file"
        )
    
    with col2:
        st.subheader("📋 Upload Data")
        data_file = st.file_uploader(
            "Choose data file",
            type=['xlsx', 'xls', 'csv'],
            help="Upload your data file to map to the template"
        )
    
    # Image upload section
    st.subheader("🖼️ Upload Images (Optional)")
    uploaded_images = {}
    
    image_files = st.file_uploader(
        "Choose image files",
        type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
        accept_multiple_files=True,
        help="Upload images to be inserted into the template"
    )
    
    if image_files:
        for img_file in image_files:
            try:
                # Convert to base64 for storage
                img_data = img_file.read()
                img_b64 = base64.b64encode(img_data).decode()
                
                # Get image info
                pil_img = Image.open(io.BytesIO(img_data))
                
                uploaded_images[img_file.name] = {
                    'data': img_b64,
                    'format': pil_img.format,
                    'size': pil_img.size,
                    'filename': img_file.name
                }
                
                # Show preview
                st.image(pil_img, caption=img_file.name, width=150)
                
            except Exception as e:
                st.error(f"Error processing image {img_file.name}: {e}")
    
    # Processing section
    if template_file and data_file:
        try:
            # Load data
            if data_file.name.endswith('.csv'):
                data_df = pd.read_csv(data_file)
            else:
                data_df = pd.read_excel(data_file)
            
            st.success(f"✅ Data loaded: {len(data_df)} rows, {len(data_df.columns)} columns")
            
            # Show data preview
            with st.expander("📊 Data Preview"):
                st.dataframe(data_df.head())
            
            # Analyze template
            with st.spinner("🔍 Analyzing template and finding mappable fields..."):
                # Save template temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_template:
                    tmp_template.write(template_file.read())
                    template_path = tmp_template.name
                
                # Find template fields and image areas
                template_fields, image_areas = st.session_state.enhanced_mapper.find_template_fields_with_context_and_images(template_path)
                
                # Perform mapping
                mapping_results = st.session_state.enhanced_mapper.map_data_with_section_context(template_fields, data_df)
            
            st.success(f"✅ Found {len(template_fields)} mappable fields and {len(image_areas)} image areas")
            
            # Show mapping results
            st.subheader("🎯 Mapping Results")
            
            # Create mapping summary
            mapped_count = sum(1 for m in mapping_results.values() if m['is_mappable'])
            unmapped_count = len(mapping_results) - mapped_count
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Fields", len(mapping_results))
            with col2:
                st.metric("Mapped Fields", mapped_count)
            with col3:
                st.metric("Unmapped Fields", unmapped_count)
            
            # Detailed mapping table
            with st.expander("📋 Detailed Mapping Results"):
                mapping_df = pd.DataFrame([
                    {
                        'Template Field': m['template_field'],
                        'Data Column': m['data_column'] if m['data_column'] else 'No Match',
                        'Section': m['section_context'] if m['section_context'] else 'General',
                        'Similarity': f"{m['similarity']:.2f}" if m['similarity'] > 0 else "0.00",
                        'Status': '✅ Mapped' if m['is_mappable'] else '❌ Unmapped'
                    }
                    for coord, m in mapping_results.items()
                ])
                st.dataframe(mapping_df, use_container_width=True)
            
            # Image areas information
            if image_areas:
                with st.expander("🖼️ Image Upload Areas Found"):
                    image_df = pd.DataFrame([
                        {
                            'Position': area['position'],
                            'Type': area['type'].replace('_', ' ').title(),
                            'Description': area['text']
                        }
                        for area in image_areas
                    ])
                    st.dataframe(image_df, use_container_width=True)
            
            # Generate filled template
            if st.button("🚀 Generate Filled Template", type="primary"):
                with st.spinner("📝 Filling template with data and images..."):
                    try:
                        filled_workbook, filled_count, images_added = st.session_state.enhanced_mapper.fill_template_with_data_and_images(
                            template_path, mapping_results, data_df, uploaded_images
                        )
                        
                        if filled_workbook:
                            # Save filled template
                            output_buffer = io.BytesIO()
                            filled_workbook.save(output_buffer)
                            filled_workbook.close()
                            output_buffer.seek(0)
                            
                            # Success metrics
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Fields Filled", filled_count)
                            with col2:
                                st.metric("Images Added", images_added)
                            with col3:
                                st.metric("Success Rate", f"{(filled_count/len(mapping_results)*100):.1f}%")
                            
                            # Download button
                            st.download_button(
                                label="📥 Download Filled Template",
                                data=output_buffer.getvalue(),
                                file_name=f"filled_template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                            st.success("✅ Template filled successfully!")
                            
                            # Store in session for analytics
                            template_id = f"template_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                            st.session_state.templates[template_id] = {
                                'timestamp': datetime.now(),
                                'fields_total': len(mapping_results),
                                'fields_mapped': mapped_count,
                                'fields_filled': filled_count,
                                'images_added': images_added,
                                'success_rate': filled_count/len(mapping_results)*100 if mapping_results else 0
                            }
                        
                    except Exception as e:
                        st.error(f"❌ Error generating template: {e}")
                    finally:
                        # Clean up temporary file
                        try:
                            os.unlink(template_path)
                        except:
                            pass
        
        except Exception as e:
            st.error(f"❌ Error processing files: {e}")

def show_analytics():
    """Analytics dashboard"""
    st.header("📊 Analytics Dashboard")
    
    if not st.session_state.templates:
        st.info("No templates processed yet. Start by mapping some templates!")
        return
    
    # Overall metrics
    templates = st.session_state.templates
    total_templates = len(templates)
    avg_success_rate = np.mean([t['success_rate'] for t in templates.values()])
    total_fields = sum([t['fields_total'] for t in templates.values()])
    total_mapped = sum([t['fields_mapped'] for t in templates.values()])
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Templates", total_templates)
    with col2:
        st.metric("Avg Success Rate", f"{avg_success_rate:.1f}%")
    with col3:
        st.metric("Total Fields", total_fields)
    with col4:
        st.metric("Fields Mapped", total_mapped)
    
    # Recent activity
    st.subheader("Recent Activity")
    recent_df = pd.DataFrame([
        {
            'Template ID': tid,
            'Timestamp': info['timestamp'].strftime('%Y-%m-%d %H:%M'),
            'Fields Total': info['fields_total'],
            'Fields Mapped': info['fields_mapped'],
            'Success Rate': f"{info['success_rate']:.1f}%"
        }
        for tid, info in sorted(templates.items(), key=lambda x: x[1]['timestamp'], reverse=True)
    ])
    st.dataframe(recent_df, use_container_width=True)

def show_settings():
    """Settings panel for admin users"""
    st.header("⚙️ Settings")
    
    st.subheader("Mapping Configuration")
    
    # Similarity threshold
    current_threshold = st.session_state.enhanced_mapper.similarity_threshold
    new_threshold = st.slider(
        "Similarity Threshold",
        min_value=0.1,
        max_value=1.0,
        value=current_threshold,
        step=0.05,
        help="Higher values require closer matches"
    )
    
    if new_threshold != current_threshold:
        st.session_state.enhanced_mapper.similarity_threshold = new_threshold
        st.success("Threshold updated!")
    
    # Clear data
    st.subheader("Data Management")
    if st.button("🗑️ Clear All Template Data", type="secondary"):
        st.session_state.templates = {}
        st.success("All template data cleared!")
    
    # System info
    st.subheader("System Information")
    st.write(f"**NLTK Available:** {'✅ Yes' if NLTK_READY else '❌ No'}")
    st.write(f"**Advanced NLP:** {'✅ Enabled' if ADVANCED_NLP else '❌ Disabled'}")
    st.write(f"**Templates in Memory:** {len(st.session_state.templates)}")

# Main application flow
def main():
    if not st.session_state.authenticated:
        show_login()
    else:
        show_main_app()

if __name__ == "__main__":
    main()
