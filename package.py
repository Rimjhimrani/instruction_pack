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
from PIL import Image as PILImage
import zipfile

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

class EnhancedTemplateMapperWithImages:
    def __init__(self):
        self.similarity_threshold = 0.3
        self.stop_words = {
            'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from',
            'has', 'he', 'in', 'is', 'it', 'its', 'of', 'on', 'that', 'the',
            'to', 'was', 'will', 'with', 'or', 'but', 'not', 'this', 'have',
            'had', 'what', 'when', 'where', 'who', 'which', 'why', 'how'
        }
        
        # Image placeholder patterns
        self.image_patterns = [
            r'upload\s+image', r'insert\s+image', r'add\s+image',
            r'image\s+here', r'photo\s+here', r'picture\s+here',
            r'upload\s+photo', r'insert\s+photo', r'add\s+photo',
            r'reference\s+image', r'primary\s+packaging', r'secondary\s+packaging',
            r'current\s+packaging', r'approved\s+by', r'received\s+by'
        ]
        
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
    
    def extract_images_from_excel(self, excel_file_path):
        """Extract images from Excel file"""
        images = []
        try:
            # Open Excel file as ZIP to access media files
            with zipfile.ZipFile(excel_file_path, 'r') as zip_file:
                # Look for media files in xl/media/ directory
                media_files = [f for f in zip_file.namelist() if f.startswith('xl/media/')]
                
                for media_file in media_files:
                    try:
                        # Extract image data
                        image_data = zip_file.read(media_file)
                        
                        # Get file extension
                        file_ext = os.path.splitext(media_file)[1].lower()
                        
                        # Create PIL Image to validate and get info
                        image = PILImage.open(io.BytesIO(image_data))
                        
                        images.append({
                            'filename': os.path.basename(media_file),
                            'data': image_data,
                            'format': image.format,
                            'size': image.size,
                            'mode': image.mode,
                            'extension': file_ext
                        })
                        
                        print(f"Extracted image: {media_file} - {image.size} - {image.format}")
                        
                    except Exception as e:
                        print(f"Error extracting image {media_file}: {e}")
                        continue
            
            # Also try to extract using openpyxl (for embedded images)
            try:
                workbook = openpyxl.load_workbook(excel_file_path)
                worksheet = workbook.active
                
                if hasattr(worksheet, '_images'):
                    for img in worksheet._images:
                        try:
                            # Get image data
                            img_data = img._data()
                            pil_img = PILImage.open(io.BytesIO(img_data))
                            
                            images.append({
                                'filename': f'embedded_image_{len(images)}.{pil_img.format.lower()}',
                                'data': img_data,
                                'format': pil_img.format,
                                'size': pil_img.size,
                                'mode': pil_img.mode,
                                'extension': f'.{pil_img.format.lower()}',
                                'anchor': getattr(img, 'anchor', None)
                            })
                            
                        except Exception as e:
                            print(f"Error extracting embedded image: {e}")
                            continue
                
                workbook.close()
                
            except Exception as e:
                print(f"Error using openpyxl for image extraction: {e}")
        
        except Exception as e:
            st.error(f"Error extracting images from Excel: {e}")
        
        return images
    
    def find_image_placeholders(self, worksheet):
        """Find cells that contain image upload placeholders"""
        image_placeholders = []
        
        try:
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value:
                        cell_text = str(cell.value).lower().strip()
                        
                        # Check against image patterns
                        for pattern in self.image_patterns:
                            if re.search(pattern, cell_text):
                                image_placeholders.append({
                                    'cell': cell.coordinate,
                                    'row': cell.row,
                                    'column': cell.column,
                                    'text': cell.value,
                                    'pattern_matched': pattern
                                })
                                print(f"Found image placeholder: {cell.coordinate} - '{cell.value}'")
                                break
        
        except Exception as e:
            st.error(f"Error finding image placeholders: {e}")
        
        return image_placeholders
    
    def insert_image_into_template(self, worksheet, placeholder, image_data, image_index):
        """Insert image into worksheet at placeholder location"""
        try:
            # Create temporary image file
            with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{image_data["extension"]}') as tmp_file:
                tmp_file.write(image_data['data'])
                temp_image_path = tmp_file.name
            
            # Create openpyxl Image object
            img = OpenpyxlImage(temp_image_path)
            
            # Resize image to fit in cell area (adjust as needed)
            max_width = 200  # pixels
            max_height = 150  # pixels
            
            # Calculate scaling to maintain aspect ratio
            width_ratio = max_width / image_data['size'][0]
            height_ratio = max_height / image_data['size'][1]
            scale_ratio = min(width_ratio, height_ratio, 1.0)  # Don't scale up
            
            img.width = int(image_data['size'][0] * scale_ratio)
            img.height = int(image_data['size'][1] * scale_ratio)
            
            # Position the image
            # Try to place it near the placeholder cell
            target_cell = worksheet.cell(row=placeholder['row'], column=placeholder['column'] + 1)
            if not target_cell.value:  # If next cell is empty, use it
                img.anchor = target_cell.coordinate
            else:
                # Find nearby empty area
                for col_offset in range(1, 5):
                    for row_offset in range(0, 3):
                        check_cell = worksheet.cell(
                            row=placeholder['row'] + row_offset, 
                            column=placeholder['column'] + col_offset
                        )
                        if not check_cell.value:
                            img.anchor = check_cell.coordinate
                            break
                    else:
                        continue
                    break
                else:
                    # Default to next column
                    img.anchor = target_cell.coordinate
            
            # Add image to worksheet
            worksheet.add_image(img)
            
            # Clean up temporary file
            os.unlink(temp_image_path)
            
            return True
            
        except Exception as e:
            st.error(f"Error inserting image: {e}")
            if 'temp_image_path' in locals():
                try:
                    os.unlink(temp_image_path)
                except:
                    pass
            return False
    
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
                                    
                                    if keyword_processed == cell_text or keyword_processed in cell_text or cell_text in keyword_processed:
                                        return section_name
                                    
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
                r'l[-\s]*mm', r'w[-\s]*mm', r'h[-\s]*mm',
                r'l\b', r'w\b', r'h\b',
                r'packaging\s+type', r'qty[/\s]*pack',
                r'part\s+[lwh]', r'component\s+[lwh]',
                r'length', r'width', r'height',
                r'quantity', r'pack\s+weight', r'total',
                r'empty\s+weight', r'weight', r'unit\s+weight',
                r'code', r'name', r'description',
                r'vendor', r'supplier', r'customer',
                r'date', r'revision', r'reference',
                r'part\s+no', r'part\s+number'
            ]
            
            for pattern in mappable_patterns:
                if re.search(pattern, text):
                    return True
            
            if text.endswith(':'):
                return True
                
            return False
        except Exception as e:
            st.error(f"Error in is_mappable_field: {e}")
            return False
    
    def find_template_fields_with_context(self, template_file):
        """Find template fields with enhanced section context information"""
        fields = {}
        
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
        
        return fields
    
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
                        
                        for template_field_key, data_column_pattern in section_mappings.items():
                            if template_field_key in field_value or field_value in template_field_key:
                                for data_col in data_columns:
                                    if data_column_pattern.lower() == data_col.lower():
                                        best_match = data_col
                                        best_score = 1.0
                                        break
                                
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
                    continue
                    
        except Exception as e:
            st.error(f"Error in map_data_with_section_context: {e}")
            
        return mapping_results
    
    def find_data_cell_for_label(self, worksheet, field_info):
        """Find data cell for a label with improved merged cell handling"""
        try:
            row = field_info['row']
            col = field_info['column']
        
            def is_suitable_data_cell(cell_coord):
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
            
            # Look right of label first
            for offset in range(1, 6):
                target_col = col + offset
                if target_col <= worksheet.max_column:
                    cell_coord = worksheet.cell(row=row, column=target_col).coordinate
                    if is_suitable_data_cell(cell_coord):
                        return cell_coord
            
            # Look below label
            for offset in range(1, 4):
                target_row = row + offset
                if target_row <= worksheet.max_row:
                    cell_coord = worksheet.cell(row=target_row, column=col).coordinate
                    if is_suitable_data_cell(cell_coord):
                        return cell_coord
            
            return None
            
        except Exception as e:
            return None
    
    def fill_template_with_data_and_images(self, template_file, mapping_results, data_df, images=None):
        """Fill template with mapped data and insert images"""
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
            
            filled_count = 0
            images_inserted = 0
            
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
                    continue
            
            # Insert images if available
            if images:
                image_placeholders = self.find_image_placeholders(worksheet)
                
                if image_placeholders:
                    st.info(f"Found {len(image_placeholders)} image placeholders in template")
                    
                    # Match images to placeholders (up to 4 images as requested)
                    max_images = min(len(images), len(image_placeholders), 4)
                    
                    for i in range(max_images):
                        try:
                            placeholder = image_placeholders[i]
                            image_data = images[i]
                            
                            if self.insert_image_into_template(worksheet, placeholder, image_data, i):
                                images_inserted += 1
                                st.success(f"✅ Inserted image {i+1} at {placeholder['cell']}")
                            else:
                                st.warning(f"⚠️ Failed to insert image {i+1}")
                                
                        except Exception as e:
                            st.error(f"Error inserting image {i+1}: {e}")
                            continue
                else:
                    st.warning("No image placeholders found in template")
            
            return workbook, filled_count, images_inserted
            
        except Exception as e:
            st.error(f"Error filling template: {e}")
            return None, 0, 0

# Initialize session state - THIS IS THE KEY FIX
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'user_role' not in st.session_state:
    st.session_state.user_role = None
if 'templates' not in st.session_state:
    st.session_state.templates = {}
if 'enhanced_mapper' not in st.session_state:
    # Fixed: Use the correct class name
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
    st.markdown("### Advanced packaging template processing with section-aware mapping and image support")
    
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
        
        st.info("**Demo Credentials:**\n- Admin: admin / admin123\n- User: user1 / user123")

def show_enhanced_processor():
    st.header("🚀 Enhanced Template Processor with Images")
    
    # File upload section
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Upload Template")
        template_file = st.file_uploader(
            "Choose Excel template file",
            type=['xlsx', 'xlsm'],
            key="template_upload"
        )
        
        if template_file:
            st.success(f"✅ Template uploaded: {template_file.name}")
            
            # Save template temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(template_file.getvalue())
                template_path = tmp_file.name
            
            # Extract images from template
            try:
                template_images = st.session_state.enhanced_mapper.extract_images_from_excel(template_path)
                if template_images:
                    st.info(f"🖼️ Found {len(template_images)} images in template")
                    
                    # Show image preview
                    with st.expander("Preview Template Images"):
                        for i, img_info in enumerate(template_images[:4]):  # Show first 4
                            col_img1, col_img2 = st.columns([1, 3])
                            with col_img1:
                                st.write(f"**Image {i+1}:**")
                                st.write(f"Size: {img_info['size']}")
                                st.write(f"Format: {img_info['format']}")
                            with col_img2:
                                try:
                                    pil_image = PILImage.open(io.BytesIO(img_info['data']))
                                    st.image(pil_image, width=200, caption=img_info['filename'])
                                except:
                                    st.write("Could not display image")
                else:
                    st.info("No images found in template")
            except Exception as e:
                st.warning(f"Could not extract images: {e}")
                template_images = []
    
    with col2:
        st.subheader("📊 Upload Data")
        data_file = st.file_uploader(
            "Choose Excel data file",
            type=['xlsx', 'xlsm', 'csv'],
            key="data_upload"
        )
        
        if data_file:
            st.success(f"✅ Data uploaded: {data_file.name}")
            
            # Load data
            try:
                if data_file.name.endswith('.csv'):
                    data_df = pd.read_csv(data_file)
                else:
                    data_df = pd.read_excel(data_file)
                
                st.write(f"**Data Shape:** {data_df.shape}")
                st.write("**Columns:**", list(data_df.columns))
                
                # Show data preview
                with st.expander("Preview Data"):
                    st.dataframe(data_df.head())
                    
            except Exception as e:
                st.error(f"Error loading data: {e}")
                data_df = None
    
    # Processing section
    if template_file and data_file and 'data_df' in locals() and data_df is not None:
        st.subheader("⚙️ Processing Options")
        
        col_opt1, col_opt2 = st.columns(2)
        
        with col_opt1:
            similarity_threshold = st.slider(
                "Similarity Threshold",
                min_value=0.1,
                max_value=1.0,
                value=0.3,
                step=0.1,
                help="Lower values = more flexible matching"
            )
            st.session_state.enhanced_mapper.similarity_threshold = similarity_threshold
        
        with col_opt2:
            process_images = st.checkbox(
                "Include Image Processing",
                value=True,
                help="Extract and insert images into template"
            )
        
        # Process button
        if st.button("🚀 Process Template", type="primary", use_container_width=True):
            with st.spinner("Processing template..."):
                try:
                    # Find template fields
                    template_fields = st.session_state.enhanced_mapper.find_template_fields_with_context(template_path)
                    
                    if not template_fields:
                        st.error("No mappable fields found in template")
                        return
                    
                    st.success(f"Found {len(template_fields)} template fields")
                    
                    # Map data to template
                    mapping_results = st.session_state.enhanced_mapper.map_data_with_section_context(
                        template_fields, data_df
                    )
                    
                    # Show mapping results
                    st.subheader("📋 Mapping Results")
                    
                    mapped_count = sum(1 for m in mapping_results.values() if m['is_mappable'])
                    st.info(f"Successfully mapped {mapped_count} out of {len(mapping_results)} fields")
                    
                    # Display mapping table
                    mapping_display = []
                    for coord, mapping in mapping_results.items():
                        mapping_display.append({
                            'Cell': coord,
                            'Template Field': mapping['template_field'],
                            'Data Column': mapping['data_column'] or 'No match',
                            'Section': mapping['section_context'] or 'General',
                            'Similarity': f"{mapping['similarity']:.2%}",
                            'Status': '✅ Mapped' if mapping['is_mappable'] else '❌ Not mapped'
                        })
                    
                    mapping_df = pd.DataFrame(mapping_display)
                    st.dataframe(mapping_df, use_container_width=True)
                    
                    # Fill template
                    images_to_use = template_images if process_images else None
                    
                    filled_workbook, filled_count, images_inserted = st.session_state.enhanced_mapper.fill_template_with_data_and_images(
                        template_path, mapping_results, data_df, images_to_use
                    )
                    
                    if filled_workbook:
                        st.success(f"✅ Filled {filled_count} fields")
                        if process_images and images_inserted > 0:
                            st.success(f"✅ Inserted {images_inserted} images")
                        
                        # Save filled template
                        output_buffer = io.BytesIO()
                        filled_workbook.save(output_buffer)
                        output_buffer.seek(0)
                        
                        # Download button
                        st.download_button(
                            label="📥 Download Filled Template",
                            data=output_buffer.getvalue(),
                            file_name=f"filled_template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        # Cleanup
                        filled_workbook.close()
                    
                except Exception as e:
                    st.error(f"Processing failed: {e}")
                    import traceback
                    st.error(traceback.format_exc())
                
                finally:
                    # Cleanup temporary files
                    try:
                        if 'template_path' in locals():
                            os.unlink(template_path)
                    except:
                        pass

def show_mapping_rules():
    st.header("📋 Mapping Rules & Configuration")
    
    st.subheader("📦 Section-Based Mapping Rules")
    
    for section_name, section_info in st.session_state.enhanced_mapper.section_mappings.items():
        with st.expander(f"{section_name.replace('_', ' ').title()} Section"):
            st.write("**Section Keywords:**")
            for keyword in section_info['section_keywords']:
                st.write(f"- {keyword}")
            
            st.write("**Field Mappings:**")
            for template_field, data_column in section_info['field_mappings'].items():
                st.write(f"- `{template_field}` → `{data_column}`")
    
    st.subheader("🖼️ Image Processing Patterns")
    st.write("The system automatically detects these image placeholder patterns:")
    for pattern in st.session_state.enhanced_mapper.image_patterns:
        st.code(pattern)

def main():
    if not st.session_state.authenticated:
        show_login()
    else:
        # Navigation
        st.sidebar.title(f"Welcome, {st.session_state.name}")
        st.sidebar.write(f"Role: {st.session_state.user_role}")
        
        page = st.sidebar.selectbox(
            "Navigate",
            ["Enhanced Processor", "Mapping Rules", "Help"]
        )
        
        if st.sidebar.button("Logout"):
            for key in list(st.session_state.keys()):
                if key != 'enhanced_mapper':  # Keep the mapper instance
                    del st.session_state[key]
            st.rerun()
        
        # Show selected page
        if page == "Enhanced Processor":
            show_enhanced_processor()
        elif page == "Mapping Rules":
            show_mapping_rules()
        elif page == "Help":
            show_help()

def show_help():
    st.header("📚 Help & Documentation")
    
    st.subheader("🚀 Getting Started")
    st.markdown("""
    1. **Upload Template**: Choose an Excel file (.xlsx/.xlsm) containing your template
    2. **Upload Data**: Choose an Excel or CSV file containing your data
    3. **Configure**: Adjust similarity threshold and image processing options
    4. **Process**: Click "Process Template" to generate filled template
    5. **Download**: Get your completed template with data and images
    """)
    
    st.subheader("✨ Key Features")
    st.markdown("""
    - **Section-Aware Mapping**: Automatically detects packaging sections (Primary, Secondary, Part Information)
    - **Advanced Text Matching**: Uses multiple similarity algorithms for better field matching
    - **Image Processing**: Extracts and inserts images from templates
    - **Smart Field Detection**: Identifies mappable fields automatically
    - **Flexible Configuration**: Adjustable similarity thresholds
    """)
    
    st.subheader("🔧 Technical Requirements")
    st.markdown("""
    - **Template Format**: Excel files (.xlsx, .xlsm)
    - **Data Format**: Excel (.xlsx, .xlsm) or CSV files
    - **Image Support**: PNG, JPEG, GIF, BMP formats
    - **Maximum Images**: Up to 4 images per template
    """)
    
    st.subheader("⚠️ Troubleshooting")
    st.markdown("""
    - **No mappable fields found**: Check if template contains recognizable field labels
    - **Low mapping accuracy**: Try lowering the similarity threshold
    - **Images not inserting**: Ensure template has image placeholder text
    - **Processing errors**: Check file formats and data structure
    """)

if __name__ == "__main__":
    main()
