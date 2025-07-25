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
from PIL import Image, ImageDraw, ImageFont
import base64
import zipfile

# Configure Streamlit page
st.set_page_config(
    page_title="AI Template Mapper - Enhanced Image Extraction",
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

class EnhancedTemplateMapper:
    def __init__(self):
        self.similarity_threshold = 0.3
        self.stop_words = {
            'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from',
            'has', 'he', 'in', 'is', 'it', 'its', 'of', 'on', 'that', 'the',
            'to', 'was', 'will', 'with', 'or', 'but', 'not', 'this', 'have',
            'had', 'what', 'when', 'where', 'who', 'which', 'why', 'how'
        }
        
        # Enhanced section-based mapping rules for packaging instructions
        self.section_mappings = {
            'vendor_information': {
                'section_keywords': [
                    'vendor information', 'vendor', 'supplier information', 'supplier'
                ],
                'field_mappings': {
                    'vendor': 'Vendor Name',
                    'code': 'Vendor Code',
                    'location': 'Vendor Location'
                }
            },
            'part_information': {
                'section_keywords': [
                    'part information', 'part', 'component', 'item information'
                ],
                'field_mappings': {
                    'part no': 'Part No',
                    'part number': 'Part No',
                    'description': 'Part Description',
                    'unit weight': 'Part Unit Weight',
                    'l': 'Part L',
                    'w': 'Part W', 
                    'h': 'Part H',
                    'length': 'Part L',
                    'width': 'Part W',
                    'height': 'Part H'
                }
            },
            'primary_packaging': {
                'section_keywords': [
                    'primary packaging instruction', 'primary packaging', 'primary', 
                    'internal', 'primary / internal', 'inner packaging'
                ],
                'field_mappings': {
                    'packaging type': 'Primary Packaging Type',
                    'primary packaging type': 'Primary Packaging Type',
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
                    'outer', 'external', 'outer / external', 'outer packaging'
                ],
                'field_mappings': {
                    'packaging type': 'Secondary Packaging Type',
                    'secondary packaging type': 'Secondary Packaging Type',
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
            'current_packaging': {
                'section_keywords': [
                    'current packaging', 'existing packaging', 'present packaging'
                ],
                'field_mappings': {
                    'current': 'Current Packaging Image',
                    'existing': 'Existing Packaging Image'
                }
            },
            'reference_images': {
                'section_keywords': [
                    'reference images', 'reference image', 'pictures', 'photos',
                    'primary packaging', 'secondary packaging', 'shipping packaging'
                ],
                'field_mappings': {
                    'primary packaging': 'Primary Packaging Image',
                    'secondary packaging': 'Secondary Packaging Image', 
                    'shipping packaging': 'Shipping Packaging Image',
                    'current packaging': 'Current Packaging Image'
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
    
    def identify_section_context(self, worksheet, row, col, max_search_rows=25):
        """Enhanced section identification with better pattern matching"""
        try:
            section_context = None
            
            # Search upwards and in nearby cells for section headers
            for search_row in range(max(1, row - max_search_rows), row + 5):
                for search_col in range(max(1, col - 20), min(worksheet.max_column + 1, col + 20)):
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
                                    if 'packaging' in keyword_processed and 'packaging' in cell_text:
                                        if 'current' in cell_text:
                                            return 'current_packaging'
                                        elif 'primary' in cell_text or 'internal' in cell_text:
                                            return 'primary_packaging'
                                        elif 'secondary' in cell_text or 'outer' in cell_text or 'external' in cell_text:
                                            return 'secondary_packaging'
                                    elif 'part' in keyword_processed and 'part' in cell_text:
                                        return 'part_information'
                                    elif 'vendor' in keyword_processed and 'vendor' in cell_text:
                                        return 'vendor_information'
                                    elif ('reference' in keyword_processed and 'image' in keyword_processed) or 'pictures' in cell_text:
                                        return 'reference_images'
                    except:
                        continue
            
            return section_context
        except Exception as e:
            st.error(f"Error in identify_section_context: {e}")
            return None
    
    def extract_images_from_template_improved(self, template_file):
        """Improved image extraction from Excel template with better positioning and context detection"""
        images_info = {}
        
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
            
            # Method 1: Extract from worksheet._images (standard embedded images)
            if hasattr(worksheet, '_images') and worksheet._images:
                st.info(f"Found {len(worksheet._images)} embedded images using standard method")
                
                for i, img in enumerate(worksheet._images):
                    try:
                        # Get image position information
                        anchor = img.anchor
                        row, col = 0, 0
                        
                        # Try different anchor types
                        if hasattr(anchor, '_from') and anchor._from:
                            row = anchor._from.row if hasattr(anchor._from, 'row') else 0
                            col = anchor._from.col if hasattr(anchor._from, 'col') else 0
                        elif hasattr(anchor, 'row') and hasattr(anchor, 'col'):
                            row = anchor.row
                            col = anchor.col
                        
                        # Determine image context based on position and surrounding text
                        section_context = self.identify_section_context(worksheet, row + 1, col + 1, max_search_rows=30)
                        
                        # Try to get a more descriptive name based on context
                        image_name = f"image_{i+1}"
                        if section_context:
                            section_name = section_context.replace('_', ' ').title()
                            image_name = f"{section_name}_Image_{i+1}"
                        
                        # Extract image data
                        try:
                            img_data = img._data()
                        except:
                            # Alternative method to get image data
                            img_data = img.ref
                        
                        images_info[image_name] = {
                            'data': img_data,
                            'position': f"Row {row + 1}, Col {col + 1}",
                            'section_context': section_context or 'unknown',
                            'anchor': anchor,
                            'image_index': i,
                            'extraction_method': 'standard'
                        }
                        
                    except Exception as e:
                        st.warning(f"Error extracting image {i}: {e}")
                        continue
            
            # Method 2: Check for images in drawing parts (alternative method)
            try:
                if hasattr(worksheet, '_drawing') and worksheet._drawing:
                    drawing = worksheet._drawing
                    if hasattr(drawing, '_charts') or hasattr(drawing, 'charts'):
                        st.info("Checking alternative drawing objects...")
                        # Additional logic for other drawing objects can be added here
            except Exception as e:
                st.warning(f"Alternative image extraction failed: {e}")
            
            # Method 3: Scan for image-related cell content and shapes
            try:
                # Look for cells that might reference images or have special formatting
                for row in range(1, min(worksheet.max_row + 1, 100)):  # Limit search to reasonable range
                    for col in range(1, min(worksheet.max_column + 1, 50)):
                        cell = worksheet.cell(row=row, column=col)
                        if cell.value:
                            cell_text = str(cell.value).lower()
                            # Look for image-related keywords
                            if any(keyword in cell_text for keyword in ['image', 'picture', 'photo', 'current packaging']):
                                section_context = self.identify_section_context(worksheet, row, col)
                                if section_context and section_context not in [img['section_context'] for img in images_info.values()]:
                                    # This might be a placeholder for an image
                                    placeholder_name = f"Placeholder_{section_context}_{row}_{col}"
                                    images_info[placeholder_name] = {
                                        'data': None,
                                        'position': f"Row {row}, Col {col}",
                                        'section_context': section_context,
                                        'anchor': None,
                                        'image_index': -1,
                                        'extraction_method': 'placeholder',
                                        'cell_text': cell_text
                                    }
            except Exception as e:
                st.warning(f"Placeholder detection failed: {e}")
            
            workbook.close()
            
        except Exception as e:
            st.error(f"Error extracting images: {e}")
        
        return images_info
    
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
                r'part\s+no', r'part\s+number',             # Part identification
                r'location', r'address'                     # Location fields
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
                                
                                # Identify section context with enhanced detection
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
            for offset in range(1, 8):
                target_col = col + offset
                if target_col <= worksheet.max_column:
                    cell_coord = worksheet.cell(row=row, column=target_col).coordinate
                    if is_suitable_data_cell(cell_coord):
                        return cell_coord
            
            # Strategy 2: Look below label
            for offset in range(1, 5):
                target_row = row + offset
                if target_row <= worksheet.max_row:
                    cell_coord = worksheet.cell(row=target_row, column=col).coordinate
                    if is_suitable_data_cell(cell_coord):
                        return cell_coord
            
            # Strategy 3: Look in nearby area
            for r_offset in range(-1, 4):
                for c_offset in range(-1, 8):
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
    
    def fill_template_with_data(self, template_file, mapping_results, data_df):
        """Fill template with mapped data"""
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
            
            filled_count = 0
            
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
            
            return workbook, filled_count
            
        except Exception as e:
            st.error(f"Error filling template: {e}")
            return None, 0

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'user_role' not in st.session_state:
    st.session_state.user_role = None
if 'templates' not in st.session_state:
    st.session_state.templates = {}
if 'enhanced_mapper' not in st.session_state:
    st.session_state.enhanced_mapper = EnhancedTemplateMapper()

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
    st.title("🤖 Enhanced AI Template Mapper")
    st.markdown("### Advanced packaging template processing with improved image extraction")
    
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

def display_template_preview(template_name, template_info):
    """Display template preview with improved image extraction"""
    try:
        # Create temporary template file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(template_info['file_data'])
            template_path = tmp_file.name
        
        # Extract images from template using improved method
        images_info = st.session_state.enhanced_mapper.extract_images_from_template_improved(template_path)
        
        st.write(f"**Template:** {template_name}")
        st.write(f"**Description:** {template_info.get('description', 'No description available')}")
        st.write(f"**Uploaded:** {template_info.get('upload_date', 'Unknown')}")
        
        # Display extracted images information
        if images_info:
            st.subheader("📸 Extracted Images")
            
            cols = st.columns(min(3, len(images_info)))
            for idx, (img_name, img_info) in enumerate(images_info.items()):
                with cols[idx % 3]:
                    st.write(f"**{img_name}**")
                    st.write(f"Position: {img_info['position']}")
                    st.write(f"Context: {img_info['section_context']}")
                    st.write(f"Method: {img_info['extraction_method']}")
                    
                    if img_info['data'] is not None:
                        try:
                            # Try to display the image
                            image_data = img_info['data']
                            if isinstance(image_data, bytes):
                                st.image(image_data, caption=img_name, width=200)
                            else:
                                st.info("Image data available but format not displayable")
                        except Exception as e:
                            st.warning(f"Could not display image: {e}")
                    else:
                        st.info("Image placeholder detected")
        else:
            st.info("No images found in template")
        
        # Clean up temporary file
        os.unlink(template_path)
        
    except Exception as e:
        st.error(f"Error displaying template preview: {e}")

def show_template_management():
    """Template management interface"""
    st.header("📋 Template Management")
    
    # Template upload section
    with st.expander("➕ Upload New Template", expanded=True):
        template_name = st.text_input("Template Name")
        template_description = st.text_area("Template Description")
        uploaded_template = st.file_uploader(
            "Choose Excel template file",
            type=['xlsx', 'xls'],
            help="Upload your Excel template with packaging instruction fields"
        )
        
        if st.button("Upload Template"):
            if uploaded_template and template_name:
                try:
                    # Store template
                    template_data = {
                        'file_data': uploaded_template.read(),
                        'description': template_description,
                        'upload_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'uploaded_by': st.session_state.username
                    }
                    
                    st.session_state.templates[template_name] = template_data
                    st.success(f"Template '{template_name}' uploaded successfully!")
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"Error uploading template: {e}")
            else:
                st.warning("Please provide template name and file")
    
    # Display existing templates
    if st.session_state.templates:
        st.subheader("📁 Existing Templates")
        
        for template_name, template_info in st.session_state.templates.items():
            with st.expander(f"📄 {template_name}"):
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    display_template_preview(template_name, template_info)
                
                with col2:
                    if st.button(f"Delete", key=f"delete_{template_name}"):
                        del st.session_state.templates[template_name]
                        st.success(f"Template '{template_name}' deleted!")
                        st.rerun()
    else:
        st.info("No templates uploaded yet")

def show_data_processing():
    """Data processing and mapping interface"""
    st.header("🔄 Data Processing & Mapping")
    
    if not st.session_state.templates:
        st.warning("Please upload templates first in Template Management")
        return
    
    # Template selection
    selected_template = st.selectbox(
        "Select Template",
        options=list(st.session_state.templates.keys()),
        help="Choose the template to map data to"
    )
    
    # Data file upload
    uploaded_data = st.file_uploader(
        "Upload Data File",
        type=['xlsx', 'xls', 'csv'],
        help="Upload your data file with values to map to the template"
    )
    
    if uploaded_data and selected_template:
        try:
            # Load data
            if uploaded_data.name.endswith('.csv'):
                data_df = pd.read_csv(uploaded_data)
            else:
                data_df = pd.read_excel(uploaded_data)
            
            st.success(f"Data loaded: {len(data_df)} rows, {len(data_df.columns)} columns")
            
            # Show data preview
            with st.expander("📊 Data Preview"):
                st.dataframe(data_df.head())
            
            # Process mapping
            if st.button("🔍 Analyze Template Fields"):
                with st.spinner("Analyzing template fields..."):
                    # Create temporary template file
                    template_info = st.session_state.templates[selected_template]
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                        tmp_file.write(template_info['file_data'])
                        template_path = tmp_file.name
                    
                    try:
                        # Find template fields
                        mapper = st.session_state.enhanced_mapper
                        template_fields = mapper.find_template_fields_with_context(template_path)
                        
                        st.success(f"Found {len(template_fields)} template fields")
                        
                        # Perform mapping
                        st.write("✅ Mapped fields:")
                        for k, m in mapping_results.items():
                            st.write(f"{k}: {m['template_field']} → {m['data_column']} (Mapped: {m['is_mappable']})")
                        
                        # Display mapping results
                        st.subheader("🎯 Field Mapping Results")
                        
                        # Create mapping summary
                        mapping_summary = []
                        for coord, mapping in mapping_results.items():
                            mapping_summary.append({
                                'Template Field': mapping['template_field'],
                                'Data Column': mapping['data_column'] if mapping['data_column'] else 'Not Mapped',
                                'Similarity': f"{mapping['similarity']:.2f}" if mapping['similarity'] > 0 else "0.00",
                                'Section': mapping['section_context'] or 'Unknown',
                                'Position': coord,
                                'Status': '✅ Mapped' if mapping['is_mappable'] else '❌ Not Mapped'
                            })
                        
                        mapping_df = pd.DataFrame(mapping_summary)
                        st.dataframe(mapping_df, use_container_width=True)
                        
                        # Statistics
                        mapped_count = sum(1 for m in mapping_results.values() if m['is_mappable'])
                        total_count = len(mapping_results)
                        mapping_rate = (mapped_count / total_count * 100) if total_count > 0 else 0
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Total Fields", total_count)
                        with col2:
                            st.metric("Mapped Fields", mapped_count)
                        with col3:
                            st.metric("Mapping Rate", f"{mapping_rate:.1f}%")
                        
                        # Generate filled template
                        if st.button("📝 Generate Filled Template"):
                            with st.spinner("Generating filled template..."):
                                try:
                                    filled_workbook, filled_count = mapper.fill_template_with_data(
                                        template_path, mapping_results, data_df
                                    )
                                    st.write("DEBUG: Filled workbook object:", filled_workbook)
                                    st.write("DEBUG: Number of fields filled:", filled_count)
                                    
                                    if filled_workbook:
                                        # Save filled template
                                        output_buffer = io.BytesIO()
                                        filled_workbook.save(output_buffer)
                                        output_buffer.seek(0)
                                        
                                        st.success(f"Template filled! {filled_count} fields populated.")
                                        
                                        # Download button
                                        st.download_button(
                                            label="📥 Download Filled Template",
                                            data=output_buffer.getvalue(),
                                            file_name=f"filled_{selected_template}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        )
                                    else:
                                        st.error("Failed to generate filled template")
                                        
                                except Exception as e:
                                    st.error(f"Error generating filled template: {e}")
                    
                    finally:
                        # Clean up temporary file
                        os.unlink(template_path)
                        
        except Exception as e:
            st.error(f"Error processing data: {e}")

def show_analytics():
    """Analytics and reporting interface"""
    st.header("📊 Analytics & Reports")
    
    if not st.session_state.templates:
        st.warning("No templates available for analysis")
        return
    
    # Template analytics
    st.subheader("📋 Template Statistics")
    
    template_stats = []
    for template_name, template_info in st.session_state.templates.items():
        try:
            # Create temporary file to analyze
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(template_info['file_data'])
                template_path = tmp_file.name
            
            # Analyze template
            mapper = st.session_state.enhanced_mapper
            template_fields = mapper.find_template_fields_with_context(template_path)
            images_info = mapper.extract_images_from_template_improved(template_path)
            
            template_stats.append({
                'Template Name': template_name,
                'Total Fields': len(template_fields),
                'Images Found': len(images_info),
                'Upload Date': template_info.get('upload_date', 'Unknown'),
                'Uploaded By': template_info.get('uploaded_by', 'Unknown')
            })
            
            os.unlink(template_path)
            
        except Exception as e:
            st.error(f"Error analyzing template {template_name}: {e}")
    
    if template_stats:
        stats_df = pd.DataFrame(template_stats)
        st.dataframe(stats_df, use_container_width=True)
        
        # Visual analytics
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📈 Fields per Template")
            st.bar_chart(stats_df.set_index('Template Name')['Total Fields'])
        
        with col2:
            st.subheader("🖼️ Images per Template")
            st.bar_chart(stats_df.set_index('Template Name')['Images Found'])

def main():
    """Main application function"""
    
    # Authentication check
    if not st.session_state.authenticated:
        show_login()
        return
    
    # Sidebar navigation
    st.sidebar.title(f"Welcome, {st.session_state.name}")
    st.sidebar.write(f"Role: {st.session_state.user_role}")
    
    # Navigation menu
    menu_options = ["Template Management", "Data Processing", "Analytics"]
    if st.session_state.user_role == "admin":
        menu_options.append("System Settings")
    
    selected_menu = st.sidebar.selectbox("Navigation", menu_options)
    
    # Logout button
    if st.sidebar.button("Logout"):
        st.session_state.authenticated = False
        st.session_state.user_role = None
        st.rerun()
    
    # Main content based on menu selection
    if selected_menu == "Template Management":
        show_template_management()
    elif selected_menu == "Data Processing":
        show_data_processing()
    elif selected_menu == "Analytics":
        show_analytics()
    elif selected_menu == "System Settings" and st.session_state.user_role == "admin":
        st.header("⚙️ System Settings")
        st.info("System settings panel - Feature under development")
        
        # Advanced settings
        st.subheader("🔧 Advanced Configuration")
        
        similarity_threshold = st.slider(
            "Similarity Threshold",
            min_value=0.1,
            max_value=1.0,
            value=st.session_state.enhanced_mapper.similarity_threshold,
            step=0.05,
            help="Minimum similarity score for field matching"
        )
        
        if st.button("Update Settings"):
            st.session_state.enhanced_mapper.similarity_threshold = similarity_threshold
            st.success("Settings updated successfully!")
    
    # Footer
    st.sidebar.markdown("---")
    st.sidebar.markdown("🤖 **Enhanced AI Template Mapper**")
    st.sidebar.markdown("Advanced packaging template processing")
    if ADVANCED_NLP:
        st.sidebar.success("🚀 Advanced NLP: Enabled")
    else:
        st.sidebar.warning("⚠️ Advanced NLP: Disabled")

if __name__ == "__main__":
    main()
