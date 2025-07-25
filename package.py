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

# Configure Streamlit page
st.set_page_config(
    page_title="AI Template Mapper - Enhanced with Image Support",
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
            'reference_images': {
                'section_keywords': [
                    'reference images', 'reference image', 'pictures', 'photos',
                    'primary packaging', 'secondary packaging', 'shipping packaging'
                ],
                'field_mappings': {
                    'primary packaging': 'Primary Packaging Image',
                    'secondary packaging': 'Secondary Packaging Image',
                    'shipping packaging': 'Shipping Packaging Image'
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
    
    def identify_section_context(self, worksheet, row, col, max_search_rows=20):
        """Enhanced section identification with better pattern matching"""
        try:
            section_context = None
            
            # Search upwards and in nearby cells for section headers
            for search_row in range(max(1, row - max_search_rows), row + 3):
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
                                    
                                    # Check for key words within the text
                                    if 'packaging' in keyword_processed and 'packaging' in cell_text:
                                        if 'primary' in cell_text or 'internal' in cell_text:
                                            return 'primary_packaging'
                                        elif 'secondary' in cell_text or 'outer' in cell_text or 'external' in cell_text:
                                            return 'secondary_packaging'
                                    elif 'part' in keyword_processed and 'part' in cell_text:
                                        return 'part_information'
                                    elif 'vendor' in keyword_processed and 'vendor' in cell_text:
                                        return 'vendor_information'
                                    elif 'reference' in keyword_processed and 'image' in keyword_processed:
                                        return 'reference_images'
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
    
    def extract_images_from_template(self, template_file):
        """Extract images from Excel template"""
        images_info = {}
        
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
            
            # Check if worksheet has images
            if hasattr(worksheet, '_images') and worksheet._images:
                for i, img in enumerate(worksheet._images):
                    try:
                        # Get image position
                        anchor = img.anchor
                        if hasattr(anchor, '_from'):
                            col = anchor._from.col
                            row = anchor._from.row
                        else:
                            col = 0
                            row = 0
                        
                        # Determine image context based on position
                        section_context = self.identify_section_context(worksheet, row + 1, col + 1)
                        
                        # Extract image data
                        img_data = img._data()
                        
                        images_info[f"image_{i}"] = {
                            'data': img_data,
                            'position': f"Row {row + 1}, Col {col + 1}",
                            'section_context': section_context,
                            'anchor': anchor
                        }
                        
                    except Exception as e:
                        print(f"Error extracting image {i}: {e}")
                        continue
            
            workbook.close()
            
        except Exception as e:
            st.error(f"Error extracting images: {e}")
        
        return images_info
    
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

# Initialize session state
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
    st.markdown("### Advanced packaging template processing with image support and section-aware mapping")
    
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
    """Display template preview with extracted images"""
    try:
        # Create temporary template file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(template_info['file_data'])
            template_path = tmp_file.name
        
        # Extract images from template
        images_info = st.session_state.enhanced_mapper.extract_images_from_template(template_path)
        
        # Analyze template structure
        template_fields = st.session_state.enhanced_mapper.find_template_fields_with_context(template_path)
        
        os.unlink(template_path)
        
        # Display template structure
        st.subheader(f"📋 Template Structure: {template_name}")
        
        # Group fields by section
        section_analysis = {}
        for coord, field in template_fields.items():
            section = field.get('section_context', 'general')
            if section not in section_analysis:
                section_analysis[section] = []
            section_analysis[section].append(field)
        
        # Display sections
        for section, fields in section_analysis.items():
            section_name = section.replace('_', ' ').title() if section else 'General'
            with st.expander(f"📦 {section_name} Section - {len(fields)} fields", expanded=False):
                fields_df = pd.DataFrame([
                    {
                        'Field': field['value'],
                        'Position': f"Row {field['row']}, Col {field['column']}",
                        'Mappable': '✅' if field['is_mappable'] else '❌',
                        'Section': section_name
                    }
                    for field in fields
                ])
                st.dataframe(fields_df, use_container_width=True)
        
        # Display extracted images
        if images_info:
            st.subheader("🖼️ Extracted Template Images")
            
            cols = st.columns(min(3, len(images_info)))
            for i, (img_key, img_info) in enumerate(images_info.items()):
                with cols[i % 3]:
                    try:
                        # Display image
                        st.image(img_info['data'], caption=f"Position: {img_info['position']}")
                        st.write(f"**Section:** {img_info.get('section_context', 'Unknown')}")
                    except Exception as e:
                        st.error(f"Error displaying image: {e}")
        else:
            st.info("No images found in template")
            
    except Exception as e:
        st.error(f"Error analyzing template: {e}")

def show_enhanced_processor():
    st.header("🚀 Enhanced AI Data Processor")
    st.info("Upload your packaging data and template. AI will intelligently map fields with image support!")
    
    # Show enhanced mapping rules
    with st.expander("📋 Enhanced Section Mapping Rules", expanded=False):
        st.markdown("""
        **Vendor Information Section:**
        - `Vendor` ← `Vendor Name`
        - `Code` ← `Vendor Code`
        - `Location` ← `Vendor Location`
        
        **Part Information Section:**
        - `Part No.` ← `Part No`
        - `Description` ← `Part Description`
        - `L/W/H` ← `Part L/W/H`
        - `Unit Weight` ← `Part Unit Weight`
        
        **Primary Packaging Section:**
        - `Packaging Type` ← `Primary Packaging Type`
        - `L-mm/W-mm/H-mm` ← `Primary L-mm/W-mm/H-mm`
        - `Qty/Pack` ← `Primary Qty/Pack`
        - `Empty Weight/Pack Weight` ← `Primary Empty Weight/Pack Weight`
        
        **Secondary Packaging Section:**
        - `Packaging Type` ← `Secondary Packaging Type`
        - `L-mm/W-mm/H-mm` ← `Secondary L-mm/W-mm/H-mm`
        - `Qty/Pack` ← `Secondary Qty/Pack`
        - `Empty Weight/Pack Weight` ← `Secondary Empty Weight/Pack Weight`
        
        **Reference Images Section:**
        - `Primary Packaging` ← `Primary Packaging Image`
        - `Secondary Packaging` ← `Secondary Packaging Image`
        - `Shipping Packaging` ← `Shipping Packaging Image`
        """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Upload Data File")
        data_file = st.file_uploader(
            "Choose CSV or Excel file",
            type=['csv', 'xlsx', 'xls'],
            key="data_upload",
            help="Upload your packaging data file"
        )
        
        if data_file:
            try:
                # Read data file
                if data_file.name.endswith('.csv'):
                    data_df = pd.read_csv(data_file)
                else:
                    data_df = pd.read_excel(data_file)
                
                st.success(f"✅ Data loaded: {len(data_df)} rows, {len(data_df.columns)} columns")
                
                # Show data preview
                with st.expander("🔍 Data Preview", expanded=False):
                    st.dataframe(data_df.head(10), use_container_width=True)
                
                # Show column analysis
                with st.expander("📋 Column Analysis", expanded=False):
                    cols_info = []
                    for col in data_df.columns:
                        non_null_count = data_df[col].count()
                        null_count = len(data_df) - non_null_count
                        data_type = str(data_df[col].dtype)
                        
                        cols_info.append({
                            'Column': col,
                            'Non-Null': non_null_count,
                            'Null': null_count,
                            'Data Type': data_type,
                            'Sample Value': str(data_df[col].dropna().iloc[0]) if non_null_count > 0 else 'N/A'
                        })
                    
                    cols_df = pd.DataFrame(cols_info)
                    st.dataframe(cols_df, use_container_width=True)
                
            except Exception as e:
                st.error(f"Error reading data file: {e}")
                data_df = None
        else:
            data_df = None
    
    with col2:
        st.subheader("📄 Upload Template")
        template_file = st.file_uploader(
            "Choose Excel template",
            type=['xlsx', 'xls'],
            key="template_upload",
            help="Upload your Excel template with images"
        )
        
        if template_file:
            try:
                # Save template temporarily for analysis
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                    tmp_file.write(template_file.read())
                    template_path = tmp_file.name
                
                # Analyze template
                template_fields = st.session_state.enhanced_mapper.find_template_fields_with_context(template_path)
                images_info = st.session_state.enhanced_mapper.extract_images_from_template(template_path)
                
                st.success(f"✅ Template loaded: {len(template_fields)} fields, {len(images_info)} images")
                
                # Show template analysis
                with st.expander("🔍 Template Analysis", expanded=False):
                    # Group by sections
                    section_counts = {}
                    for field in template_fields.values():
                        section = field.get('section_context', 'general')
                        section_counts[section] = section_counts.get(section, 0) + 1
                    
                    st.write("**Sections Detected:**")
                    for section, count in section_counts.items():
                        section_name = section.replace('_', ' ').title() if section else 'General'
                        st.write(f"- {section_name}: {count} fields")
                    
                    if images_info:
                        st.write(f"**Images Found:** {len(images_info)}")
                        for img_key, img_info in images_info.items():
                            st.write(f"- {img_key}: {img_info.get('section_context', 'Unknown')} section")
                
                os.unlink(template_path)
                
            except Exception as e:
                st.error(f"Error analyzing template: {e}")
                template_fields = None
                images_info = None
        else:
            template_fields = None
            images_info = None
    
    # Processing section
    if data_df is not None and template_file is not None:
        st.header("🎯 Intelligent Field Mapping")
        
        # Configuration options
        col1, col2, col3 = st.columns(3)
        
        with col1:
            similarity_threshold = st.slider(
                "Similarity Threshold",
                min_value=0.1,
                max_value=1.0,
                value=0.3,
                step=0.1,
                help="Minimum similarity score for field matching"
            )
            st.session_state.enhanced_mapper.similarity_threshold = similarity_threshold
        
        with col2:
            process_images = st.checkbox(
                "Process Images",
                value=True,
                help="Include image processing in template filling"
            )
        
        with col3:
            auto_fill = st.checkbox(
                "Auto-fill Template",
                value=True,
                help="Automatically fill template with first data row"
            )
        
        if st.button("🚀 Start Enhanced Processing", type="primary", use_container_width=True):
            with st.spinner("Processing with AI-powered mapping..."):
                try:
                    # Create temporary template file
                    template_file.seek(0)
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                        tmp_file.write(template_file.read())
                        template_path = tmp_file.name
                    
                    # Get template fields
                    template_fields = st.session_state.enhanced_mapper.find_template_fields_with_context(template_path)
                    
                    # Perform enhanced mapping
                    mapping_results = st.session_state.enhanced_mapper.map_data_with_section_context(
                        template_fields, data_df
                    )
                    
                    # Display mapping results
                    st.subheader("📊 Mapping Results")
                    
                    # Create mapping summary
                    mapping_summary = []
                    successful_mappings = 0
                    
                    for coord, mapping in mapping_results.items():
                        section_name = mapping.get('section_context', 'general')
                        section_name = section_name.replace('_', ' ').title() if section_name else 'General'
                        
                        mapping_summary.append({
                            'Template Field': mapping['template_field'],
                            'Data Column': mapping['data_column'] if mapping['data_column'] else '❌ No Match',
                            'Similarity': f"{mapping['similarity']:.2f}" if mapping['similarity'] > 0 else '0.00',
                            'Section': section_name,
                            'Status': '✅ Mapped' if mapping['is_mappable'] else '❌ Not Mapped'
                        })
                        
                        if mapping['is_mappable']:
                            successful_mappings += 1
                    
                    mapping_df = pd.DataFrame(mapping_summary)
                    
                    # Display statistics
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Fields", len(mapping_results))
                    with col2:
                        st.metric("Successful Mappings", successful_mappings)
                    with col3:
                        mapping_rate = (successful_mappings / len(mapping_results)) * 100 if mapping_results else 0
                        st.metric("Mapping Rate", f"{mapping_rate:.1f}%")
                    with col4:
                        st.metric("Data Rows", len(data_df))
                    
                    # Display detailed mapping table
                    st.dataframe(mapping_df, use_container_width=True)
                    
                    # Group by sections for better visualization
                    st.subheader("📦 Section-wise Mapping Analysis")
                    
                    section_groups = mapping_df.groupby('Section')
                    for section_name, group in section_groups:
                        with st.expander(f"📋 {section_name} Section", expanded=False):
                            st.dataframe(group[['Template Field', 'Data Column', 'Similarity', 'Status']], use_container_width=True)
                    
                    # Auto-fill template if requested
                    if auto_fill and successful_mappings > 0:
                        st.subheader("📝 Template Filling")
                        
                        with st.spinner("Filling template with data..."):
                            filled_workbook, filled_count = st.session_state.enhanced_mapper.fill_template_with_data(
                                template_path, mapping_results, data_df
                            )
                            
                            if filled_workbook and filled_count > 0:
                                st.success(f"✅ Successfully filled {filled_count} fields in template!")
                                
                                # Save filled template
                                output_buffer = io.BytesIO()
                                filled_workbook.save(output_buffer)
                                output_buffer.seek(0)
                                
                                # Create download button
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                filename = f"filled_template_{timestamp}.xlsx"
                                
                                st.download_button(
                                    label="📥 Download Filled Template",
                                    data=output_buffer.read(),
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                                
                                filled_workbook.close()
                            else:
                                st.warning("⚠️ No data was filled in the template. Check your mappings.")
                    
                    # Process multiple rows if requested
                    if len(data_df) > 1:
                        st.subheader("📊 Batch Processing")
                        
                        process_all = st.checkbox(
                            f"Process all {len(data_df)} rows",
                            help="Generate filled templates for all data rows"
                        )
                        
                        if process_all and st.button("🔄 Process All Rows", type="secondary"):
                            with st.spinner(f"Processing {len(data_df)} rows..."):
                                
                                # Create ZIP file for multiple templates
                                zip_buffer = io.BytesIO()
                                
                                with tempfile.TemporaryDirectory() as temp_dir:
                                    for idx, row in data_df.iterrows():
                                        try:
                                            # Create single-row dataframe
                                            single_row_df = pd.DataFrame([row])
                                            
                                            # Fill template
                                            filled_workbook, filled_count = st.session_state.enhanced_mapper.fill_template_with_data(
                                                template_path, mapping_results, single_row_df
                                            )
                                            
                                            if filled_workbook and filled_count > 0:
                                                # Save to temp directory
                                                temp_filename = f"template_row_{idx+1}.xlsx"
                                                temp_path = os.path.join(temp_dir, temp_filename)
                                                filled_workbook.save(temp_path)
                                                filled_workbook.close()
                                            
                                        except Exception as e:
                                            st.error(f"Error processing row {idx+1}: {e}")
                                            continue
                                    
                                    # Create ZIP archive
                                    shutil.make_archive(
                                        os.path.join(temp_dir, "filled_templates"), 
                                        'zip', 
                                        temp_dir
                                    )
                                    
                                    # Read ZIP file
                                    zip_path = os.path.join(temp_dir, "filled_templates.zip")
                                    if os.path.exists(zip_path):
                                        with open(zip_path, 'rb') as zip_file:
                                            zip_data = zip_file.read()
                                        
                                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                        zip_filename = f"batch_templates_{timestamp}.zip"
                                        
                                        st.download_button(
                                            label=f"📥 Download All Templates ({len(data_df)} files)",
                                            data=zip_data,
                                            file_name=zip_filename,
                                            mime="application/zip",
                                            use_container_width=True
                                        )
                                        
                                        st.success(f"✅ Batch processing completed for {len(data_df)} rows!")
                    
                    # Clean up
                    os.unlink(template_path)
                    
                except Exception as e:
                    st.error(f"Error during processing: {e}")
                    import traceback
                    st.error(traceback.format_exc())

def show_template_manager():
    st.header("📚 Template Manager")
    
    # Template upload section
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("📤 Upload New Template")
        
        with st.form("template_upload_form"):
            template_name = st.text_input("Template Name", placeholder="e.g., Packaging Instructions v2.0")
            template_description = st.text_area("Description", placeholder="Describe this template...")
            template_file = st.file_uploader("Choose Excel template", type=['xlsx', 'xls'])
            
            if st.form_submit_button("Upload Template", use_container_width=True):
                if template_name and template_file:
                    try:
                        # Store template
                        template_data = {
                            'name': template_name,
                            'description': template_description,
                            'file_data': template_file.read(),
                            'filename': template_file.name,
                            'upload_date': datetime.now().isoformat(),
                            'uploaded_by': st.session_state.username
                        }
                        
                        st.session_state.templates[template_name] = template_data
                        st.success(f"✅ Template '{template_name}' uploaded successfully!")
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"Error uploading template: {e}")
                else:
                    st.error("Please provide template name and file")
    
    with col2:
        st.subheader("📊 Template Statistics")
        st.metric("Total Templates", len(st.session_state.templates))
        
        if st.session_state.templates:
            latest_template = max(
                st.session_state.templates.values(),
                key=lambda x: x['upload_date']
            )
            st.metric("Latest Upload", latest_template['name'])
    
    # Template list section
    if st.session_state.templates:
        st.subheader("📋 Stored Templates")
        
        for template_name, template_info in st.session_state.templates.items():
            with st.expander(f"📄 {template_name}", expanded=False):
                col1, col2, col3 = st.columns([2, 1, 1])
                
                with col1:
                    st.write(f"**Description:** {template_info.get('description', 'No description')}")
                    st.write(f"**Filename:** {template_info['filename']}")
                    st.write(f"**Uploaded:** {template_info['upload_date'][:10]}")
                    st.write(f"**By:** {template_info.get('uploaded_by', 'Unknown')}")
                
                with col2:
                    if st.button(f"📥 Download", key=f"download_{template_name}"):
                        st.download_button(
                            label="Download Template",
                            data=template_info['file_data'],
                            file_name=template_info['filename'],
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_btn_{template_name}"
                        )
                
                with col3:
                    if st.session_state.user_role == 'admin':
                        if st.button(f"🗑️ Delete", key=f"delete_{template_name}"):
                            del st.session_state.templates[template_name]
                            st.success(f"Template '{template_name}' deleted!")
                            st.rerun()
                
                # Show template preview
                display_template_preview(template_name, template_info)
    
    else:
        st.info("📝 No templates uploaded yet. Upload your first template above!")

def main():
    if not st.session_state.authenticated:
        show_login()
        return
    
    # Sidebar navigation
    with st.sidebar:
        st.title(f"👋 Welcome, {st.session_state.name}")
        st.write(f"**Role:** {st.session_state.user_role.title()}")
        
        st.divider()
        
        # Navigation menu
        page = st.selectbox(
            "🧭 Navigation",
            ["🚀 Enhanced Processor", "📚 Template Manager", "⚙️ Settings"],
            key="nav_selection"
        )
        
        st.divider()
        
        # System information
        st.subheader("🔧 System Info")
        st.write(f"**NLP Status:** {'✅ Advanced' if ADVANCED_NLP else '⚠️ Basic'}")
        st.write(f"**NLTK Ready:** {'✅ Yes' if NLTK_READY else '❌ No'}")
        st.write(f"**Templates:** {len(st.session_state.templates)}")
        
        st.divider()
        
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.user_role = None
            st.rerun()
    
    # Main content
    if page == "🚀 Enhanced Processor":
        show_enhanced_processor()
    elif page == "📚 Template Manager":
        show_template_manager()
    elif page == "⚙️ Settings":
        st.header("⚙️ Settings")
        st.info("Settings panel - Coming soon!")
        
        # Current settings display
        with st.expander("🔧 Current Configuration", expanded=True):
            st.write(f"**Similarity Threshold:** {st.session_state.enhanced_mapper.similarity_threshold}")
            st.write(f"**Advanced NLP:** {'Enabled' if ADVANCED_NLP else 'Disabled'}")
            st.write(f"**User Role:** {st.session_state.user_role}")

if __name__ == "__main__":
    main()
