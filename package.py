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
import re
import io
import tempfile
import shutil
from pathlib import Path

# Configure Streamlit page
st.set_page_config(
    page_title="AI Template Mapper - Enhanced",
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
                            
                            # Debug: Print cell text for troubleshooting
                            if 'packaging' in cell_text or 'part' in cell_text:
                                print(f"Debug: Found potential section at {cell.coordinate}: '{cell_text}'")
                            
                            # Check for section keywords with more flexible matching
                            for section_name, section_info in self.section_mappings.items():
                                for keyword in section_info['section_keywords']:
                                    keyword_processed = self.preprocess_text(keyword)
                                    
                                    # Exact match
                                    if keyword_processed == cell_text:
                                        print(f"Debug: Exact section match - {section_name} at {cell.coordinate}")
                                        return section_name
                                    
                                    # Partial match for key phrases
                                    if keyword_processed in cell_text or cell_text in keyword_processed:
                                        print(f"Debug: Partial section match - {section_name} at {cell.coordinate}")
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
    
    def find_template_fields_with_context(self, template_file):
        """Find template fields with enhanced section context information"""
        fields = {}
        
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
            
            merged_ranges = worksheet.merged_cells.ranges
            
            print("Debug: Scanning template for fields...")
            
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
                                
                                print(f"Debug: Field '{cell_value}' at {cell_coord} -> Section: {section_context}")
                                
                                fields[cell_coord] = {
                                    'value': cell_value,
                                    'row': cell.row,
                                    'column': cell.column,
                                    'merged_range': merged_range,
                                    'section_context': section_context,
                                    'is_mappable': True
                                }
                    except Exception as e:
                        print(f"Error processing cell {cell.coordinate}: {e}")
                        continue
            
            workbook.close()
            print(f"Debug: Found {len(fields)} mappable fields")
            
        except Exception as e:
            st.error(f"Error reading template: {e}")
        
        return fields
    
    def map_data_with_section_context(self, template_fields, data_df):
        """Enhanced mapping with better section-aware logic"""
        mapping_results = {}
        
        try:
            data_columns = data_df.columns.tolist()
            print(f"Debug: Available data columns: {data_columns}")
            
            for coord, field in template_fields.items():
                try:
                    best_match = None
                    best_score = 0.0
                    field_value = field['value'].lower().strip()
                    section_context = field.get('section_context')
                    
                    print(f"Debug: Processing field '{field['value']}' in section '{section_context}'")
                    
                    # Try section-based mapping first
                    if section_context and section_context in self.section_mappings:
                        section_mappings = self.section_mappings[section_context]['field_mappings']
                        
                        # Look for direct field matches within section
                        for template_field_key, data_column_pattern in section_mappings.items():
                            if template_field_key in field_value or field_value in template_field_key:
                                print(f"Debug: Found section match for '{template_field_key}' -> '{data_column_pattern}'")
                                
                                # Look for exact match first
                                for data_col in data_columns:
                                    if data_column_pattern.lower() == data_col.lower():
                                        best_match = data_col
                                        best_score = 1.0
                                        print(f"Debug: Exact match found: {data_col}")
                                        break
                                
                                # If no exact match, try similarity matching
                                if not best_match:
                                    for data_col in data_columns:
                                        similarity = self.calculate_similarity(data_column_pattern, data_col)
                                        if similarity > best_score and similarity >= self.similarity_threshold:
                                            best_score = similarity
                                            best_match = data_col
                                            print(f"Debug: Similarity match found: {data_col} (score: {similarity:.2f})")
                                break
                    
                    # Fallback to general similarity matching
                    if not best_match:
                        print(f"Debug: Using fallback matching for '{field_value}'")
                        for data_col in data_columns:
                            similarity = self.calculate_similarity(field_value, data_col)
                            if similarity > best_score and similarity >= self.similarity_threshold:
                                best_score = similarity
                                best_match = data_col
                                print(f"Debug: Fallback match: {data_col} (score: {similarity:.2f})")
                    
                    mapping_results[coord] = {
                        'template_field': field['value'],
                        'data_column': best_match,
                        'similarity': best_score,
                        'field_info': field,
                        'section_context': section_context,
                        'is_mappable': best_match is not None
                    }
                    
                    print(f"Debug: Final mapping for '{field['value']}': {best_match} (score: {best_score:.2f})")
                        
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
    st.markdown("### Advanced packaging template processing with section-aware mapping")
    
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
    st.header("🚀 Enhanced AI Data Processor")
    st.info("Upload your packaging data and template. AI will intelligently map fields based on section context!")
    
    # Show enhanced mapping rules
    with st.expander("📋 Enhanced Section Mapping Rules", expanded=False):
        st.markdown("""
        **Primary Packaging Section:**
        - `Primary Packaging Type` ← `Primary Packaging Type`
        - `L-mm` ← `Primary L-mm`
        - `W-mm` ← `Primary W-mm`  
        - `H-mm` ← `Primary H-mm`
        - `Qty/Pack` ← `Primary Qty/Pack`
        - `Empty Weight` ← `Primary Empty Weight`
        - `Pack Weight` ← `Primary Pack Weight`
        
        **Secondary Packaging Section:**
        - `Secondary Packaging Type` ← `Secondary Packaging Type`
        - `L-mm` ← `Secondary L-mm`
        - `W-mm` ← `Secondary W-mm`
        - `H-mm` ← `Secondary H-mm`
        - `Qty/Pack` ← `Secondary Qty/Pack`
        - `Empty Weight` ← `Secondary Empty Weight`
        - `Pack Weight` ← `Secondary Pack Weight`
        
        **Part Information Section:**
        - `L` ← `Part L`
        - `W` ← `Part W`
        - `H` ← `Part H`
        - `Part No.` ← `Part No`
        - `Description` ← `Part Description`
        - `Unit Weight` ← `Part Unit Weight`
        """)
    
    # Data file upload
    data_file = st.file_uploader("Upload Data File", type=['csv', 'xlsx'])
    
    # Template selection
    if st.session_state.templates:
        selected_template = st.selectbox(
            "Select Template",
            options=list(st.session_state.templates.keys()),
            format_func=lambda x: f"{x} ({st.session_state.templates[x].get('type', 'Standard')})"
        )
    else:
        st.warning("No templates available. Please upload a template first.")
        return
    
    if data_file and selected_template:
        try:
            # Load data
            if data_file.name.lower().endswith('.csv'):
                data_df = pd.read_csv(data_file)
            else:
                data_df = pd.read_excel(data_file)
            
            st.subheader("📊 Data Preview")
            st.dataframe(data_df.head(), use_container_width=True)
            
            if st.button("🚀 Process with Enhanced AI", type="primary"):
                with st.spinner("🤖 Enhanced AI is analyzing sections and mapping fields..."):
                    # Get template info
                    template_info = st.session_state.templates[selected_template]
                    
                    # Create temporary template file
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                        tmp_file.write(template_info['file_data'])
                        template_path = tmp_file.name
                    
                    # Enhanced AI mapping with section context
                    template_fields = st.session_state.enhanced_mapper.find_template_fields_with_context(template_path)
                    mapping_results = st.session_state.enhanced_mapper.map_data_with_section_context(template_fields, data_df)
                    
                    # Fill template
                    filled_workbook, filled_count = st.session_state.enhanced_mapper.fill_template_with_data(
                        template_path, mapping_results, data_df
                    )
                    
                    os.unlink(template_path)
                
                if filled_workbook:
                    st.success(f"✅ Enhanced processing complete! Mapped {filled_count} fields with section awareness.")
                    
                    # Show enhanced mapping results
                    st.subheader("🎯 Section-Aware Mapping Results")
                    
                    # Group by section context
                    section_groups = {}
                    for mapping in mapping_results.values():
                        section = mapping.get('section_context', 'general')
                        if section not in section_groups:
                            section_groups[section] = {'mapped': [], 'unmapped': []}
                        
                        if mapping['is_mappable']:
                            section_groups[section]['mapped'].append(mapping)
                        else:
                            section_groups[section]['unmapped'].append(mapping)
                    
                    # Display results by section
                    for section, group in section_groups.items():
                        section_name = section.replace('_', ' ').title() if section else 'General'
                        with st.expander(f"📦 {section_name} Section", expanded=True):
                            col1, col2 = st.columns(2)
                            with col1:
                                st.write(f"**✅ Mapped ({len(group['mapped'])}):**")
                                for mapping in group['mapped']:
                                    confidence = mapping['similarity'] * 100
                                    st.write(f"• {mapping['template_field']} ← {mapping['data_column']} ({confidence:.1f}%)")
                            with col2:
                                st.write(f"**❌ Unmapped ({len(group['unmapped'])}):**")
                                for mapping in group['unmapped']:
                                    st.write(f"• {mapping['template_field']} (no match found)")
                    
                    # Save and download filled template
                    output = io.BytesIO()
                    filled_workbook.save(output)
                    output.seek(0)
                    filled_workbook.close()
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"filled_{selected_template}_{timestamp}.xlsx"
                    
                    st.download_button(
                        label="📥 Download Filled Template",
                        data=output.getvalue(),
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                else:
                    st.error("❌ Failed to process template. Please check the file format.")
        
        except Exception as e:
            st.error(f"❌ Error processing files: {str(e)}")

def show_template_manager():
    st.header("📂 Template Manager")
    
    if st.session_state.user_role == 'admin':
        st.subheader("➕ Upload New Template")
        
        template_name = st.text_input("Template Name")
        template_type = st.selectbox("Template Type", ["Standard", "Primary Packaging", "Secondary Packaging", "Part Information"])
        template_file = st.file_uploader("Upload Template File", type=['xlsx'])
        
        if st.button("Upload Template") and template_name and template_file:
            try:
                st.session_state.templates[template_name] = {
                    'type': template_type,
                    'file_data': template_file.read(),
                    'uploaded_by': st.session_state.username,
                    'upload_date': datetime.now().isoformat()
                }
                st.success(f"✅ Template '{template_name}' uploaded successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error uploading template: {str(e)}")
    
    # Show existing templates
    st.subheader("📋 Available Templates")
    if st.session_state.templates:
        for template_name, template_info in st.session_state.templates.items():
            with st.expander(f"📄 {template_name}", expanded=False):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"**Type:** {template_info.get('type', 'Standard')}")
                    st.write(f"**Uploaded by:** {template_info.get('uploaded_by', 'Unknown')}")
                    st.write(f"**Upload date:** {template_info.get('upload_date', 'Unknown')}")
                
                with col2:
                    if st.session_state.user_role == 'admin':
                        if st.button(f"🗑️ Delete", key=f"delete_{template_name}"):
                            del st.session_state.templates[template_name]
                            st.success(f"Template '{template_name}' deleted!")
                            st.rerun()
                
                # Download template
                st.download_button(
                    label="📥 Download Template",
                    data=template_info['file_data'],
                    file_name=f"{template_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{template_name}"
                )
    else:
        st.info("No templates available. Upload a template to get started.")

def show_mapping_analysis():
    st.header("📊 Mapping Analysis")
    
    if not st.session_state.templates:
        st.warning("No templates available for analysis.")
        return
    
    selected_template = st.selectbox(
        "Select Template for Analysis",
        options=list(st.session_state.templates.keys())
    )
    
    if selected_template:
        template_info = st.session_state.templates[selected_template]
        
        try:
            # Create temporary template file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(template_info['file_data'])
                template_path = tmp_file.name
            
            # Analyze template structure
            template_fields = st.session_state.enhanced_mapper.find_template_fields_with_context(template_path)
            os.unlink(template_path)
            
            st.subheader("🔍 Template Field Analysis")
            
            # Group fields by section
            section_analysis = {}
            for coord, field in template_fields.items():
                section = field.get('section_context', 'general')
                if section not in section_analysis:
                    section_analysis[section] = []
                section_analysis[section].append(field)
            
            # Display analysis
            for section, fields in section_analysis.items():
                section_name = section.replace('_', ' ').title() if section else 'General'
                with st.expander(f"📦 {section_name} Section - {len(fields)} fields", expanded=True):
                    fields_df = pd.DataFrame([
                        {
                            'Field': field['value'],
                            'Position': f"Row {field['row']}, Col {field['column']}",
                            'Mappable': '✅' if field['is_mappable'] else '❌',
                            'Merged': '✅' if field['merged_range'] else '❌'
                        }
                        for field in fields
                    ])
                    st.dataframe(fields_df, use_container_width=True)
            
            # Summary statistics
            st.subheader("📈 Summary Statistics")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Fields", len(template_fields))
            with col2:
                mappable_count = sum(1 for f in template_fields.values() if f['is_mappable'])
                st.metric("Mappable Fields", mappable_count)
            with col3:
                section_count = len(set(f.get('section_context', 'general') for f in template_fields.values()))
                st.metric("Sections Detected", section_count)
            with col4:
                merged_count = sum(1 for f in template_fields.values() if f['merged_range'])
                st.metric("Merged Cells", merged_count)
        
        except Exception as e:
            st.error(f"❌ Error analyzing template: {str(e)}")

def show_settings():
    st.header("⚙️ Settings")
    
    st.subheader("🎯 Mapping Configuration")
    
    # Similarity threshold
    new_threshold = st.slider(
        "Similarity Threshold",
        min_value=0.1,
        max_value=1.0,
        value=st.session_state.enhanced_mapper.similarity_threshold,
        step=0.05,
        help="Minimum similarity score for field mapping (higher = more strict)"
    )
    
    if new_threshold != st.session_state.enhanced_mapper.similarity_threshold:
        st.session_state.enhanced_mapper.similarity_threshold = new_threshold
        st.success(f"✅ Similarity threshold updated to {new_threshold}")
    
    # Advanced NLP status
    st.subheader("🧠 AI Capabilities")
    col1, col2 = st.columns(2)
    
    with col1:
        nlp_status = "🟢 Enabled" if ADVANCED_NLP else "🔴 Disabled"
        st.write(f"**Advanced NLP:** {nlp_status}")
        
        nltk_status = "🟢 Ready" if NLTK_READY else "🔴 Not Ready"
        st.write(f"**NLTK Library:** {nltk_status}")
    
    with col2:
        if not ADVANCED_NLP:
            st.info("💡 Install `nltk` and `scikit-learn` for enhanced text matching capabilities.")
    
    # System information
    st.subheader("ℹ️ System Information")
    info_df = pd.DataFrame([
        {"Component": "Python Libraries", "Status": "sklearn, nltk" if ADVANCED_NLP else "Basic only"},
        {"Component": "Text Processing", "Status": "Advanced" if ADVANCED_NLP else "Basic"},
        {"Component": "Section Detection", "Status": "Enhanced" if ADVANCED_NLP else "Standard"},
        {"Component": "Template Storage", "Status": f"{len(st.session_state.templates)} templates"}
    ])
    st.dataframe(info_df, use_container_width=True)

def main():
    if not st.session_state.authenticated:
        show_login()
        return
    
    # Sidebar navigation
    with st.sidebar:
        st.write(f"👋 Welcome, **{st.session_state.name}**")
        st.write(f"Role: {st.session_state.user_role.title()}")
        
        st.divider()
        
        page = st.radio(
            "Navigation",
            ["🚀 Enhanced Processor", "📂 Template Manager", "📊 Mapping Analysis", "⚙️ Settings"],
            key="navigation"
        )
        
        st.divider()
        
        if st.button("🚪 Logout"):
            st.session_state.authenticated = False
            st.session_state.user_role = None
            st.rerun()
        
        # System status
        st.write("**System Status:**")
        st.write(f"🧠 AI: {'Advanced' if ADVANCED_NLP else 'Basic'}")
        st.write(f"📋 Templates: {len(st.session_state.templates)}")
    
    # Main content
    if page == "🚀 Enhanced Processor":
        show_enhanced_processor()
    elif page == "📂 Template Manager":
        show_template_manager()
    elif page == "📊 Mapping Analysis":
        show_mapping_analysis()
    elif page == "⚙️ Settings":
        show_settings()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"❌ Application Error: {str(e)}")
        st.write("Please refresh the page or contact support if the issue persists.")
