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
        """FIXED: Improved image extraction from Excel template"""
        images_info = {}
        
        try:
            # Load workbook without data_only to preserve images
            workbook = openpyxl.load_workbook(template_file, data_only=False)
            worksheet = workbook.active
            
            st.info(f"🔍 Analyzing worksheet: {worksheet.title}")
            
            # Method 1: Extract embedded images from worksheet._images
            if hasattr(worksheet, '_images') and worksheet._images:
                st.success(f"✅ Found {len(worksheet._images)} embedded images using standard method")
                
                for i, img in enumerate(worksheet._images):
                    try:
                        # Get image anchor information
                        anchor = img.anchor
                        row, col = 1, 1  # Default position
                        
                        # Handle different anchor types
                        if hasattr(anchor, '_from') and anchor._from:
                            row = getattr(anchor._from, 'row', 0) + 1
                            col = getattr(anchor._from, 'col', 0) + 1
                        elif hasattr(anchor, 'row') and hasattr(anchor, 'col'):
                            row = anchor.row + 1
                            col = anchor.col + 1
                        
                        # Determine section context
                        section_context = self.identify_section_context(worksheet, row, col, max_search_rows=30)
                        
                        # Generate descriptive name
                        if section_context:
                            section_name = section_context.replace('_', ' ').title()
                            image_name = f"{section_name}_Image_{i+1}"
                        else:
                            image_name = f"Template_Image_{i+1}"
                        
                        # Extract image data - FIXED METHOD
                        try:
                            # Get image data using proper method
                            if hasattr(img, '_data') and callable(img._data):
                                img_data = img._data()
                            elif hasattr(img, 'ref'):
                                img_data = img.ref
                            elif hasattr(img, '_image'):
                                img_data = img._image
                            else:
                                img_data = None
                            
                            images_info[image_name] = {
                                'data': img_data,
                                'position': f"Row {row}, Col {col}",
                                'section_context': section_context or 'unknown',
                                'anchor': str(anchor),
                                'image_index': i,
                                'extraction_method': 'embedded_standard',
                                'size': getattr(img, 'width', 0) if hasattr(img, 'width') else 0
                            }
                            
                        except Exception as e:
                            st.warning(f"⚠️ Could not extract data for image {i}: {e}")
                            # Still record the image existence
                            images_info[image_name] = {
                                'data': None,
                                'position': f"Row {row}, Col {col}",
                                'section_context': section_context or 'unknown',
                                'anchor': str(anchor),
                                'image_index': i,
                                'extraction_method': 'embedded_metadata_only',
                                'error': str(e)
                            }
                            
                    except Exception as e:
                        st.error(f"❌ Error processing image {i}: {e}")
                        continue
            else:
                st.info("ℹ️ No embedded images found in worksheet._images")
            
            # Method 2: Check worksheet drawing parts (alternative extraction)
            try:
                if hasattr(worksheet, '_drawing') and worksheet._drawing:
                    drawing = worksheet._drawing
                    st.info("🔍 Checking worksheet drawings...")
                    
                    # Check for charts or other drawing objects
                    if hasattr(drawing, 'charts') and drawing.charts:
                        st.info(f"Found {len(drawing.charts)} chart objects")
                    
                    # Look for image relationships in drawing
                    if hasattr(drawing, '_rels') and drawing._rels:
                        st.info(f"Found {len(drawing._rels)} drawing relationships")
                        for rel_id, rel in drawing._rels.items():
                            if 'image' in rel.target.lower():
                                st.info(f"Found image relationship: {rel.target}")
                                
            except Exception as e:
                st.warning(f"⚠️ Drawing analysis failed: {e}")
            
            # Method 3: Scan for image placeholders and references
            try:
                st.info("🔍 Scanning for image placeholders...")
                placeholder_count = 0
                
                for row in range(1, min(worksheet.max_row + 1, 100)):
                    for col in range(1, min(worksheet.max_column + 1, 50)):
                        cell = worksheet.cell(row=row, column=col)
                        if cell.value:
                            cell_text = str(cell.value).lower()
                            
                            # Look for image-related keywords
                            image_keywords = ['image', 'picture', 'photo', 'current packaging', 
                                            'reference', 'attach', 'insert', 'paste image']
                            
                            if any(keyword in cell_text for keyword in image_keywords):
                                section_context = self.identify_section_context(worksheet, row, col)
                                
                                placeholder_name = f"Placeholder_{section_context or 'Unknown'}_{row}_{col}"
                                
                                # Avoid duplicates
                                if placeholder_name not in images_info:
                                    images_info[placeholder_name] = {
                                        'data': None,
                                        'position': f"Row {row}, Col {col}",
                                        'section_context': section_context or 'unknown',
                                        'anchor': None,
                                        'image_index': -1,
                                        'extraction_method': 'text_placeholder',
                                        'cell_text': cell_text,
                                        'placeholder': True
                                    }
                                    placeholder_count += 1
                
                if placeholder_count > 0:
                    st.info(f"📝 Found {placeholder_count} image placeholders")
                    
            except Exception as e:
                st.warning(f"⚠️ Placeholder detection failed: {e}")
            
            # Method 4: Check for merged cells that might contain image areas
            try:
                merged_ranges = list(worksheet.merged_cells.ranges)
                if merged_ranges:
                    st.info(f"🔍 Checking {len(merged_ranges)} merged cell ranges for image areas...")
                    
                    for i, merged_range in enumerate(merged_ranges):
                        # Check if merged area is large enough to contain an image
                        min_col, min_row, max_col, max_row = merged_range.bounds
                        width = max_col - min_col + 1
                        height = max_row - min_row + 1
                        
                        # Consider it a potential image area if it's reasonably sized
                        if width >= 3 and height >= 3:
                            section_context = self.identify_section_context(worksheet, min_row, min_col)
                            
                            area_name = f"Merged_Area_{section_context or 'Unknown'}_{i+1}"
                            
                            images_info[area_name] = {
                                'data': None,
                                'position': f"Rows {min_row}-{max_row}, Cols {min_col}-{max_col}",
                                'section_context': section_context or 'unknown',
                                'anchor': str(merged_range),
                                'image_index': -1,
                                'extraction_method': 'merged_cell_area',
                                'dimensions': f"{width}x{height} cells",
                                'bounds': merged_range.bounds
                            }
                            
            except Exception as e:
                st.warning(f"⚠️ Merged cell analysis failed: {e}")
            
            workbook.close()
            
            # Summary
            if images_info:
                embedded_count = len([img for img in images_info.values() if img['extraction_method'].startswith('embedded')])
                placeholder_count = len([img for img in images_info.values() if img['extraction_method'] == 'text_placeholder'])
                merged_count = len([img for img in images_info.values() if img['extraction_method'] == 'merged_cell_area'])
                
                st.success(f"📊 Image extraction summary: {embedded_count} embedded, {placeholder_count} placeholders, {merged_count} merged areas")
            else:
                st.warning("⚠️ No images or image areas found in template")
            
        except Exception as e:
            st.error(f"❌ Error during image extraction: {e}")
        
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
        """FIXED: Find template fields with enhanced section context information"""
        fields = {}
        
        try:
            workbook = openpyxl.load_workbook(template_file, data_only=False)
            worksheet = workbook.active
            
            merged_ranges = list(worksheet.merged_cells.ranges)
            
            st.info(f"🔍 Scanning worksheet with {worksheet.max_row} rows and {worksheet.max_column} columns")
            
            field_count = 0
            for row in worksheet.iter_rows():
                for cell in row:
                    try:
                        if cell.value is not None:
                            cell_value = str(cell.value).strip()
                            
                            if cell_value and self.is_mappable_field(cell_value):
                                cell_coord = cell.coordinate
                                merged_range = None
                                
                                # Check if cell is part of merged range
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
                                    'is_mappable': True,
                                    'clean_value': self.preprocess_text(cell_value)
                                }
                                field_count += 1
                                
                    except Exception as e:
                        continue
            
            st.success(f"✅ Found {field_count} mappable template fields")
            workbook.close()
            
        except Exception as e:
            st.error(f"❌ Error reading template: {e}")
        
        return fields
    
    def map_data_with_section_context(self, template_fields, data_df):
        """Enhanced mapping with better section-aware logic"""
        mapping_results = {}
        
        try:
            data_columns = data_df.columns.tolist()
            st.info(f"🔄 Mapping {len(template_fields)} template fields to {len(data_columns)} data columns")
            
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
                    st.error(f"❌ Error mapping field {coord}: {e}")
                    continue
            
            mapped_count = sum(1 for m in mapping_results.values() if m['is_mappable'])
            st.success(f"✅ Successfully mapped {mapped_count}/{len(mapping_results)} fields")
                    
        except Exception as e:
            st.error(f"❌ Error in map_data_with_section_context: {e}")
            
        return mapping_results
    
    def find_data_cell_for_label(self, worksheet, field_info):
        """FIXED: Find data cell for a label with improved logic"""
        try:
            row = field_info['row']
            col = field_info['column']
            merged_ranges = list(worksheet.merged_cells.ranges)
        
            def is_suitable_data_cell(target_row, target_col):
                """Check if a cell position is suitable for data entry"""
                try:
                    if target_row <= 0 or target_col <= 0:
                        return False, None
                    if target_row > worksheet.max_row or target_col > worksheet.max_column:
                        return False, None
                        
                    cell = worksheet.cell(row=target_row, column=target_col)
                    cell_coord = cell.coordinate
                    
                    # Skip if it's a merged cell (not the anchor)
                    for merged_range in merged_ranges:
                        if cell_coord in merged_range and cell_coord != merged_range.start_cell.coordinate:
                            return False, None
                    
                    # Check cell content
                    if cell.value is None or str(cell.value).strip() == "":
                        return True, cell_coord
                    
                    # Check for data placeholder patterns
                    cell_text = str(cell.value).lower().strip()
                    data_patterns = [
                        r'^_+$', r'^\.*$', r'^-+$', r'^\s*$',  # Empty patterns
                        r'enter', r'fill', r'data', r'value',  # Placeholder text
                        r'insert', r'type', r'add'
                    ]
                    
                    for pattern in data_patterns:
                        if re.search(pattern, cell_text):
                            return True, cell_coord
                    
                    # If cell has non-placeholder content, it's not suitable
                    return False, None
                    
                except Exception as e:
                    return False, None
            
            # Strategy 1: Look immediately to the right (most common pattern)
            for offset in range(1, 6):
                is_suitable, cell_coord = is_suitable_data_cell(row, col + offset)
                if is_suitable:
                    return cell_coord
            
            # Strategy 2: Look below the label
            for offset in range(1, 4):
                is_suitable, cell_coord = is_suitable_data_cell(row + offset, col)
                if is_suitable:
                    return cell_coord
            
            # Strategy 3: Look diagonally (right-down)
            for offset in range(1, 3):
                is_suitable, cell_coord = is_suitable_data_cell(row + offset, col + offset)
                if is_suitable:
                    return cell_coord
            
            # Strategy 4: Look in nearby cells (broader search)
            for row_offset in range(-1, 3):
                for col_offset in range(-1, 4):
                    if row_offset == 0 and col_offset == 0:
                        continue
                    is_suitable, cell_coord = is_suitable_data_cell(row + row_offset, col + col_offset)
                    if is_suitable:
                        return cell_coord
            
            # Default fallback: return adjacent cell
            return worksheet.cell(row=row, column=col + 1).coordinate
            
        except Exception as e:
            st.error(f"Error in find_data_cell_for_label: {e}")
            return None
    
    def populate_template_with_data(self, template_file, data_df, mapping_results, images_info):
        """FIXED: Populate template with data and images"""
        try:
            # Create a copy of the template
            temp_dir = tempfile.mkdtemp()
            temp_template_path = os.path.join(temp_dir, "populated_template.xlsx")
            shutil.copy2(template_file, temp_template_path)
            
            # Load the copied template
            workbook = openpyxl.load_workbook(temp_template_path)
            worksheet = workbook.active
            
            # Track populated fields
            populated_count = 0
            
            st.info("🔄 Populating template with data...")
            
            # Populate data fields
            for coord, mapping in mapping_results.items():
                try:
                    if mapping['is_mappable'] and mapping['data_column']:
                        field_info = mapping['field_info']
                        data_column = mapping['data_column']
                        
                        # Get data value (first row of data)
                        if not data_df.empty and data_column in data_df.columns:
                            data_value = data_df[data_column].iloc[0]
                            
                            if pd.notna(data_value):
                                # Find appropriate data cell
                                data_cell_coord = self.find_data_cell_for_label(worksheet, field_info)
                                
                                if data_cell_coord:
                                    data_cell = worksheet[data_cell_coord]
                                    data_cell.value = str(data_value)
                                    
                                    # Apply basic formatting
                                    data_cell.font = Font(size=10)
                                    data_cell.alignment = Alignment(horizontal='left', vertical='center')
                                    
                                    populated_count += 1
                                    
                except Exception as e:
                    st.warning(f"⚠️ Could not populate field {coord}: {e}")
                    continue
            
            # Handle images
            image_count = 0
            for image_name, image_info in images_info.items():
                try:
                    if image_info.get('data') and image_info.get('section_context'):
                        # This is where you would add image insertion logic
                        # For now, we'll add a placeholder text
                        section = image_info['section_context']
                        
                        # Find a suitable cell for image placeholder
                        if 'position' in image_info:
                            position_info = image_info['position']
                            # Extract row/col from position string
                            if 'Row' in position_info and 'Col' in position_info:
                                try:
                                    row_match = re.search(r'Row (\d+)', position_info)
                                    col_match = re.search(r'Col (\d+)', position_info)
                                    
                                    if row_match and col_match:
                                        target_row = int(row_match.group(1))
                                        target_col = int(col_match.group(1))
                                        
                                        cell = worksheet.cell(row=target_row, column=target_col)
                                        cell.value = f"[{image_name}]"
                                        cell.font = Font(italic=True, color="0000FF")
                                        image_count += 1
                                        
                                except Exception as e:
                                    continue
                        
                except Exception as e:
                    st.warning(f"⚠️ Could not process image {image_name}: {e}")
                    continue
            
            # Save the populated template
            workbook.save(temp_template_path)
            workbook.close()
            
            st.success(f"✅ Template populated: {populated_count} fields, {image_count} image placeholders")
            
            return temp_template_path
            
        except Exception as e:
            st.error(f"❌ Error populating template: {e}")
            return None

def create_download_link(file_path, filename):
    """Create a download link for the populated template"""
    try:
        with open(file_path, "rb") as file:
            file_data = file.read()
        
        b64_data = base64.b64encode(file_data).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_data}" download="{filename}">📥 Download Populated Template</a>'
        return href
    except Exception as e:
        st.error(f"Error creating download link: {e}")
        return None

def main():
    """Main Streamlit application"""
    st.title("🤖 AI Template Mapper - Enhanced Image Extraction")
    st.markdown("Upload your Excel template and data file to automatically populate the template with intelligent field mapping.")
    
    # Initialize the mapper
    mapper = EnhancedTemplateMapper()
    
    # Sidebar configuration
    st.sidebar.header("⚙️ Configuration")
    mapper.similarity_threshold = st.sidebar.slider(
        "Similarity Threshold", 
        min_value=0.1, 
        max_value=1.0, 
        value=0.3, 
        step=0.1,
        help="Minimum similarity score for field matching"
    )
    
    # File uploads
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Template File")
        template_file = st.file_uploader(
            "Upload Excel Template", 
            type=['xlsx', 'xls'],
            help="Upload your Excel template with fields to be populated"
        )
    
    with col2:
        st.subheader("📊 Data File")
        data_file = st.file_uploader(
            "Upload Data File", 
            type=['xlsx', 'xls', 'csv'],
            help="Upload your data file with values to populate the template"
        )
    
    if template_file and data_file:
        try:
            # Save uploaded files temporarily
            temp_dir = tempfile.mkdtemp()
            template_path = os.path.join(temp_dir, template_file.name)
            data_path = os.path.join(temp_dir, data_file.name)
            
            with open(template_path, "wb") as f:
                f.write(template_file.getbuffer())
            
            with open(data_path, "wb") as f:
                f.write(data_file.getbuffer())
            
            # Process the files
            st.header("🔍 Analysis Results")
            
            # Extract images from template
            with st.expander("🖼️ Image Extraction Results", expanded=True):
                images_info = mapper.extract_images_from_template_improved(template_path)
                
                if images_info:
                    st.write(f"**Found {len(images_info)} image elements:**")
                    
                    for img_name, img_info in images_info.items():
                        with st.container():
                            col1, col2, col3 = st.columns([2, 2, 2])
                            
                            with col1:
                                st.write(f"**{img_name}**")
                                st.write(f"Position: {img_info.get('position', 'Unknown')}")
                            
                            with col2:
                                st.write(f"Section: {img_info.get('section_context', 'Unknown')}")
                                st.write(f"Method: {img_info.get('extraction_method', 'Unknown')}")
                            
                            with col3:
                                if img_info.get('data'):
                                    st.success("✅ Data Available")
                                else:
                                    st.info("ℹ️ Placeholder/Reference")
                else:
                    st.info("No images found in template")
            
            # Find template fields
            st.subheader("📋 Template Field Analysis")
            template_fields = mapper.find_template_fields_with_context(template_path)
            
            if template_fields:
                # Display fields by section
                sections = {}
                for coord, field in template_fields.items():
                    section = field.get('section_context', 'Unknown')
                    if section not in sections:
                        sections[section] = []
                    sections[section].append((coord, field))
                
                for section_name, section_fields in sections.items():
                    with st.expander(f"📁 {section_name.replace('_', ' ').title()} ({len(section_fields)} fields)"):
                        for coord, field in section_fields:
                            st.write(f"**{coord}**: {field['value']}")
            
            # Load and process data
            st.subheader("📊 Data File Analysis")
            if data_file.name.endswith('.csv'):
                data_df = pd.read_csv(data_path)
            else:
                data_df = pd.read_excel(data_path)
            
            st.write(f"**Data shape:** {data_df.shape[0]} rows × {data_df.shape[1]} columns")
            st.write("**Columns:**", ", ".join(data_df.columns.tolist()))
            
            # Show data preview
            with st.expander("👀 Data Preview"):
                st.dataframe(data_df.head())
            
            # Perform mapping
            st.subheader("🔗 Field Mapping Results")
            mapping_results = mapper.map_data_with_section_context(template_fields, data_df)
            
            # Display mapping results
            mapped_fields = [m for m in mapping_results.values() if m['is_mappable']]
            unmapped_fields = [m for m in mapping_results.values() if not m['is_mappable']]
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("✅ Mapped Fields", len(mapped_fields))
                if mapped_fields:
                    with st.expander("View Mapped Fields"):
                        for mapping in mapped_fields:
                            st.write(f"**{mapping['template_field']}** → **{mapping['data_column']}** ({mapping['similarity']:.2f})")
            
            with col2:
                st.metric("❌ Unmapped Fields", len(unmapped_fields))
                if unmapped_fields:
                    with st.expander("View Unmapped Fields"):
                        for mapping in unmapped_fields:
                            st.write(f"**{mapping['template_field']}** (no match found)")
            
            # Generate populated template
            if st.button("🚀 Generate Populated Template", type="primary"):
                with st.spinner("Generating populated template..."):
                    populated_template_path = mapper.populate_template_with_data(
                        template_path, data_df, mapping_results, images_info
                    )
                    
                    if populated_template_path:
                        st.success("✅ Template populated successfully!")
                        
                        # Create download link
                        download_link = create_download_link(
                            populated_template_path, 
                            f"populated_{template_file.name}"
                        )
                        
                        if download_link:
                            st.markdown(download_link, unsafe_allow_html=True)
                            
                        # Show summary
                        st.info("📋 **Summary:** The template has been populated with your data. Download the file above to view the results.")
            
            # Cleanup
            shutil.rmtree(temp_dir, ignore_errors=True)
            
        except Exception as e:
            st.error(f"❌ Error processing files: {e}")
            st.exception(e)
    
    else:
        st.info("👆 Please upload both template and data files to begin analysis.")
        
        # Show feature information
        st.markdown("""
        ### ✨ Features
        
        - **🔍 Smart Field Detection**: Automatically identifies mappable fields in Excel templates
        - **🧠 Intelligent Matching**: Uses advanced NLP techniques for field-to-column mapping
        - **📁 Section-Aware Processing**: Understands template structure and context
        - **🖼️ Enhanced Image Extraction**: Detects and processes embedded images and placeholders
        - **⚙️ Configurable Similarity**: Adjustable matching threshold for optimal results
        - **📥 Easy Download**: Get your populated template with one click
        
        ### 📋 Supported File Types
        - **Templates**: Excel files (.xlsx, .xls)
        - **Data**: Excel files (.xlsx, .xls) and CSV files (.csv)
        """)

if __name__ == "__main__":
    main()
