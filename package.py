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
from collections import defaultdict
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
    """Handles image extraction from Excel files with improved duplicate handling"""
    
    def __init__(self):
        self.supported_formats = ['.png', '.jpg', '.jpeg', '.gif', '.bmp']
    
    def identify_image_upload_areas(self, worksheet):
        """Identify areas in template designated for image uploads"""
        upload_areas = []
        processed_cells = set()
        
        try:
            # Enhanced image keywords with better matching
            image_keywords = {
                'primary': ['primary packaging', 'primary pack', 'primary'],
                'secondary': ['secondary packaging', 'secondary pack', 'secondary', 'sec'],
                'current': ['current packaging', 'current', 'existing packaging', 'existing'],
                'label': ['label', 'product label', 'labeling', 'labels']
            }
            
            # Search through first few rows for column headers
            for row_num in range(1, min(15, worksheet.max_row + 1)):
                for col_num in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=row_num, column=col_num)
                    cell_coord = f"{row_num}_{col_num}"
                    
                    if cell_coord in processed_cells or not cell.value:
                        continue
                        
                    cell_text = str(cell.value).lower().strip()
                    
                    # Find the best matching category
                    best_match = None
                    best_score = 0
                    
                    for category, keywords in image_keywords.items():
                        for keyword in keywords:
                            # Calculate match score based on keyword length and position
                            if keyword in cell_text:
                                score = len(keyword)
                                # Bonus for exact match
                                if keyword == cell_text:
                                    score += 10
                                # Bonus for word boundary match
                                if cell_text.startswith(keyword) or cell_text.endswith(keyword):
                                    score += 5
                                
                                if score > best_score:
                                    best_match = (category, keyword)
                                    best_score = score
                    
                    if best_match:
                        upload_areas.append({
                            'position': cell.coordinate,
                            'row': row_num + 1,  # Position image below header
                            'column': col_num,
                            'text': cell.value,
                            'type': best_match[0],
                            'header_text': cell_text,
                            'matched_keyword': best_match[1],
                            'match_score': best_score
                        })
                        processed_cells.add(cell_coord)
                        print(f"Found {best_match[0]} area at {cell.coordinate}: '{cell.value}' (score: {best_score})")

            # Sort by match score (best matches first)
            upload_areas.sort(key=lambda x: x['match_score'], reverse=True)
            return upload_areas
            
        except Exception as e:
            st.error(f"Error identifying image upload areas: {e}")
            return []

    def extract_images_from_excel(self, excel_file_path):
        """Extract unique images from Excel file with better deduplication"""
        try:
            images = {}
            workbook = openpyxl.load_workbook(excel_file_path)
            image_hashes = set()  # Track image hashes to avoid duplicates
            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
                print(f"Processing sheet: {sheet_name}")

                if hasattr(worksheet, '_images') and worksheet._images:
                    print(f"Found {len(worksheet._images)} images in {sheet_name}")

                    for idx, img in enumerate(worksheet._images):
                        try:
                            # Get image data
                            image_data = img._data()

                            # Create hash of image data to detect duplicates
                            image_hash = hashlib.md5(image_data).hexdigest()

                            if image_hash in image_hashes:
                                print(f"Skipping duplicate image in {sheet_name}")
                                continue
                            image_hashes.add(image_hash)

                            # Create PIL Image
                            pil_image = Image.open(io.BytesIO(image_data))

                            # Get image position
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

                            # Classify image type from sheet name
                            sheet_name_lc = sheet_name.lower()
                            if any(word in sheet_name_lc for word in ['primary', 'sheet1', 'internal']):
                                image_type = 'primary'
                            elif any(word in sheet_name_lc for word in ['secondary', 'sheet2', 'external']):
                                image_type = 'secondary'
                            elif any(word in sheet_name_lc for word in ['current', 'sheet3', 'existing']):
                                image_type = 'current'
                            elif any(word in sheet_name_lc for word in ['label', 'sheet4']):
                                image_type = 'label'
                            else:
                                image_type = f'image_{len(images) + 1}'
                            image_key = f"{image_type}_{sheet_name}_{position}"
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

                            print(f"Extracted unique image: {image_key} at position {position}")
                        except Exception as e:
                            print(f"Error extracting image {idx} from {sheet_name}: {e}")
                            continue
                else:
                    print(f"No images found in sheet: {sheet_name}")
            workbook.close()
            print(f"Total unique extracted images: {len(images)}")
            return {'all_sheets': images}
        except Exception as e:
            st.error(f"Error extracting images: {e}")
            print(f"Error in extract_images_from_excel: {e}")
            return {}

    def add_images_to_template(self, worksheet, uploaded_images, image_areas):
        """Add uploaded images to template with improved matching and proper placement"""
        try:
            added_images = 0
            temp_image_paths = []
            used_images = set()

            # Sort image areas by match score and column
            image_areas_sorted = sorted(image_areas, key=lambda x: (x.get('match_score', 0), x['column']), reverse=True)

            print(f"Processing {len(image_areas_sorted)} image areas")
            print(f"Available images: {list(uploaded_images.keys())}")

            # Initialize placement counters for grid-based types
            if not hasattr(self, "_placement_counters"):
                self._placement_counters = defaultdict(int)
            def cell_range(start_col, end_col, start_row, end_row):
                return [
                    (column_index_from_string(col), row)
                    for col in range(ord(start_col), ord(end_col)+1)
                    for row in range(start_row, end_row+1)
                ]

            # Define cell grids (defined outside loop to avoid scope issues)
            placement_grids = {
                'primary': [(column_index_from_string(col), row) for col in ['A','B','C','D'] for row in range(41, 49)],     # A–D, rows 44–51
                'secondary': [(column_index_from_string(col), row) for col in ['F','G','H','I','J'] for row in range(41, 49)],  # E–K, rows 44–51
                'label': [(column_index_from_string(col), row) for col in ['K','L','M','N','O','P'] for row in range(41, 49)]      # M–R, rows 44–51
            }

            for area in image_areas_sorted:
                area_type = area['type']
                header_text = area.get('header_text', '').lower()

                print(f"Processing area: {area_type} at {area['position']} (score: {area.get('match_score', 0)})")

                matching_image = None
                matching_key = None

                # Match image by type
                for img_key, img_data in uploaded_images.items():
                    if img_key in used_images:
                        continue
                    img_type = img_data.get('type', '').lower()
                    img_key_lower = img_key.lower()

                    if area_type == img_type:
                        matching_image = img_data
                        matching_key = img_key
                        break
                # Fallback to first available
                if not matching_image:
                    for img_key, img_data in uploaded_images.items():
                        if img_key not in used_images:
                            matching_image = img_data
                            matching_key = img_key
                            break

                if matching_image and matching_key:
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                            image_bytes = base64.b64decode(matching_image['data'])
                            tmp_img.write(image_bytes)
                            tmp_img_path = tmp_img.name
                        img = OpenpyxlImage(tmp_img_path)

                        # Set image size
                        width_cm, height_cm = (8.3, 8.3) if area_type == 'current' else (4.3, 4.3)
                        img.width = int(width_cm * 37.8)
                        img.height = int(height_cm * 37.8)

                        # Placement logic
                        if area_type in placement_grids:
                            index = self._placement_counters[area_type]
                            grid = placement_grids[area_type]
                            if index < len(grid):
                                col, row = grid[index]
                                cell_coord = f"{get_column_letter(col)}{row}"
                                self._placement_counters[area_type] += 1
                            else:
                                print(f"⚠️ No more space left for {area_type} images. Skipping.")
                                continue
                        elif area_type == 'current':
                            current_area = next((a for a in image_areas if a['type'] == 'current'), None)
                            if current_area:
                                cell_coord = f"{get_column_letter(current_area['column'])}{current_area['row']}"
                            else:
                                print("⚠️ No valid header found for current packaging image. Skipping.")
                                continue
                        else:
                            cell_coord = "A1"  # fallback
                        img.anchor = cell_coord
                        worksheet.add_image(img)

                        temp_image_paths.append(tmp_img_path)
                        used_images.add(matching_key)
                        added_images += 1

                        print(f"✅ Added {area_type} image at {cell_coord} ({width_cm}x{height_cm} cm)")

                    except Exception as e:
                        print(f"❌ Could not add {area_type} image: {e}")
                        st.warning(f"Could not add {area_type} image: {e}")
                        continue
                else:
                    print(f"⚠️ No matching image found for {area_type} at {area['position']}")
            print(f"Total images added: {added_images}")
            return added_images, temp_image_paths
        except Exception as e:
            st.error(f"Error adding images to template: {e}")
            print(f"Error in add_images_to_template: {e}")
            return 0, []

    def process_extracted_images_for_template(self, extracted_images):
        """Convert extracted images to format expected by template filling"""
        processed_images = {}
        if extracted_images and 'all_sheets' in extracted_images:
            for image_key, img_data in extracted_images['all_sheets'].items():
                processed_images[image_key] = img_data
        return processed_images
        
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
            temp_image_paths = []
            used_images = set()
            for area in image_areas:
                area_type = area['type']
                label_text = area.get('text', '').lower()  # ✅ Define label_text here

                matching_image = None

                for label, img_data in uploaded_images.items():
                    if label in used_images:
                        continue
                    label_lower = label.lower()
                    if (
                        area_type in label_lower
                        or area_type.replace('_', ' ') in label_lower
                        or 'primary' in label_lower and area_type == 'primary_packaging'
                        or 'secondary' in label_lower and area_type == 'secondary_packaging'
                        or 'current' in label_lower and area_type == 'current_packaging'
                        or label_lower in label_text
                        or label_text in label_lower
                    ):
                        matching_image = img_data
                        used_images.add(label)
                        break
                        
                if not matching_image:
                    for label, img_data in uploaded_images.items():
                        if label not in used_images:
                            matching_image = img_data
                            used_images.add(label)
                            break
                if matching_image:
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                            image_bytes = base64.b64decode(matching_image['data'])
                            tmp_img.write(image_bytes)
                            tmp_img_path = tmp_img.name
                        img = OpenpyxlImage(tmp_img_path)
                        img.width = 250
                        img.height = 150

                        cell_coord = f"{get_column_letter(area['column'])}{area['row']}"
                        worksheet.add_image(img, cell_coord)

                        temp_image_paths.append(tmp_img_path)
                        added_images += 1
                    except Exception as e:
                        st.warning(f"Could not add image to {area['position']}: {e}")
                        continue
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
                # First, identify image upload areas
                _, image_areas = self.find_template_fields_with_context_and_images(template_file)
                images_added, temp_image_paths = self.image_extractor.add_images_to_template(worksheet, uploaded_images, image_areas)
                
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
    st.title("🤖 Enhanced AI Template Mapper with Images")
    
    # Header with user info
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown(f"Welcome, **{st.session_state.name}** ({st.session_state.user_role})")
    with col3:
        if st.button("Logout"):
            st.session_state.authenticated = False
            st.session_state.user_role = None
            st.rerun()
    
    st.markdown("---")
    
    # Sidebar for file uploads
    with st.sidebar:
        st.header("📁 File Upload")
        
        # Template upload
        st.subheader("Excel Template")
        template_file = st.file_uploader(
            "Upload Excel Template",
            type=['xlsx', 'xls'],
            help="Upload the Excel template file"
        )
        
        # Data upload
        st.subheader("Data File")
        data_file = st.file_uploader(
            "Upload Data File",
            type=['xlsx', 'xls', 'csv'],
            help="Upload the data file to map to template (images will be extracted from Excel files)"
        )
        
        # Settings
        st.subheader("⚙️ Settings")
        similarity_threshold = st.slider(
            "Similarity Threshold",
            min_value=0.1,
            max_value=1.0,
            value=0.3,
            step=0.1,
            help="Minimum similarity score for field matching"
        )
        
        st.session_state.enhanced_mapper.similarity_threshold = similarity_threshold
    
    # Main content area
    if template_file and data_file:
        try:
            # Extract images from data file (only if it's Excel)
            extracted_images = {}
            if data_file.name.endswith(('.xlsx', '.xls')):
                # Create temporary file for data file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_data:
                    tmp_data.write(data_file.getvalue())
                    data_path = tmp_data.name
                
                st.info("🔍 Extracting images from data file...")
                with st.spinner("Extracting images from Excel file..."):
                    extracted_images = st.session_state.enhanced_mapper.image_extractor.extract_images_from_excel(data_path)
                
                # Clean up data file copy
                try:
                    os.unlink(data_path)
                except:
                    pass
            
            # Read data file
            if data_file.name.endswith('.csv'):
                data_df = pd.read_csv(data_file)
            else:
                data_df = pd.read_excel(data_file)
            
            # Create temporary file for template
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_template:
                tmp_template.write(template_file.getvalue())
                template_path = tmp_template.name
            
            # Process template and find fields
            st.subheader("📋 Template Analysis")
            
            with st.spinner("Analyzing template fields and image areas..."):
                template_fields, image_areas = st.session_state.enhanced_mapper.find_template_fields_with_context_and_images(template_path)
            
            if template_fields:
                st.success(f"Found {len(template_fields)} mappable fields")
                
                # Show template fields
                with st.expander("Template Fields Details", expanded=False):
                    fields_df = pd.DataFrame([
                        {
                            'Position': coord,
                            'Field': field['value'],
                            'Section': field.get('section_context', 'Unknown'),
                            'Row': field['row'],
                            'Column': field['column']
                        }
                        for coord, field in template_fields.items()
                    ])
                    st.dataframe(fields_df, use_container_width=True)
                
                # Show image areas
                if image_areas:
                    st.info(f"Found {len(image_areas)} image upload areas in template")
                    with st.expander("Image Upload Areas", expanded=False):
                        image_df = pd.DataFrame(image_areas)
                        st.dataframe(image_df, use_container_width=True)
                
                # Show extracted images from data file
                if extracted_images:
                    total_images = sum(len(sheet_images) for sheet_images in extracted_images.values())
                    st.success(f"🖼️ Extracted {total_images} images from data file")
                    
                    with st.expander("Extracted Images from Data File", expanded=True):
                        for sheet_name, sheet_images in extracted_images.items():
                            if sheet_images:
                                st.write(f"**Sheet: {sheet_name}**")
                                cols = st.columns(min(3, len(sheet_images)))
                                
                                for idx, (position, img_data) in enumerate(sheet_images.items()):
                                    with cols[idx % 3]:
                                        st.write(f"Position: {position}")
                                        # Display image thumbnail
                                        img_bytes = base64.b64decode(img_data['data'])
                                        st.image(img_bytes, width=150)
                                        st.write(f"Size: {img_data['size']}")
                else:
                    if data_file.name.endswith(('.xlsx', '.xls')):
                        st.info("No images found in the data file")
                    else:
                        st.info("CSV files don't contain images. Use Excel files to include images.")
                
                # Data mapping
                st.subheader("🔗 Field Mapping")
                
                with st.spinner("Mapping template fields to data columns..."):
                    mapping_results = st.session_state.enhanced_mapper.map_data_with_section_context(
                        template_fields, data_df
                    )
                
                if mapping_results:
                    # Show mapping results
                    mapping_df = pd.DataFrame([
                        {
                            'Template Field': mapping['template_field'],
                            'Data Column': mapping['data_column'] if mapping['data_column'] else 'No Match',
                            'Similarity': f"{mapping['similarity']:.2f}" if mapping['similarity'] > 0 else "0.00",
                            'Section': mapping.get('section_context', 'Unknown'),
                            'Status': '✅ Mapped' if mapping['is_mappable'] else '❌ No Match'
                        }
                        for mapping in mapping_results.values()
                    ])
                    
                    st.dataframe(mapping_df, use_container_width=True)
                    
                    # Statistics
                    mapped_count = sum(1 for m in mapping_results.values() if m['is_mappable'])
                    total_count = len(mapping_results)
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Fields", total_count)
                    with col2:
                        st.metric("Mapped Fields", mapped_count)
                    with col3:
                        st.metric("Mapping Rate", f"{(mapped_count/total_count*100):.1f}%")
                    with col4:
                        images_count = sum(len(sheet_images) for sheet_images in extracted_images.values()) if extracted_images else 0
                        st.metric("Images Found", images_count)
                    
                    # Fill template
                    st.subheader("📝 Fill Template")
                    
                    if st.button("Generate Filled Template", type="primary", use_container_width=True):
                        with st.spinner("Filling template with data and extracted images..."):
                            try:
                                # Convert extracted images to the format expected by fill_template_with_data_and_images
                                processed_images = {}
                                if extracted_images:
                                    for sheet_name, sheet_images in extracted_images.items():
                                        for position, img_data in sheet_images.items():
                                            # Create a unique key for each image
                                            image_key = f"{sheet_name}_{position}"
                                            processed_images[image_key] = img_data
                                
                                workbook, filled_count, images_added, temp_image_paths = st.session_state.enhanced_mapper.fill_template_with_data_and_images(
                                    template_path, mapping_results, data_df, processed_images
                                )
                                
                                if workbook:
                                    # Save filled template
                                    output_buffer = io.BytesIO()
                                    workbook.save(output_buffer)
                                    for path in temp_image_paths:
                                        try:
                                            os.unlink(path)
                                        except Exception as e:
                                            st.warning(f"Failed to delete temp file {path}: {e}")
                                    output_buffer.seek(0)
                                    
                                    # Success message
                                    st.success(f"Template filled successfully!")
                                    st.info(f"📊 Filled {filled_count} data fields")
                                    if images_added > 0:
                                        st.info(f"🖼️ Added {images_added} images from data file")
                                    elif processed_images:
                                        st.warning("Images were found but could not be placed in template areas")
                                    
                                    # Download button
                                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                    filename = f"filled_template_{timestamp}.xlsx"
                                    
                                    st.download_button(
                                        label="📥 Download Filled Template",
                                        data=output_buffer.getvalue(),
                                        file_name=filename,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True
                                    )
                                    
                                    workbook.close()
                                else:
                                    st.error("Failed to fill template")
                                    
                            except Exception as e:
                                st.error(f"Error filling template: {e}")
                                st.exception(e)
                
                else:
                    st.warning("No mapping results generated")
            
            else:
                st.warning("No mappable fields found in template")
            
            # Clean up temporary file
            try:
                os.unlink(template_path)
            except:
                pass
                
        except Exception as e:
            st.error(f"Error processing files: {e}")
            st.exception(e)
    
    else:
        st.info("👆 Please upload both an Excel template and a data file to begin")
        
        # Show demo information
        st.markdown("### 🎯 Features")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **Template Processing:**
            - 📋 Smart field detection
            - 🎯 Section-aware mapping
            - 🔄 Merged cell handling
            - 📏 Packaging-specific patterns
            """)
            
        with col2:
            st.markdown("""
            **Image Processing:**
            - 🖼️ Auto image extraction from Excel data files
            - 📍 Smart image placement in templates
            - 🎨 Format conversion and optimization
            - 📦 Packaging image area detection
            """)
        
        st.markdown("""
        ### 📚 Supported Sections
        - **Primary Packaging**: Internal packaging dimensions and specifications
        - **Secondary Packaging**: Outer packaging details
        - **Part Information**: Component specifications and measurements
        
        ### 🖼️ Image Processing
        - Images are automatically extracted from Excel data files
        - Supports multiple image formats (PNG, JPG, GIF, BMP)
        - Images are intelligently placed in designated template areas
        - No manual image upload required - everything is automated!
        """)
        
# Main application logic
def main():
    if not st.session_state.authenticated:
        show_login()
    else:
        show_main_app()

if __name__ == "__main__":
    main()
