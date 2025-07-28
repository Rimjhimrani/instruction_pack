import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import get_column_letter
import tempfile
import os
import hashlib
import io
import base64
import re
import zipfile
from datetime import datetime
from difflib import SequenceMatcher
from PIL import Image
import numpy as np

# Advanced NLP imports (optional)
try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    from nltk.corpus import stopwords
    from nltk.tokenize import word_tokenize
    import nltk
    nltk.download('punkt', quiet=True)
    nltk.download('stopwords', quiet=True)
    ADVANCED_NLP = True
    NLTK_READY = True
except ImportError:
    ADVANCED_NLP = False
    NLTK_READY = False

# Set page config
st.set_page_config(
    page_title="AI Template Mapper",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

class ImageExtractor:
    """Handles image extraction from Excel files"""
    
    def __init__(self):
        self.supported_formats = ['.png', '.jpg', '.jpeg', '.gif', '.bmp']
    
    def extract_images_from_excel(self, excel_file_path):
        """Extract all images from Excel file with enhanced categorization"""
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
                            
                            # Get image position and determine type
                            anchor = img.anchor
                            if hasattr(anchor, '_from'):
                                col = anchor._from.col
                                row = anchor._from.row
                                position = f"{get_column_letter(col + 1)}{row + 1}"
                                
                                # Determine image type based on position and nearby text
                                image_type = self.categorize_image_by_context(worksheet, row + 1, col + 1)
                            else:
                                position = f"Image_{idx + 1}"
                                image_type = "uncategorized"
                            
                            # Convert to base64 for storage
                            buffered = io.BytesIO()
                            pil_image.save(buffered, format="PNG")
                            img_str = base64.b64encode(buffered.getvalue()).decode()
                            
                            # Use image type as key, or fallback to position
                            image_key = image_type if image_type != "uncategorized" else position
                            
                            sheet_images[image_key] = {
                                'data': img_str,
                                'format': 'PNG',
                                'size': pil_image.size,
                                'position': position,
                                'type': image_type,
                                'original_index': idx
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
    
    def categorize_image_by_context(self, worksheet, img_row, img_col, search_range=10):
        """Categorize image based on nearby text context"""
        try:
            # Define patterns for different packaging types
            packaging_patterns = {
                'primary_packaging': [
                    'primary packaging', 'primary', 'internal packaging', 'inner packaging',
                    'primary / internal', 'primary/internal'
                ],
                'secondary_packaging': [
                    'secondary packaging', 'secondary', 'outer packaging', 'external packaging',
                    'outer / external', 'outer/external', 'box', 'carton'
                ],
                'current_packaging': [
                    'current packaging', 'existing packaging', 'current', 'as-is packaging',
                    'present packaging'
                ]
            }
            
            # Search in nearby cells for context clues
            for search_row in range(max(1, img_row - search_range), min(worksheet.max_row + 1, img_row + search_range)):
                for search_col in range(max(1, img_col - search_range), min(worksheet.max_column + 1, img_col + search_range)):
                    try:
                        cell = worksheet.cell(row=search_row, column=search_col)
                        if cell.value:
                            cell_text = str(cell.value).lower().strip()
                            
                            # Check against patterns
                            for package_type, patterns in packaging_patterns.items():
                                for pattern in patterns:
                                    if pattern in cell_text:
                                        return package_type
                    except:
                        continue
            
            # Fallback: try to determine by column headers
            header_row = 1
            try:
                header_cell = worksheet.cell(row=header_row, column=img_col)
                if header_cell.value:
                    header_text = str(header_cell.value).lower().strip()
                    for package_type, patterns in packaging_patterns.items():
                        for pattern in patterns:
                            if pattern in header_text:
                                return package_type
            except:
                pass
            
            return "uncategorized"
            
        except Exception as e:
            print(f"Error categorizing image: {e}")
            return "uncategorized"
    
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
    
    def extract_data_and_images_from_excel(self, excel_file):
        """Extract both data and images from the uploaded Excel file"""
        try:
            # Save uploaded file temporarily to extract images
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(excel_file.getvalue())
                tmp_file_path = tmp_file.name
            
            # Extract images from the Excel file
            extracted_images = self.image_extractor.extract_images_from_excel(tmp_file_path)
            
            # Read data using pandas
            data_df = pd.read_excel(excel_file, engine='openpyxl')
            
            # Clean up temporary file
            try:
                os.unlink(tmp_file_path)
            except:
                pass
            
            # Flatten images from all sheets into a single dictionary
            all_images = {}
            for sheet_name, sheet_images in extracted_images.items():
                for img_key, img_data in sheet_images.items():
                    # Prefix with sheet name if multiple sheets
                    if len(extracted_images) > 1:
                        final_key = f"{sheet_name}_{img_key}"
                    else:
                        final_key = img_key
                    all_images[final_key] = img_data
            
            return data_df, all_images
            
        except Exception as e:
            st.error(f"Error extracting data and images from Excel: {e}")
            return None, {}
    
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
    
    def fill_template_with_mapped_data_and_images(self, template_file, mapping_results, data_df, uploaded_images):
        """Fill template with mapped data and images"""
        try:
            workbook = openpyxl.load_workbook(template_file)
            worksheet = workbook.active
            temp_image_paths = []
            
            # Get image areas for precise placement
            _, image_areas = self.find_template_fields_with_context_and_images(template_file)
            
            # Fill data fields
            filled_fields = 0
            for coord, mapping in mapping_results.items():
                if mapping['data_column'] and mapping['is_mappable']:
                    try:
                        field_info = mapping['field_info']
                        data_column = mapping['data_column']
                        
                        # Find the appropriate data cell
                        data_cell_coord = self.find_data_cell_for_label(worksheet, field_info)
                        
                        if data_cell_coord and data_column in data_df.columns:
                            # Get the first non-null value from the data column
                            data_values = data_df[data_column].dropna()
                            if not data_values.empty:
                                value = str(data_values.iloc[0])
                                worksheet[data_cell_coord] = value
                                filled_fields += 1
                    except Exception as e:
                        st.warning(f"Could not fill field {coord}: {e}")
                        continue
            
            # Add images with precise positioning
            added_images = 0
            if uploaded_images and image_areas:
                added_images, temp_paths = self.add_images_to_template_precise(
                    worksheet, uploaded_images, image_areas
                )
                temp_image_paths.extend(temp_paths)
            
            # Save to memory
            output = io.BytesIO()
            workbook.save(output)
            output.seek(0)
            workbook.close()
            
            # Clean up temporary image files
            for temp_path in temp_image_paths:
                try:
                    if os.path.exists(temp_path):
                        os.unlink(temp_path)
                except:
                    pass
            
            return output.getvalue(), filled_fields, added_images
            
        except Exception as e:
            st.error(f"Error filling template: {e}")
            return None, 0, 0

def main():
    """Main Streamlit application"""
    st.title("🤖 AI Template Mapper with Images")
    st.markdown("---")
    
    # Initialize the mapper
    mapper = EnhancedTemplateMapperWithImages()
    
    # Sidebar configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        similarity_threshold = st.slider(
            "Similarity Threshold", 
            min_value=0.1, 
            max_value=1.0, 
            value=0.3, 
            step=0.1,
            help="Lower values = more lenient matching"
        )
        mapper.similarity_threshold = similarity_threshold
        
        st.markdown("---")
        st.markdown("### 📋 Instructions")
        st.markdown("""
        1. **Upload Template**: Excel template with fields to fill
        2. **Upload Data**: Excel file with source data
        3. **Upload Images**: Optional packaging images
        4. **Review Mappings**: Check field mappings
        5. **Download Result**: Get filled template
        """)
    
    # File upload section
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Template File")
        template_file = st.file_uploader(
            "Upload Excel template", 
            type=['xlsx', 'xls'], 
            key="template",
            help="Excel file with fields to be filled"
        )
    
    with col2:
        st.subheader("📊 Data File")
        data_file = st.file_uploader(
            "Upload Excel data", 
            type=['xlsx', 'xls'], 
            key="data",
            help="Excel file containing source data"
        )
    
    # Image upload section
    st.subheader("🖼️ Packaging Images (Optional)")
    uploaded_images_files = st.file_uploader(
        "Upload packaging images",
        type=['png', 'jpg', 'jpeg', 'gif', 'bmp'],
        accept_multiple_files=True,
        key="images",
        help="Upload images for primary, secondary, or current packaging"
    )
    
    # Process uploaded images
    uploaded_images = {}
    if uploaded_images_files:
        st.write(f"📸 Processing {len(uploaded_images_files)} uploaded images...")
        
        for idx, img_file in enumerate(uploaded_images_files):
            try:
                # Open and process image
                pil_image = Image.open(img_file)
                
                # Convert to base64
                buffered = io.BytesIO()
                pil_image.save(buffered, format="PNG")
                img_str = base64.b64encode(buffered.getvalue()).decode()
                
                # Use filename (without extension) as key
                img_name = os.path.splitext(img_file.name)[0]
                uploaded_images[img_name] = {
                    'data': img_str,
                    'format': 'PNG',
                    'size': pil_image.size,
                    'filename': img_file.name
                }
                
                # Display thumbnail
                with st.expander(f"🖼️ {img_file.name}"):
                    st.image(pil_image, width=200)
                    st.write(f"Size: {pil_image.size[0]} x {pil_image.size[1]} pixels")
                    
            except Exception as e:
                st.error(f"Error processing image {img_file.name}: {e}")
    
    # Main processing
    if template_file and data_file:
        try:
            # Extract data and images from files
            with st.spinner("🔍 Analyzing files..."):
                # Process data file
                data_df, extracted_images = mapper.extract_data_and_images_from_excel(data_file)
                
                if data_df is None:
                    st.error("❌ Could not read data file")
                    return
                
                # Find template fields and image areas
                template_fields, image_areas = mapper.find_template_fields_with_context_and_images(template_file)
                
                if not template_fields:
                    st.error("❌ No mappable fields found in template")
                    return
            
            # Combine uploaded and extracted images
            all_images = {**extracted_images, **uploaded_images}
            
            # Display analysis results
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Template Fields", len(template_fields))
            with col2:
                st.metric("Data Columns", len(data_df.columns))
            with col3:
                st.metric("Total Images", len(all_images))
            
            # Perform mapping
            with st.spinner("🎯 Mapping fields..."):
                mapping_results = mapper.map_data_with_section_context(template_fields, data_df)
            
            # Display mapping results
            st.subheader("🎯 Field Mapping Results")
            
            successful_mappings = sum(1 for m in mapping_results.values() if m['is_mappable'])
            st.write(f"**Successfully mapped:** {successful_mappings}/{len(mapping_results)} fields")
            
            # Show detailed mappings
            with st.expander("📋 View Detailed Mappings", expanded=True):
                for coord, mapping in mapping_results.items():
                    col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
                    
                    with col1:
                        st.write(f"**{mapping['template_field']}**")
                    
                    with col2:
                        if mapping['is_mappable']:
                            st.write(f"✅ {mapping['data_column']}")
                        else:
                            st.write("❌ No match found")
                    
                    with col3:
                        if mapping['is_mappable']:
                            st.write(f"{mapping['similarity']:.2f}")
                        else:
                            st.write("-")
                    
                    with col4:
                        section = mapping.get('section_context', 'General')
                        st.write(f"*{section}*")
            
            # Image areas information
            if image_areas:
                st.subheader("🖼️ Image Upload Areas")
                with st.expander("View Image Areas"):
                    for area in image_areas:
                        st.write(f"**{area['type'].replace('_', ' ').title()}**: {area['position']} ({area['header_text']})")
            
            # Generate filled template
            if st.button("🚀 Generate Filled Template", type="primary"):
                with st.spinner("✨ Filling template..."):
                    filled_template, filled_fields, added_images = mapper.fill_template_with_mapped_data_and_images(
                        template_file, mapping_results, data_df, all_images
                    )
                
                if filled_template:
                    st.success(f"✅ Template filled successfully!")
                    st.write(f"- **Fields filled:** {filled_fields}")
                    st.write(f"- **Images added:** {added_images}")
                    
                    # Generate filename
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    download_filename = f"filled_template_{timestamp}.xlsx"
                    
                    # Download button
                    st.download_button(
                        label="📥 Download Filled Template",
                        data=filled_template,
                        file_name=download_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ Failed to generate filled template")
                    
        except Exception as e:
            st.error(f"❌ Error processing files: {e}")
            st.exception(e)
    
    else:
        st.info("👆 Please upload both template and data files to begin")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666; font-size: 0.9em;'>
        🤖 AI Template Mapper | Enhanced with Image Support
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main() 
