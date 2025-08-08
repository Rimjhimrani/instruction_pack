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

# Configure Streamlit page
st.set_page_config(
    page_title="AI Packaging Template Mapper",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
                'secondary': 'I42',  # Secondary packaging at I42 (next column set)
                'label': 'Q42'  # Label at Q42 (next column set)
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
        
        # Section-based mapping rules from reference code
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
                    'type': 'Secondary Packaging Type',
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
                    'part information', 'part info', 'part', 'component', 'item'
                ],
                'field_mappings': {
                    'part no': 'Part No',
                    'part number': 'Part No',
                    'description': 'Part Description',
                    'unit weight': 'Part Unit Weight',
                    'L': 'Part L',
                    'W': 'Part W',
                    'H': 'Part H'
                }
            },
            'vendor_information': {
                'section_keywords': [
                    'vendor information', 'vendor info', 'vendor', 'supplier'
                ],
                'field_mappings': {
                    'vendor name': 'Vendor Name',
                    'name': 'Vendor Name',
                    'vendor code': 'Vendor Code',
                    'vendor location': 'Vendor Location'
                }
            }
        }

    def map_template_with_data(self, template_path, data_path):
        """Map template with data from Excel file"""
        try:
            # Read data from Excel
            data_df = pd.read_excel(data_path)
            
            # Load template
            workbook = openpyxl.load_workbook(template_path)
            worksheet = workbook.active
            
            mapped_fields = {}
            
            # Simple mapping - look for matching field names
            for index, row in data_df.iterrows():
                for col_name, value in row.items():
                    if pd.notna(value) and col_name:
                        # Clean column name for matching
                        clean_col = str(col_name).lower().strip()
                        mapped_fields[clean_col] = str(value)
            
            # Apply mappings to template
            for row_num in range(1, worksheet.max_row + 1):
                for col_num in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=row_num, column=col_num)
                    if cell.value and isinstance(cell.value, str):
                        cell_text = cell.value.lower().strip()
                        
                        # Check for direct matches
                        for field_key, field_value in mapped_fields.items():
                            if field_key in cell_text or cell_text in field_key:
                                # Look for adjacent empty cell to fill
                                adjacent_cells = [
                                    worksheet.cell(row=row_num, column=col_num + 1),
                                    worksheet.cell(row=row_num + 1, column=col_num),
                                    worksheet.cell(row=row_num, column=col_num + 2)
                                ]
                                
                                for adj_cell in adjacent_cells:
                                    if not adj_cell.value:
                                        adj_cell.value = field_value
                                        break
            
            return workbook, mapped_fields
            
        except Exception as e:
            st.error(f"Error mapping template: {e}")
            return None, {}

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
        "Seal carton box and put {Inner Qty/Pack} such carton boxes into another carton box [L-{Inner L} mm, W-{Inner W} mm, H-{Inner H} mm]",
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
        "Put {Inner Qty/Pack} such carton boxes into another carton box [L-{Inner L} mm, W-{Inner W} mm, H-{Inner H} mm]",
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
                    st.session_state.current_step = 2
                    st.rerun()
        
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
            
            if st.button("Continue to Data Upload"):
                st.session_state.current_step = 3
                st.rerun()
    
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
            
            if st.button("Continue to Auto-Fill"):
                st.session_state.current_step = 4
                st.rerun()
    
    # Step 4: Auto-Fill Template
    elif st.session_state.current_step == 4:
        st.header("🔄 Step 4: Auto-Fill Template")
        
        if st.button("Start Auto-Fill Process"):
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
                        
                        st.success(f"✅ Template auto-filled with {len(mapped_fields)} data fields!")
                        
                        # Show mapped fields
                        if mapped_fields:
                            with st.expander("View Mapped Fields"):
                                for field, value in mapped_fields.items():
                                    st.write(f"**{field}**: {value}")
                        
                        if st.button("Continue to Image Options"):
                            st.session_state.current_step = 5
                            st.rerun()
                    else:
                        st.error("Failed to process template mapping")
                        
                except Exception as e:
                    st.error(f"Error during auto-fill: {e}")
                    st.write("Traceback:", traceback.format_exc())
    
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
            if st.button("Continue to Final Generation"):
                st.session_state.current_step = 6
                st.rerun()
    
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
