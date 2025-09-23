#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import os
import re
from datetime import datetime
from werkzeug.utils import secure_filename
import json

app = Flask(__name__)
CORS(app)

# Configuration
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'

# Create folders if they don't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_dlc(dlc_string):
    """Parse DLC value from various formats"""
    if pd.isna(dlc_string):
        return 8  # Default DLC
    
    dlc_str = str(dlc_string).strip()
    
    # Try to extract number from strings like "DLC = 8"
    match = re.search(r'\d+', dlc_str)
    if match:
        dlc = int(match.group())
        # Validate DLC range (0-8 for standard CAN, 0-64 for CAN FD)
        if dlc > 64:
            return 8
        return dlc
    
    return 8  # Default DLC if parsing fails

def detect_can_id_type(can_id_int):
    """Detect if CAN ID is standard (11-bit) or extended (29-bit)"""
    if can_id_int > 0x7FF:  # Greater than 11-bit max value (2047)
        return 'extended', True
    return 'standard', False

def parse_can_id(can_id_string, force_extended=False):
    """Parse CAN ID from various formats with extended ID support"""
    if pd.isna(can_id_string):
        return None, False
    
    can_id_str = str(can_id_string).strip()
    
    # Check if explicitly marked as extended
    is_extended = force_extended or 'x' in can_id_str.lower() or 'ext' in can_id_str.lower()
    
    # Remove common prefixes and extended markers
    can_id_str = re.sub(r'0x|0X|x|X|ext|EXT', '', can_id_str)
    
    try:
        # Parse as hex
        if isinstance(can_id_string, (int, float)):
            can_id_int = int(can_id_string)
        else:
            can_id_int = int(can_id_str, 16)
        
        # Auto-detect if extended based on value
        id_type, detected_extended = detect_can_id_type(can_id_int)
        is_extended = is_extended or detected_extended
        
        # Validate range
        if is_extended:
            if can_id_int > 0x1FFFFFFF:  # 29-bit max
                return None, False
            # Format extended ID with 8 hex digits and 'x' suffix
            return f"0x{can_id_int:08X}x", True
        else:
            if can_id_int > 0x7FF:  # Standard 11-bit max
                # Automaticky převést na extended pokud je větší
                return f"0x{can_id_int:08X}x", True
            # Format standard ID with 3 hex digits
            return f"0x{can_id_int:03X}", False
            
    except (ValueError, TypeError):
        return None, False

def parse_data_bytes(data_string, dlc):
    """Parse data bytes from various formats"""
    if pd.isna(data_string):
        return ' '.join(['00'] * min(dlc, 8))
    
    data_str = str(data_string).strip()
    
    # Remove common separators and clean up
    data_str = re.sub(r'[,;:]', ' ', data_str)
    
    # Split by whitespace and filter empty strings
    bytes_list = [b for b in data_str.split() if b]
    
    # Ensure each byte is 2 hex digits
    formatted_bytes = []
    for byte in bytes_list[:dlc]:  # Only take as many bytes as DLC specifies
        byte = byte.upper().replace('0X', '')
        # Validate hex characters
        if not all(c in '0123456789ABCDEF' for c in byte):
            byte = '00'  # Replace invalid with 00
        if len(byte) == 1:
            byte = '0' + byte
        elif len(byte) > 2:
            byte = byte[-2:]  # Take last 2 characters
        formatted_bytes.append(byte)
    
    # Pad with zeros if necessary
    while len(formatted_bytes) < 8:  # Always pad to 8 bytes for CANoe
        formatted_bytes.append('00')
    
    return ' '.join(formatted_bytes)

def detect_columns(df):
    """Automatically detect relevant columns in the dataframe"""
    columns_map = {
        'can_id': None,
        'dlc': None,
        'data': None,
        'address': None,
        'timeout': None  # Přidáno pro detekci timeout sloupce
    }
    
    for col in df.columns:
        col_lower = str(col).lower()
        
        # Detect CAN ID column
        if any(keyword in col_lower for keyword in ['can', 'id', 'canid', 'can_id', 'identifier', 'pgn']):
            if columns_map['can_id'] is None:
                columns_map['can_id'] = col
        
        # Detect DLC column
        if any(keyword in col_lower for keyword in ['dlc', 'length', 'len']):
            if columns_map['dlc'] is None:
                columns_map['dlc'] = col
        
        # Detect data bytes column
        if any(keyword in col_lower for keyword in ['byte', 'data', 'payload', 'message']):
            if columns_map['data'] is None:
                columns_map['data'] = col
        
        # Detect timeout column
        if any(keyword in col_lower for keyword in ['timeout', 'time', 'delay', 'wait']):
            if columns_map['timeout'] is None:
                columns_map['timeout'] = col
        
        # Detect address column (optional)
        if any(keyword in col_lower for keyword in ['address', 'addr', 'name', 'description']):
            if columns_map['address'] is None:
                columns_map['address'] = col
    
    return columns_map

def parse_timeout(timeout_value, default_timeout):
    """Parse timeout value from Excel cell"""
    if pd.isna(timeout_value):
        return default_timeout
    
    try:
        timeout = int(float(str(timeout_value).strip()))
        # Validate reasonable timeout range (1ms to 60000ms)
        if timeout < 1:
            return default_timeout
        if timeout > 60000:
            return 60000
        return timeout
    except (ValueError, TypeError):
        return default_timeout

def create_vsq_xml_header(sequence_name="GeneratedSequence"):
    """Create the XML header for VSQ file"""
    xml_header = f'''<?xml version="1.0" encoding="utf-8"?>
<VisualSequence version="1">
  <Settings>
    <NumberOfRepetitions>1</NumberOfRepetitions>
    <StartOnMeasurementStart>False</StartOnMeasurementStart>
    <RunUntilMeasurementStop>False</RunUntilMeasurementStop>
    <DebugMode>False</DebugMode>
    <ShowCommentColumn>False</ShowCommentColumn>
    <LogToWrite>True</LogToWrite>
    <LogToFile>False</LogToFile>
    <LogFile>{sequence_name}.csv</LogFile>
    <CSVColumnSeparator>,</CSVColumnSeparator>
    <CSVDecimalSymbol>.</CSVDecimalSymbol>
    <CSVDecimalPlaces>6</CSVDecimalPlaces>
    <LogTimeStamp>False</LogTimeStamp>
    <SymbolNameDisplay>{sequence_name}</SymbolNameDisplay>
    <WaitForKeyKey />
    <CheckOutputFailedOnly>False</CheckOutputFailedOnly>
    <UseSignalLayer>False</UseSignalLayer>
    <ExecMode>Standard</ExecMode>
  </Settings>
</VisualSequence>'''
    return xml_header

def process_excel_to_vsq(file_path, output_name=None, default_timeout=3000, 
                         can_channel='CAN1', force_extended=False):
    """
    Main function to process Excel file and generate VSQ
    
    Args:
        file_path: Path to Excel file
        output_name: Custom output filename
        default_timeout: Default timeout in ms (user configurable)
        can_channel: CAN channel to use (CAN1, CAN2, etc.)
        force_extended: Force all IDs to be treated as extended
    """
    try:
        # Read Excel file
        df = pd.read_excel(file_path, sheet_name=0)
        
        # Remove completely empty rows
        df = df.dropna(how='all')
        
        # Detect columns
        columns = detect_columns(df)
        
        if not columns['can_id'] or not columns['data']:
            return None, {'success': False, 'error': 'Could not detect CAN ID or Data columns in the Excel file'}
        
        # Generate output filename
        if not output_name:
            output_name = os.path.splitext(os.path.basename(file_path))[0]
        
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{output_name}.vsq")
        
        # Create VSQ content
        vsq_lines = []
        vsq_lines.append(create_vsq_xml_header(output_name))
        
        # Process each row
        warnings = []
        processed_count = 0
        standard_ids = 0
        extended_ids = 0
        preview_lines = []  # Pro live preview
        
        for idx, row in df.iterrows():
            # Skip header rows or rows without CAN ID
            can_id, is_extended = parse_can_id(row[columns['can_id']], force_extended)
            if not can_id:
                continue
            
            # Count ID types
            if is_extended:
                extended_ids += 1
            else:
                standard_ids += 1
            
            # Get DLC
            dlc = 8  # Default
            if columns['dlc']:
                dlc = parse_dlc(row[columns['dlc']])
            
            # Get timeout (from column or default)
            timeout = default_timeout
            if columns['timeout']:
                timeout = parse_timeout(row[columns['timeout']], default_timeout)
            
            # Get data bytes
            data_bytes = parse_data_bytes(row[columns['data']], dlc)
            
            # Validate data bytes count
            actual_bytes = len([b for b in data_bytes.split() if b and b != '00'])
            if actual_bytes > dlc:
                warnings.append(f"Row {idx+1}: Data bytes ({actual_bytes}) exceed DLC ({dlc})")
            
            # Create VSQ line
            vsq_line = f"1,Send CAN Raw Frame,{can_channel}::{can_id},=,{data_bytes},{timeout},,False,False,False"
            vsq_lines.append(vsq_line)
            
            # Add to preview (first 10 lines)
            if processed_count < 10:
                preview_lines.append({
                    'line_num': processed_count + 1,
                    'can_id': can_id,
                    'is_extended': is_extended,
                    'dlc': dlc,
                    'data': data_bytes,
                    'timeout': timeout,
                    'raw': vsq_line
                })
            
            processed_count += 1
        
        # Write to file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(vsq_lines))
        
        result = {
            'success': True,
            'filename': f"{output_name}.vsq",
            'messages_processed': processed_count,
            'standard_ids': standard_ids,
            'extended_ids': extended_ids,
            'warnings': warnings,
            'detected_columns': {k: v for k, v in columns.items() if v is not None},
            'preview': preview_lines,
            'settings': {
                'default_timeout': default_timeout,
                'can_channel': can_channel,
                'force_extended': force_extended
            }
        }
        
        return output_path, result
    
    except Exception as e:
        return None, {'success': False, 'error': str(e)}

@app.route('/')
def index():
    """Main page"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing with advanced options"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file provided'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No file selected'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{timestamp}_{filename}"
        
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Get advanced parameters
        output_name = request.form.get('output_name', None)
        default_timeout = int(request.form.get('timeout', 3000))
        can_channel = request.form.get('can_channel', 'CAN1')
        force_extended = request.form.get('force_extended', 'false').lower() == 'true'
        
        # Process the file
        output_path, result = process_excel_to_vsq(
            file_path, 
            output_name, 
            default_timeout,
            can_channel,
            force_extended
        )
        
        if result['success']:
            return jsonify(result), 200
        else:
            return jsonify(result), 500
    
    return jsonify({'success': False, 'error': 'Invalid file type'}), 400

@app.route('/preview', methods=['POST'])
def preview_only():
    """Generate preview without saving file"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file provided'}), 400
    
    file = request.files['file']
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"preview_{timestamp}_{filename}"
        
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Get parameters
        default_timeout = int(request.form.get('timeout', 3000))
        can_channel = request.form.get('can_channel', 'CAN1')
        force_extended = request.form.get('force_extended', 'false').lower() == 'true'
        
        try:
            # Read and analyze file
            df = pd.read_excel(file_path, sheet_name=0)
            df = df.dropna(how='all')
            columns = detect_columns(df)
            
            preview_data = {
                'success': True,
                'total_rows': len(df),
                'detected_columns': {k: v for k, v in columns.items() if v is not None},
                'sample_data': []
            }
            
            # Generate sample preview
            for idx, row in df.head(5).iterrows():
                can_id, is_extended = parse_can_id(row[columns['can_id']], force_extended)
                if can_id:
                    dlc = parse_dlc(row[columns['dlc']]) if columns['dlc'] else 8
                    timeout = parse_timeout(row[columns['timeout']], default_timeout) if columns['timeout'] else default_timeout
                    data_bytes = parse_data_bytes(row[columns['data']], dlc) if columns['data'] else '00 00 00 00 00 00 00 00'
                    
                    preview_data['sample_data'].append({
                        'can_id': can_id,
                        'is_extended': is_extended,
                        'dlc': dlc,
                        'data': data_bytes,
                        'timeout': timeout
                    })
            
            # Clean up preview file
            os.remove(file_path)
            
            return jsonify(preview_data), 200
            
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500
    
    return jsonify({'success': False, 'error': 'Invalid file type'}), 400

@app.route('/download/<filename>')
def download_file(filename):
    """Download generated VSQ file"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy', 
        'version': '2.0.0',
        'features': ['extended_can_id', 'configurable_timeout', 'live_preview'],
        'timestamp': datetime.now().isoformat()
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)