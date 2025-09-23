#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import os
import re
from datetime import datetime
from werkzeug.utils import secure_filename
import xml.etree.ElementTree as ET
from xml.dom import minidom

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
        return int(match.group())
    
    return 8  # Default DLC if parsing fails

def parse_can_id(can_id_string):
    """Parse CAN ID from various formats"""
    if pd.isna(can_id_string):
        return None
    
    can_id_str = str(can_id_string).strip()
    
    # Remove common prefixes
    can_id_str = can_id_str.replace('0x', '').replace('0X', '')
    
    # Try to parse as hex
    try:
        # If it's already an integer (from Excel), convert to hex string
        if isinstance(can_id_string, (int, float)):
            return f"0x{int(can_id_string):03X}"
        else:
            # Parse hex string
            can_id_int = int(can_id_str, 16)
            return f"0x{can_id_int:03X}"
    except (ValueError, TypeError):
        return None

def parse_data_bytes(data_string, dlc):
    """Parse data bytes from various formats"""
    if pd.isna(data_string):
        return ' '.join(['00'] * dlc)
    
    data_str = str(data_string).strip()
    
    # Remove common separators and clean up
    data_str = re.sub(r'[,;:]', ' ', data_str)
    
    # Split by whitespace and filter empty strings
    bytes_list = [b for b in data_str.split() if b]
    
    # Ensure each byte is 2 hex digits
    formatted_bytes = []
    for byte in bytes_list[:dlc]:  # Only take as many bytes as DLC specifies
        byte = byte.upper().replace('0X', '')
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
        'address': None
    }
    
    # Convert column names to lowercase for easier matching
    columns_lower = {col: col for col in df.columns}
    for col in df.columns:
        col_lower = str(col).lower()
        
        # Detect CAN ID column
        if any(keyword in col_lower for keyword in ['can', 'id', 'canid', 'can_id', 'identifier']):
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
        
        # Detect address column (optional)
        if any(keyword in col_lower for keyword in ['address', 'addr', 'name']):
            if columns_map['address'] is None:
                columns_map['address'] = col
    
    return columns_map

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

def process_excel_to_vsq(file_path, output_name=None):
    """Main function to process Excel file and generate VSQ"""
    try:
        # Read Excel file
        df = pd.read_excel(file_path, sheet_name=0)
        
        # Remove completely empty rows
        df = df.dropna(how='all')
        
        # Detect columns
        columns = detect_columns(df)
        
        if not columns['can_id'] or not columns['data']:
            return None, "Could not detect CAN ID or Data columns in the Excel file"
        
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
        
        for idx, row in df.iterrows():
            # Skip header rows or rows without CAN ID
            can_id = parse_can_id(row[columns['can_id']])
            if not can_id:
                continue
            
            # Get DLC
            dlc = 8  # Default
            if columns['dlc']:
                dlc = parse_dlc(row[columns['dlc']])
            
            # Get data bytes
            data_bytes = parse_data_bytes(row[columns['data']], dlc)
            
            # Validate data bytes count
            actual_bytes = len([b for b in data_bytes.split() if b and b != '00'])
            if actual_bytes > dlc:
                warnings.append(f"Row {idx+1}: Data bytes ({actual_bytes}) exceed DLC ({dlc})")
            
            # Create VSQ line
            # Format: sequence_number, action, channel::address, operator, data, timeout, , bool, bool, bool
            vsq_line = f"1,Send CAN Raw Frame,CAN1::{can_id},=,{data_bytes},3000,,False,False,False"
            vsq_lines.append(vsq_line)
            processed_count += 1
        
        # Write to file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(vsq_lines))
        
        result = {
            'success': True,
            'filename': f"{output_name}.vsq",
            'messages_processed': processed_count,
            'warnings': warnings,
            'detected_columns': {k: v for k, v in columns.items() if v is not None}
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
    """Handle file upload and processing"""
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
        
        # Get custom output name if provided
        output_name = request.form.get('output_name', None)
        
        # Process the file
        output_path, result = process_excel_to_vsq(file_path, output_name)
        
        if result['success']:
            return jsonify(result), 200
        else:
            return jsonify(result), 500
    
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
    return jsonify({'status': 'healthy', 'timestamp': datetime.now().isoformat()})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)