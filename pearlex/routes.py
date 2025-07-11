from flask import Blueprint, render_template, request, jsonify, send_file
import os
import re
import tempfile
from .extractor import PEARLexExtractor

pearlex = Blueprint('pearlex', __name__, 
                   template_folder='templates',
                   static_folder='static',
                   url_prefix='/pearlex')

@pearlex.route('/')
def pearlex_index():
    return render_template('pearlex/index.html')

@pearlex.route('/scan_directory', methods=['POST'])
def scan_directory():
    """Scan directory for .sum files"""
    try:
        data = request.get_json()
        directory_path = data.get('directory', '').strip()
        
        if not directory_path or not os.path.exists(directory_path):
            return jsonify({'error': 'Invalid directory path'}), 400
        
        extractor = PEARLexExtractor()
        projects = extractor.scan_directory(directory_path)
        
        return jsonify({
            'success': True,
            'projects': projects
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@pearlex.route('/extract_data', methods=['POST'])
def extract_data():
    """Extract data from selected .sum files"""
    try:
        data = request.get_json()
        directory_path = data.get('directory', '').strip()
        selected_files = data.get('selected_files', [])
        limit_value = data.get('limit_value')
        
        if not directory_path or not selected_files:
            return jsonify({'error': 'Missing directory or selected files'}), 400
        
        extractor = PEARLexExtractor()
        extracted_data = extractor.extract_data(directory_path, selected_files, limit_value)
        
        return jsonify({
            'success': True,
            'data': extracted_data
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@pearlex.route('/export_excel', methods=['POST'])
def export_excel():
    """Export extracted data to Excel file"""
    try:
        data = request.get_json()
        extracted_data = data.get('data', [])
        limit_value = data.get('limit_value')
        
        if not extracted_data:
            return jsonify({'error': 'No data to export'}), 400
        
        extractor = PEARLexExtractor()
        file_path = extractor.export_to_excel(extracted_data, limit_value)
        
        return jsonify({
            'success': True,
            'file_path': file_path
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@pearlex.route('/download_excel/<path:file_path>')
def download_excel(file_path):
    """Download Excel file"""
    try:
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name='pearlex_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500 