from flask import Blueprint, render_template, request, jsonify, send_file
import os
import tempfile
from .extractor import PELMOExtractor

pelmoex_bp = Blueprint('pelmoex', __name__, 
                      template_folder='templates',
                      static_folder='static')

# Global extractor instance
extractor = PELMOExtractor()

@pelmoex_bp.route('/')
def index():
    return render_template('pelmoex/index.html')

@pelmoex_bp.route('/scan_directory', methods=['POST'])
def scan_directory():
    try:
        data = request.get_json()
        directory = data.get('directory', '').strip()
        
        if not directory:
            return jsonify({'error': 'Please provide a directory path'})
        
        if not os.path.exists(directory):
            return jsonify({'error': f'Directory does not exist: {directory}'})
        
        # Look for FOCUS folder
        focus_path = os.path.join(directory, "FOCUS")
        if not os.path.exists(focus_path):
            return jsonify({'error': f'FOCUS folder not found in: {directory}'})
        
        # Look for projects (folders ending with .run)
        projects = []
        for item in os.listdir(focus_path):
            item_path = os.path.join(focus_path, item)
            if os.path.isdir(item_path) and item.endswith(".run"):
                projects.append(item)
        
        projects.sort()
        
        return jsonify({
            'projects': projects,
            'focus_path': focus_path
        })
        
    except Exception as e:
        return jsonify({'error': f'Error scanning directory: {str(e)}'})

@pelmoex_bp.route('/extract_data', methods=['POST'])
def extract_data():
    try:
        data = request.get_json()
        focus_path = data.get('focus_path', '')
        selected_projects = data.get('selected_projects', [])
        limit_value = data.get('limit_value', None)
        
        if not focus_path:
            return jsonify({'error': 'No FOCUS path specified'})
        
        if not selected_projects:
            return jsonify({'error': 'No projects selected'})
        
        # Convert limit value to float if provided
        if limit_value:
            try:
                limit_value = float(limit_value)
            except ValueError:
                return jsonify({'error': 'Invalid limit value'})
        
        # Extract data
        all_rows, header, errors = extractor.extract_data(focus_path, selected_projects, limit_value)
        
        return jsonify({
            'data': all_rows,
            'header': header,
            'limit_value': limit_value,
            'row_count': len(all_rows),
            'errors': errors
        })
        
    except Exception as e:
        return jsonify({'error': f'Error extracting data: {str(e)}'})

@pelmoex_bp.route('/export_excel', methods=['POST'])
def export_excel():
    try:
        if not extractor.all_rows:
            return jsonify({'error': 'No data to export'})
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            filepath = tmp_file.name
        
        # Export to Excel
        extractor.export_to_excel(filepath)
        
        # Send file
        return send_file(
            filepath,
            as_attachment=True,
            download_name='pelmo_extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return jsonify({'error': f'Error exporting Excel: {str(e)}'})

@pelmoex_bp.route('/get_table_data')
def get_table_data():
    try:
        return jsonify({
            'data': extractor.all_rows,
            'header': ["Project", "Crop", "Scenario"] + sorted(set().union(*[set(row.keys()) for row in extractor.all_rows]) - {"Project", "Crop", "Scenario"})
        })
        
    except Exception as e:
        return jsonify({'error': f'Error getting table data: {str(e)}'}) 