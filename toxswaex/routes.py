from flask import Blueprint, render_template, request, jsonify, send_file
import os
import tempfile
from .extractor import TOXSWAExtractor

toxswaex_bp = Blueprint('toxswaex', __name__,
                        template_folder='templates', 
                        static_folder='static')

# Global extractor instance
extractor = TOXSWAExtractor()

@toxswaex_bp.route('/')
def index():
    return render_template('toxswaex/index.html')

@toxswaex_bp.route('/scan_directory', methods=['POST'])
def scan_directory():
    try:
        data = request.get_json()
        directory = data.get('directory', '').strip()
        
        if not directory:
            return jsonify({'error': 'Please provide a directory path'})
        
        if not os.path.exists(directory):
            return jsonify({'error': f'Directory does not exist: {directory}'})
        
        # Look for projects (folders containing toxswa subfolder)
        projects = []
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            if os.path.isdir(item_path):
                toxswa_path = os.path.join(item_path, "toxswa")
                if os.path.exists(toxswa_path):
                    projects.append(item)
        
        projects.sort()
        
        return jsonify({
            'projects': projects,
            'main_dir': directory
        })
        
    except Exception as e:
        return jsonify({'error': f'Error scanning directory: {str(e)}'})

@toxswaex_bp.route('/extract_data', methods=['POST'])
def extract_data():
    try:
        data = request.get_json()
        main_dir = data.get('main_dir', '')
        selected_projects = data.get('selected_projects', [])
        selected_files = data.get('selected_files', None)
        rac_value = data.get('rac_value', None)
        areic_comparison = data.get('areic_comparison', False)
        
        if not main_dir:
            return jsonify({'error': 'No main directory specified'})
        
        if not selected_projects:
            return jsonify({'error': 'No projects selected'})
        
        # Convert RAC value to float if provided
        if rac_value:
            try:
                rac_value = float(rac_value)
            except ValueError:
                return jsonify({'error': 'Invalid RAC value'})
        
        # Extract data
        all_data, errors = extractor.extract_data(
            main_dir, selected_projects, selected_files, rac_value, areic_comparison
        )
        
        # Get table data for display
        compound_type = data.get('compound_type', 'Parent')
        sort_by = data.get('sort_by', 'Filename')
        table_data, headers = extractor.get_table_data(compound_type, sort_by)
        
        return jsonify({
            'data': table_data,
            'header': headers,
            'rac_value': rac_value,
            'row_count': len(table_data),
            'errors': errors
        })
        
    except Exception as e:
        return jsonify({'error': f'Error extracting data: {str(e)}'})

@toxswaex_bp.route('/export_excel', methods=['POST'])
def export_excel():
    try:
        if not extractor.all_data:
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
            download_name='toxswa_extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return jsonify({'error': f'Error exporting Excel: {str(e)}'})

@toxswaex_bp.route('/get_table_data')
def get_table_data():
    try:
        compound_type = request.args.get('compound_type', 'Parent')
        sort_by = request.args.get('sort_by', 'Filename')
        
        table_data, headers = extractor.get_table_data(compound_type, sort_by)
        
        return jsonify({
            'data': table_data,
            'header': headers
        })
        
    except Exception as e:
        return jsonify({'error': f'Error getting table data: {str(e)}'}) 