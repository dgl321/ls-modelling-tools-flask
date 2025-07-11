from flask import Flask, render_template, Blueprint, request, jsonify, send_file
import os
import tempfile

app = Flask(__name__)

# Import the extractors
from pelmoex.extractor import PELMOExtractor
from toxswaex.extractor import TOXSWAExtractor
from pearlex.extractor import PearlGroundwaterExtractor

# Create blueprints with full functionality
pelmoex_bp = Blueprint('pelmoex', __name__, 
                      template_folder='pelmoex/templates',
                      static_folder='pelmoex/static')

# Global extractor instances
pelmo_extractor = PELMOExtractor()
pearl_extractor = PearlGroundwaterExtractor()

@pelmoex_bp.route('/')
def pelmoex_index():
    return render_template('pelmoex/index.html')

@pelmoex_bp.route('/scan_directory', methods=['POST'])
def pelmoex_scan_directory():
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
def pelmoex_extract_data():
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
        all_rows, header, errors = pelmo_extractor.extract_data(focus_path, selected_projects, limit_value)
        
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
def pelmoex_export_excel():
    try:
        if not pelmo_extractor.all_rows:
            return jsonify({'error': 'No data to export'})
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            filepath = tmp_file.name
        
        # Export to Excel
        pelmo_extractor.export_to_excel(filepath)
        
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
def pelmoex_get_table_data():
    try:
        return jsonify({
            'data': pelmo_extractor.all_rows,
            'header': ["Project", "Crop", "Scenario"] + sorted(set().union(*[set(row.keys()) for row in pelmo_extractor.all_rows]) - {"Project", "Crop", "Scenario"})
        })
        
    except Exception as e:
        return jsonify({'error': f'Error getting table data: {str(e)}'})

# TOXSWAex Blueprint
toxswaex_bp = Blueprint('toxswaex', __name__,
                        template_folder='toxswaex/templates', 
                        static_folder='toxswaex/static')

# Global TOXSWA extractor instance
toxswa_extractor = TOXSWAExtractor()

@toxswaex_bp.route('/')
def toxswaex_index():
    return render_template('toxswaex/index.html')

@toxswaex_bp.route('/scan_directory', methods=['POST'])
def toxswaex_scan_directory():
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
def toxswaex_extract_data():
    try:
        data = request.get_json()
        main_dir = data.get('main_dir', '')
        selected_projects = data.get('selected_projects', [])
        selected_files = data.get('selected_files', None)
        rac_value = data.get('rac_value', None)
        areic_comparison = data.get('areic_comparison', False)
        summary_mode = data.get('summary_mode', False)
        project_order = data.get('project_order', [])
        
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
        print(f"Extracting data from {main_dir} for projects: {selected_projects}")
        print(f"Summary mode: {summary_mode}")
        print(f"Project order: {project_order}")
        all_data, errors = toxswa_extractor.extract_data(
            main_dir, selected_projects, selected_files, rac_value, 
            areic_comparison, summary_mode, project_order
        )
        print(f"Extraction result: {len(all_data) if all_data else 0} projects, {sum(len(rows) for rows in all_data.values()) if all_data else 0} total rows")
        if errors:
            print(f"Extraction errors: {errors}")
        
        # Convert all_data to flat list for client-side processing
        all_rows = []
        headers = []
        
        if all_data:
            # Get all unique keys for headers
            all_keys = set()
            for project, rows in all_data.items():
                for row in rows:
                    all_keys.update(row.keys())
            
            # Create headers list
            headers = ["Project", "Filename", "Compound", "Scenario", "Waterbody", "Max PECsw", "Max PECsed"]
            if areic_comparison:
                headers.append("Areic dep.")
            headers.append("Route")
            if "Type" in all_keys:
                headers.append("Type")
            
            # Flatten the data
            for project, rows in all_data.items():
                for row in rows:
                    flat_row = {"Project": project}
                    for key in headers:
                        if key == "Project":
                            # Skip Project as it's already set
                            continue
                        if key in row:
                            flat_row[key] = row[key]
                        else:
                            flat_row[key] = ""
                    all_rows.append(flat_row)
        
        return jsonify({
            'data': all_rows,
            'header': headers,
            'rac_value': rac_value,
            'row_count': len(all_rows),
            'errors': errors
        })
        
    except Exception as e:
        return jsonify({'error': f'Error extracting data: {str(e)}'})

@toxswaex_bp.route('/export_excel', methods=['POST'])
def toxswaex_export_excel():
    try:
        if not toxswa_extractor.all_data:
            return jsonify({'error': 'No data to export'})
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            filepath = tmp_file.name
        
        # Export to Excel (summary sheet will be created if batch_mode and summary_mode are enabled)
        success = toxswa_extractor.export_to_excel(filepath)
        
        if not success:
            return jsonify({'error': 'Failed to export Excel file'})
        
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
def toxswaex_get_table_data():
    try:
        compound_type = request.args.get('compound_type', 'Parent')
        sort_by = request.args.get('sort_by', 'Filename')
        
        table_data, headers = toxswa_extractor.get_table_data(compound_type, sort_by)
        
        return jsonify({
            'data': table_data,
            'header': headers
        })
        
    except Exception as e:
        return jsonify({'error': f'Error getting table data: {str(e)}'})

# PEARLex Blueprint
pearlex_bp = Blueprint('pearlex', __name__,
                       template_folder='pearlex/templates',
                       static_folder='pearlex/static')

@pearlex_bp.route('/')
def pearlex_index():
    return render_template('pearlex/index.html')

@pearlex_bp.route('/scan_directory', methods=['POST'])
def pearlex_scan_directory():
    try:
        data = request.get_json()
        directory = data.get('directory', '').strip()
        
        if not directory:
            return jsonify({'error': 'Please provide a directory path'})
        
        if not os.path.exists(directory):
            return jsonify({'error': f'Directory does not exist: {directory}'})
        
        # Scan for .sum files using the exact logic from original PEARLex
        files = pearl_extractor.scan_directory(directory)
        
        return jsonify({
            'files': files,
            'main_dir': directory
        })
        
    except Exception as e:
        return jsonify({'error': f'Error scanning directory: {str(e)}'})

@pearlex_bp.route('/extract_data', methods=['POST'])
def pearlex_extract_data():
    try:
        data = request.get_json()
        main_dir = data.get('main_dir', '')
        selected_files = data.get('selected_files', [])
        compound_type = data.get('compound_type', 'Parent')
        sort_by = data.get('sort_by', 'Filename')
        limit_value = data.get('limit_value', None)
        
        if not main_dir:
            return jsonify({'error': 'No main directory specified'})
        
        if not selected_files:
            return jsonify({'error': 'No files selected'})
        
        # Convert limit value to float if provided
        if limit_value:
            try:
                limit_value = float(limit_value)
            except ValueError:
                return jsonify({'error': 'Invalid limit value'})
        
        # Extract data using the exact logic from original PEARLex
        table_data = pearl_extractor.extract_data(selected_files)
        
        # Get filtered and sorted data for display
        filtered_data = pearl_extractor.get_table_data(compound_type, sort_by, limit_value)
        
        return jsonify({
            'data': filtered_data,
            'header': ["Project", "Filename", "Compound Type", "Scenario", "Compound", "80th Percentile (µg/L)"],
            'limit_value': limit_value,
            'row_count': len(filtered_data)
        })
        
    except Exception as e:
        return jsonify({'error': f'Error extracting data: {str(e)}'})

@pearlex_bp.route('/export_excel', methods=['POST'])
def pearlex_export_excel():
    try:
        data = request.get_json()
        batch_mode = data.get('batch_mode', False)
        limit_value = data.get('limit_value', None)
        
        # Convert limit value to float if provided
        if limit_value:
            try:
                limit_value = float(limit_value)
            except ValueError:
                return jsonify({'error': 'Invalid limit value'})
        
        if batch_mode:
            # Export batches
            success, result = pearl_extractor.export_batches(limit_value)
        else:
            # Export single mode
            success, result = pearl_extractor.export_to_excel_single(limit_value)
        
        if not success:
            return jsonify({'error': result})
        
        # Send file
        return send_file(
            result,
            as_attachment=True,
            download_name='pearl_extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return jsonify({'error': f'Error exporting Excel: {str(e)}'})

@pearlex_bp.route('/add_to_batch', methods=['POST'])
def pearlex_add_to_batch():
    try:
        data = request.get_json()
        batch_name = data.get('batch_name', None)
        
        success, message = pearl_extractor.add_to_batch(batch_name)
        
        return jsonify({
            'success': success,
            'message': message,
            'batches': pearl_extractor.get_batches()
        })
        
    except Exception as e:
        return jsonify({'error': f'Error adding to batch: {str(e)}'})

@pearlex_bp.route('/clear_data', methods=['POST'])
def pearlex_clear_data():
    try:
        success, message = pearl_extractor.clear_data()
        
        return jsonify({
            'success': success,
            'message': message
        })
        
    except Exception as e:
        return jsonify({'error': f'Error clearing data: {str(e)}'})

@pearlex_bp.route('/clear_batches', methods=['POST'])
def pearlex_clear_batches():
    try:
        success, message = pearl_extractor.clear_batches()
        
        return jsonify({
            'success': success,
            'message': message
        })
        
    except Exception as e:
        return jsonify({'error': f'Error clearing batches: {str(e)}'})

@pearlex_bp.route('/get_table_data')
def pearlex_get_table_data():
    try:
        compound_type = request.args.get('compound_type', 'Parent')
        sort_by = request.args.get('sort_by', 'Filename')
        limit_value = request.args.get('limit_value', None)
        
        # Convert limit value to float if provided
        if limit_value:
            try:
                limit_value = float(limit_value)
            except ValueError:
                limit_value = None
        
        table_data = pearl_extractor.get_table_data(compound_type, sort_by, limit_value)
        
        return jsonify({
            'data': table_data,
            'header': ["Project", "Filename", "Compound Type", "Scenario", "Compound", "80th Percentile (µg/L)"]
        })
        
    except Exception as e:
        return jsonify({'error': f'Error getting table data: {str(e)}'})

# Register blueprints
app.register_blueprint(pelmoex_bp, url_prefix='/pelmoex')
app.register_blueprint(toxswaex_bp, url_prefix='/toxswaex')
app.register_blueprint(pearlex_bp, url_prefix='/pearlex')

@app.route('/')
def index():
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True) 