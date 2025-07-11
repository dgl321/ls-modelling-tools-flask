import os
import re
import json
import base64
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
import xlsxwriter
from xlsxwriter.utility import xl_range

app = Flask(__name__, template_folder='templates')
app.secret_key = 'pelmo_extractor_secret_key'

# Global variables to store session data
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

class PELMOExtractor:
    def __init__(self):
        self.main_dir = ""
        self.all_rows = []
        self.limit_value = None

    def extract_active_substance_and_metabolites(self, file_path):
        active_substance = None
        active_pec_value = None
        metabolites = []
        metabolite = None

        print(f"DEBUG: Reading file: {file_path}")
        
        with open(file_path, "r", encoding="ISO-8859-1") as file:
            for line in file:
                if "Results for ACTIVE SUBSTANCE" in line and "percolate at 1 m soil depth" in line:
                    match = re.search(r"Results for ACTIVE SUBSTANCE \((.*?)\)", line)
                    if match:
                        active_substance = match.group(1)
                        print(f"DEBUG: Found active substance: {active_substance}")
                if active_substance and "80 Perc." in line:
                    active_pec_value = line.split()[-1]
                    print(f"DEBUG: Found active PEC value: {active_pec_value}")
                if "Results for METABOLITE" in line and "percolate at 1 m soil depth" in line:
                    match = re.search(r"Results for METABOLITE.*?\((.*?)\)", line)
                    if match:
                        metabolite = match.group(1)
                        print(f"DEBUG: Found metabolite: {metabolite}")
                if metabolite and "80 Perc." in line:
                    pec_value = line.split()[-1]
                    metabolites.append((metabolite, pec_value))
                    print(f"DEBUG: Found metabolite PEC value: {metabolite} = {pec_value}")
                    metabolite = None

        print(f"DEBUG: Final results - Active: {active_substance}, PEC: {active_pec_value}, Metabolites: {metabolites}")
        return active_substance, active_pec_value, metabolites

    def extract_scenario_from_path(self, file_path):
        folder_name = os.path.basename(os.path.dirname(file_path))
        scenario = folder_name.split("_-_")[0] if "_-_" in folder_name else folder_name
        return scenario

    def extract_crop_from_path(self, file_path):
        folder_parts = os.path.normpath(file_path).split(os.sep)
        if len(folder_parts) >= 3:
            third_last_folder = folder_parts[-3]
            crop = third_last_folder.split(".run")[0] if ".run" in third_last_folder else third_last_folder
            crop = crop.replace("_-_", " ")
            return crop
        return None

    def convert_to_numeric(self, value):
        try:
            return float(value)
        except ValueError:
            return value

    def extract_data(self, main_dir, selected_projects, limit_value=None):
        """Extract data from PELMO directories"""
        self.main_dir = main_dir
        self.limit_value = limit_value
        all_rows = []
        all_extra_keys = set()
        errors = []

        for project_folder_name in selected_projects:
            project_path = os.path.join(self.main_dir, project_folder_name)
            
            # Get all crop folders
            crop_folders = [
                d for d in os.listdir(project_path)
                if os.path.isdir(os.path.join(project_path, d)) and d.endswith(".run")
            ]
            
            if not crop_folders:
                errors.append(f"No crop folders found in project '{project_folder_name}'")
                continue

            for crop_folder in crop_folders:
                crop_folder_path = os.path.join(project_path, crop_folder)
                
                # Look for scenario folders
                scenario_folders = [
                    d for d in os.listdir(crop_folder_path)
                    if os.path.isdir(os.path.join(crop_folder_path, d)) and "_-_" in d and d.endswith(".run")
                ]
                
                if not scenario_folders:
                    errors.append(f"No scenario folders found in crop folder '{crop_folder}'")
                    continue

                for scenario_folder in scenario_folders:
                    scenario_folder_path = os.path.join(crop_folder_path, scenario_folder)
                    period_plm_path = os.path.join(scenario_folder_path, "period.plm")
                    
                    if not os.path.exists(period_plm_path):
                        errors.append(f"'period.plm' not found in scenario folder '{scenario_folder}'")
                        continue
                    
                    active_substance, active_pec_value, metabolites = self.extract_active_substance_and_metabolites(period_plm_path)
                    print(f"DEBUG: Extraction results for {scenario_folder}:")
                    print(f"  Active substance: {active_substance}")
                    print(f"  Active PEC value: {active_pec_value}")
                    print(f"  Metabolites: {metabolites}")
                    
                    if not active_substance or not active_pec_value:
                        print(f"DEBUG: Skipping {scenario_folder} - missing active substance or PEC value")
                        continue
                    
                    row = {}
                    row["Project"] = project_folder_name
                    row["Crop"] = self.extract_crop_from_path(period_plm_path)
                    row["Scenario"] = self.extract_scenario_from_path(period_plm_path)
                    
                    active_col = f"{active_substance} µg/l"
                    row[active_col] = self.convert_to_numeric(active_pec_value)
                    all_extra_keys.add(active_col)
                    
                    for met, pec in metabolites:
                        colname = f"{met} µg/l"
                        row[colname] = self.convert_to_numeric(pec)
                        all_extra_keys.add(colname)
                    
                    print(f"DEBUG: Created row: {row}")
                    all_rows.append(row)

        # Build the table header
        header = ["Project", "Crop", "Scenario"] + sorted(all_extra_keys)
        print(f"DEBUG: Final header: {header}")
        print(f"DEBUG: All extra keys: {sorted(all_extra_keys)}")
        
        for row in all_rows:
            for key in header:
                if key not in row:
                    row[key] = ""
        
        print(f"DEBUG: Final rows count: {len(all_rows)}")
        for i, row in enumerate(all_rows):
            print(f"DEBUG: Row {i}: {row}")

        self.all_rows = all_rows
        return all_rows, header, errors

    def export_to_excel(self, filepath):
        """Export data to Excel file"""
        workbook = xlsxwriter.Workbook(filepath)
        
        # Group rows by project
        projects = {}
        for row in self.all_rows:
            proj = row["Project"]
            projects.setdefault(proj, []).append(row)

        used_sheet_names = set()
        for project, rows in projects.items():
            # Create sheet name - remove .run extension and clean up
            sheet_name = project[:-4] if project.lower().endswith(".run") else project
            # Replace problematic characters and ensure it's under 31 chars
            sheet_name = sheet_name.replace('_-_', '_').replace('-', '_')
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:28]  # Leave room for potential numbering
            orig_sheet_name = sheet_name
            count = 1
            while sheet_name in used_sheet_names:
                sheet_name = f"{orig_sheet_name}_{count}"
                if len(sheet_name) > 31:
                    sheet_name = f"{orig_sheet_name[:25]}_{count}"
                count += 1
            used_sheet_names.add(sheet_name)

            worksheet = workbook.add_worksheet(sheet_name)
            
            # Re-calculate header for this project
            header = ["Project", "Crop", "Scenario"]
            extra_keys = set()
            for row in rows:
                for key in row.keys():
                    if key not in ["Project", "Crop", "Scenario"]:
                        extra_keys.add(key)
            header.extend(sorted(extra_keys))

            header_format = workbook.add_format({"bold": True, "bg_color": "#DFF0D8", "border": 1})
            for col, header_text in enumerate(header):
                worksheet.write(0, col, header_text, header_format)

            for r, row in enumerate(rows, start=1):
                for col, key in enumerate(header):
                    value = row.get(key, "")
                    try:
                        num_value = float(value)
                        worksheet.write_number(r, col, num_value)
                    except ValueError:
                        worksheet.write(r, col, value)

            # Adjust column widths
            for col in range(len(header)):
                max_width = len(header[col])
                for row in rows:
                    cell_text = str(row.get(header[col], ""))
                    max_width = max(max_width, len(cell_text))
                worksheet.set_column(col, col, max_width + 2)

            # Apply conditional formatting if limit is set
            if self.limit_value is not None:
                red_format = workbook.add_format({"font_color": "red"})
                green_format = workbook.add_format({"font_color": "green"})
                for col in range(3, len(header)):
                    cell_range = xl_range(1, col, len(rows), col)
                    worksheet.conditional_format(cell_range, {
                        "type": "cell",
                        "criteria": ">=",
                        "value": self.limit_value,
                        "format": red_format,
                    })
                    worksheet.conditional_format(cell_range, {
                        "type": "cell",
                        "criteria": "<",
                        "value": self.limit_value,
                        "format": green_format,
                    })

        workbook.close()

# Global extractor instance
extractor = PELMOExtractor()

@app.route('/')
def index():
    return render_template('pelmoex/index.html')

@app.route('/scan_directory', methods=['POST'])
def scan_directory():
    """Scan directory for FOCUS folder and return project list"""
    try:
        data = request.get_json()
        directory = data.get('directory')
        
        if not directory or not os.path.exists(directory):
            return jsonify({'error': 'Directory does not exist'}), 400
        
        # Look for FOCUS folder
        focus_folder = None
        for item in os.listdir(directory):
            if os.path.isdir(os.path.join(directory, item)) and item.upper() == "FOCUS":
                focus_folder = os.path.join(directory, item)
                break
        
        if not focus_folder:
            return jsonify({'error': 'FOCUS folder not found in the selected directory'}), 400
        
        # List project folders
        project_folders = [
            f for f in os.listdir(focus_folder)
            if os.path.isdir(os.path.join(focus_folder, f)) and f.endswith(".run")
        ]
        
        return jsonify({
            'focus_path': focus_folder,
            'projects': sorted(project_folders)
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/extract_data', methods=['POST'])
def extract_data():
    """Extract data from selected projects"""
    try:
        data = request.get_json()
        focus_path = data.get('focus_path')
        selected_projects = data.get('selected_projects', [])
        limit_value = data.get('limit_value')
        
        print(f"DEBUG: focus_path = {focus_path}")
        print(f"DEBUG: selected_projects = {selected_projects}")
        
        if not focus_path or not selected_projects:
            return jsonify({'error': 'Missing required parameters'}), 400
        
        # Convert limit value
        if limit_value:
            if limit_value == "0.1":
                limit_value = 0.1
            elif limit_value == "0.001":
                limit_value = 0.001
            else:
                limit_value = None
        
        # Extract data
        rows, header, errors = extractor.extract_data(focus_path, selected_projects, limit_value)
        
        if not rows:
            return jsonify({'error': 'No data extracted. Please check your project folders.'}), 400
        
        # Store data in session for export
        session['extracted_data'] = {
            'rows': rows,
            'header': header,
            'limit_value': limit_value
        }
        
        return jsonify({
            'success': True,
            'data': rows,
            'header': header,
            'errors': errors,
            'row_count': len(rows)
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/export_excel', methods=['POST'])
def export_excel():
    """Export data to Excel file"""
    try:
        if 'extracted_data' not in session:
            return jsonify({'error': 'No data to export'}), 400
        
        data = session['extracted_data']
        rows = data['rows']
        header = data['header']
        limit_value = data.get('limit_value')
        
        # Create temporary file
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_path = tmp_file.name
        
        # Export to Excel
        extractor.all_rows = rows
        extractor.limit_value = limit_value
        extractor.export_to_excel(tmp_path)
        
        return send_file(
            tmp_path,
            as_attachment=True,
            download_name='pelmo_extracted_data.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/get_table_data')
def get_table_data():
    """Get current table data"""
    if 'extracted_data' in session:
        return jsonify(session['extracted_data'])
    return jsonify({'data': [], 'header': []})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 