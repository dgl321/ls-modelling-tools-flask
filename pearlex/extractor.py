import os
import re
import xlsxwriter
import tempfile
import json

class PEARLexExtractor:
    def __init__(self):
        self.all_data = []
        self.main_dir = ""
        
    def scan_directory(self, directory_path):
        """Scan directory for .sum files and return project list"""
        projects = []
        sum_filepaths = []
        
        if not os.path.exists(directory_path):
            return projects
            
        # Scan main directory
        for f in os.listdir(directory_path):
            if f.endswith(".sum"):
                sum_filepaths.append(os.path.join(directory_path, f))
                
        # Scan subdirectories
        for sub in os.listdir(directory_path):
            subp = os.path.join(directory_path, sub)
            if os.path.isdir(subp):
                for ff in os.listdir(subp):
                    if ff.endswith(".sum"):
                        sum_filepaths.append(os.path.join(subp, ff))
        
        # Create project list with file information
        for path in sum_filepaths:
            filename = os.path.basename(path)
            relative_path = os.path.relpath(path, directory_path)
            projects.append({
                'filename': filename,
                'path': relative_path,
                'full_path': path
            })
            
        return projects
    
    def extract_data(self, main_dir, selected_files, limit_value=None):
        """Extract data from selected .sum files"""
        self.all_data = []
        self.main_dir = main_dir
        
        for file_info in selected_files:
            filepath = os.path.join(main_dir, file_info['path'])
            if not os.path.exists(filepath):
                continue
                
            try:
                with open(filepath, "r", encoding="ISO-8859-1") as f:
                    content = f.read()
            except Exception as e:
                print(f"Cannot read {filepath}: {str(e)}")
                continue
                
            # Extract project name
            p = re.search(r"Application_scheme\s+(\S+)", content)
            project = p.group(1) if p else "Unknown"
            
            # Extract scenario/location
            s = re.search(r"Location\s*[:]*\s*(.*)", content)
            scenario_raw = s.group(1).strip() if s else "Unknown"
            scenario = scenario_raw.capitalize()
            
            # Extract compound results
            comp_list = re.findall(r"Result_(\S+)\s+([\d.]+)", content)
            for i, (comp, val_str) in enumerate(comp_list):
                ctype = "Parent" if i == 0 else "Metabolite"
                try:
                    val = float(val_str)
                except:
                    val = 0.0
                self.all_data.append([
                    project, 
                    os.path.basename(filepath), 
                    scenario, 
                    comp, 
                    val, 
                    ctype
                ])
        
        return self.format_data_for_display(limit_value)
    
    def format_data_for_display(self, limit_value=None):
        """Format extracted data for display in the web interface"""
        if not self.all_data:
            return []
            
        # Convert limit_value to float if provided
        limit_val = None
        if limit_value:
            try:
                limit_val = float(limit_value)
            except:
                pass
        
        formatted_data = []
        for row in self.all_data:
            # Format: [Project, Filename, Compound Type, Scenario, Compound, Value]
            display_row = {
                'project': row[0],
                'filename': row[1],
                'compound_type': row[5],
                'scenario': row[2],
                'compound': row[3],
                'value': row[4],
                'exceeds_limit': False
            }
            
            # Check if value exceeds limit
            if limit_val is not None and row[4] > limit_val:
                display_row['exceeds_limit'] = True
                
            formatted_data.append(display_row)
            
        return formatted_data
    
    def export_to_excel(self, extracted_data, limit_value=None):
        """Export data to Excel file"""
        if not extracted_data:
            raise ValueError("No data to export")
            
        # Create temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        file_path = temp_file.name
        temp_file.close()
        
        # Convert limit_value to float if provided
        limit_val = None
        if limit_value:
            try:
                limit_val = float(limit_value)
            except:
                pass
        
        # Separate parent and metabolite data
        parents = []
        metabolites = []
        for row in extracted_data:
            if row['compound_type'] == 'Parent':
                parents.append([
                    row['project'],
                    row['filename'],
                    row['scenario'],
                    row['compound'],
                    row['value']
                ])
            else:
                metabolites.append([
                    row['project'],
                    row['filename'],
                    row['scenario'],
                    row['compound'],
                    row['value']
                ])
        
        # Sort data
        parents.sort(key=lambda x: x[1].lower())
        metabolites.sort(key=lambda x: x[1].lower())
        
        # Create Excel workbook
        wb = xlsxwriter.Workbook(file_path)
        ws = wb.add_worksheet("Results")
        
        # Define formats
        header_fmt = wb.add_format({"bold": True, "bg_color": "#82C940"})
        red_fmt = wb.add_format({"font_color": "red"})
        
        columns = [
            "Project",
            "Filename", 
            "Scenario",
            "Compound",
            "80th Percentile (µg/L)",
        ]
        
        def write_table(start_row, data_rows, title):
            col_width = [len(c) for c in columns]
            ws.write(start_row, 0, title, wb.add_format({"bold": True}))
            row_cursor = start_row + 1
            
            # Write headers
            for col_i, h in enumerate(columns):
                ws.write(row_cursor, col_i, h, header_fmt)
            row_cursor += 1
            
            # Write data
            for rdat in data_rows:
                for cc, valx in enumerate(rdat):
                    txt = str(valx)
                    if cc == 4:  # Value column
                        try:
                            fv = float(txt)
                            ws.write_number(row_cursor, cc, fv)
                            if limit_val is not None and fv > limit_val:
                                ws.write(row_cursor, cc, fv, red_fmt)
                        except:
                            ws.write(row_cursor, cc, txt)
                        col_width[cc] = max(col_width[cc], len(txt))
                    else:
                        ws.write(row_cursor, cc, txt)
                        col_width[cc] = max(col_width[cc], len(txt))
                row_cursor += 1
            
            # Set column widths
            for cidx in range(len(columns)):
                ws.set_column(cidx, cidx, col_width[cidx] + 2)
            
            return row_cursor
        
        # Write parent and metabolite tables
        nextrow = write_table(0, parents, "Parent Table")
        nextrow += 1
        write_table(nextrow, metabolites, "Metabolite Table")
        
        wb.close()
        return file_path 