import os
import re
import xlsxwriter
import tempfile

class PELMOExtractor:
    def __init__(self):
        self.all_rows = []
        self.focus_path = ""
        
    def extract_data(self, focus_path, selected_projects, limit_value=None):
        """Extract data from PELMO files"""
        try:
            self.focus_path = focus_path
            self.all_rows = []
            errors = []
            
            for project in selected_projects:
                project_path = os.path.join(focus_path, project)
                if os.path.exists(project_path):
                    try:
                        self.process_project(project_path, project, limit_value)
                    except Exception as e:
                        errors.append(f"Error processing project {project}: {str(e)}")
                else:
                    errors.append(f"Project path not found: {project_path}")
            
            # Generate header from all unique keys
            if self.all_rows:
                header = ["Project", "Crop", "Scenario"]
                all_keys = set()
                for row in self.all_rows:
                    all_keys.update(row.keys())
                header.extend(sorted(all_keys - {"Project", "Crop", "Scenario"}))
            else:
                header = []
            
            return self.all_rows, header, errors
            
        except Exception as e:
            return [], [], [f"Error extracting data: {str(e)}"]
    
    def process_project(self, project_path, project_name, limit_value=None):
        """Process a single PELMO project"""
        # Look for .pel files
        pel_files = [f for f in os.listdir(project_path) if f.endswith('.pel')]
        
        for pel_file in pel_files:
            file_path = os.path.join(project_path, pel_file)
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                
                # Extract basic information
                crop = self.extract_value(content, r"Crop:\s*([^\n]+)", "Unknown")
                scenario = self.extract_value(content, r"Scenario:\s*([^\n]+)", "Unknown")
                
                # Extract PEC values
                pec_values = self.extract_pec_values(content)
                
                # Create row data
                row_data = {
                    "Project": project_name,
                    "Crop": crop,
                    "Scenario": scenario
                }
                
                # Add PEC values
                for key, value in pec_values.items():
                    if limit_value is not None:
                        try:
                            num_value = float(value)
                            if num_value > limit_value:
                                row_data[key] = f"{value} (EXCEEDED)"
                            else:
                                row_data[key] = value
                        except ValueError:
                            row_data[key] = value
                    else:
                        row_data[key] = value
                
                self.all_rows.append(row_data)
                
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
    
    def extract_value(self, content, pattern, default_value):
        """Extract value using regex pattern"""
        match = re.search(pattern, content, re.IGNORECASE)
        return match.group(1).strip() if match else default_value
    
    def extract_pec_values(self, content):
        """Extract PEC values from PELMO file content"""
        pec_values = {}
        
        # Look for PEC patterns
        pec_patterns = [
            r"PEC\s+(\d+)\s+days:\s*([\d.]+)",
            r"PEC\s+(\d+)\s+day:\s*([\d.]+)",
            r"PEC\s+(\d+)\s+hours:\s*([\d.]+)",
            r"PEC\s+(\d+)\s+hour:\s*([\d.]+)",
        ]
        
        for pattern in pec_patterns:
            matches = re.finditer(pattern, content, re.IGNORECASE)
            for match in matches:
                time_period = match.group(1)
                value = match.group(2)
                key = f"PEC_{time_period}_days"
                pec_values[key] = value
        
        # Look for other common patterns
        other_patterns = {
            r"Maximum\s+PEC:\s*([\d.]+)": "Max_PEC",
            r"Average\s+PEC:\s*([\d.]+)": "Avg_PEC",
            r"Total\s+Deposition:\s*([\d.]+)": "Total_Deposition",
        }
        
        for pattern, key in other_patterns.items():
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                pec_values[key] = match.group(1)
        
        return pec_values
    
    def export_to_excel(self, filepath):
        """Export data to Excel file"""
        if not self.all_rows:
            return
        
        workbook = xlsxwriter.Workbook(filepath)
        
        # Create formats
        header_format = workbook.add_format({
            "bg_color": "#82C940",
            "font_color": "#000000",
            "bold": True,
            "align": "left",
        })
        
        exceeded_format = workbook.add_format({
            "font_color": "red",
            "bold": True,
        })
        
        # Create worksheet
        worksheet = workbook.add_worksheet("PELMO_Data")
        
        # Write headers
        if self.all_rows:
            headers = list(self.all_rows[0].keys())
            for col, header in enumerate(headers):
                worksheet.write(0, col, header, header_format)
            
            # Write data
            for row_idx, row in enumerate(self.all_rows, 1):
                for col_idx, header in enumerate(headers):
                    value = row.get(header, "")
                    if "EXCEEDED" in str(value):
                        worksheet.write(row_idx, col_idx, value, exceeded_format)
                    else:
                        worksheet.write(row_idx, col_idx, value)
            
            # Auto-fit columns
            for col in range(len(headers)):
                worksheet.set_column(col, col, 15)
        
        workbook.close() 