import os
import re
import xlsxwriter
from io import BytesIO
from flask import send_file

class PearlGroundwaterExtractor:
    def __init__(self):
        self.main_dir = ""
        self.sum_filepaths = []
        self.all_data = []
        self.batches = []

    def scan_directory(self, directory_path):
        """Scan directory for .sum files and return list of found files"""
        self.main_dir = directory_path
        self.sum_filepaths.clear()
        
        if not self.main_dir or not os.path.exists(self.main_dir):
            return []
            
        # Scan main directory
        for f in os.listdir(self.main_dir):
            if f.endswith(".sum"):
                self.sum_filepaths.append(os.path.join(self.main_dir, f))
        
        # Scan subdirectories
        for sub in os.listdir(self.main_dir):
            subp = os.path.join(self.main_dir, sub)
            if os.path.isdir(subp):
                for ff in os.listdir(subp):
                    if ff.endswith(".sum"):
                        self.sum_filepaths.append(os.path.join(subp, ff))
        
        return [os.path.basename(path) for path in self.sum_filepaths]

    def extract_data(self, selected_files):
        """Extract data from selected .sum files"""
        self.all_data.clear()
        
        for filename in selected_files:
            # Find the full path for this filename
            file_path = None
            for path in self.sum_filepaths:
                if os.path.basename(path) == filename:
                    file_path = path
                    break
            
            if not file_path:
                continue
                
            try:
                with open(file_path, "r", encoding="ISO-8859-1") as f:
                    content = f.read()
            except Exception as e:
                print(f"Cannot read {file_path}: {str(e)}")
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
                    os.path.basename(file_path), 
                    scenario, 
                    comp, 
                    val, 
                    ctype
                ])
        
        return self.get_table_data()

    def get_table_data(self, compound_type="Parent", sort_by="Filename", limit_val=None):
        """Get filtered and sorted data for table display"""
        # Filter by compound type
        data_for_table = [r for r in self.all_data if r[5] == compound_type]
        
        # Sort data
        if sort_by == "Filename":
            data_for_table.sort(key=lambda x: x[1].lower())
        elif sort_by == "Compound":
            data_for_table.sort(key=lambda x: x[3].lower())
        elif sort_by == "Scenario":
            data_for_table.sort(key=lambda x: x[2].lower())
        
        # Prepare table data
        table_data = []
        for row_data in data_for_table:
            display = [
                row_data[0],  # Project
                row_data[1],  # Filename
                row_data[5],  # Compound Type
                row_data[2],  # Scenario
                row_data[3],  # Compound
                row_data[4],  # 80th Percentile
            ]
            
            # Check if value exceeds limit for highlighting
            exceeds_limit = False
            if limit_val is not None:
                try:
                    if float(row_data[4]) > limit_val:
                        exceeds_limit = True
                except:
                    pass
            
            table_data.append({
                'data': display,
                'exceeds_limit': exceeds_limit
            })
        
        return table_data

    def add_to_batch(self, batch_name=None):
        """Add current data to batch"""
        if not self.all_data:
            return False, "No data to add to batch"
        
        if not batch_name or not batch_name.strip():
            batch_name = f"Batch_{len(self.batches) + 1}"
        
        new_copy = [row[:] for row in self.all_data]
        self.batches.append((batch_name, new_copy))
        
        return True, f"Batch '{batch_name}' added successfully"

    def clear_data(self):
        """Clear all extracted data"""
        self.all_data.clear()
        return True, "Data cleared"

    def clear_batches(self):
        """Clear all batches"""
        self.batches.clear()
        return True, "Batches cleared"

    def export_to_excel_single(self, limit_val=None):
        """Export single mode data to Excel"""
        if not self.all_data:
            return False, "No data to export"
        
        try:
            # Separate parent and metabolite data
            parents = []
            mets = []
            for row in self.all_data:
                if row[5] == "Parent":
                    parents.append(row)
                else:
                    mets.append(row)
            
            parents.sort(key=lambda x: x[1].lower())
            mets.sort(key=lambda x: x[1].lower())
            
            # Create Excel file in memory
            output = BytesIO()
            wb = xlsxwriter.Workbook(output)
            ws = wb.add_worksheet("Results")
            
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
                    reorder = [rdat[0], rdat[1], rdat[2], rdat[3], rdat[4]]
                    for cc, valx in enumerate(reorder):
                        txt = str(valx)
                        if cc == 4:  # 80th Percentile column
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
            write_table(nextrow, mets, "Metabolite Table")
            
            wb.close()
            output.seek(0)
            
            return True, output
            
        except Exception as e:
            return False, f"Export error: {str(e)}"

    def export_batches(self, limit_val=None):
        """Export batch mode data to Excel"""
        if not self.batches:
            return False, "No batches to export"
        
        try:
            output = BytesIO()
            wb = xlsxwriter.Workbook(output)
            hfmt = wb.add_format({"bold": True, "bg_color": "#82C940"})
            redf = wb.add_format({"font_color": "red"})
            
            columns = [
                "Project",
                "Filename",
                "Scenario", 
                "Compound",
                "80th Percentile (µg/L)",
            ]

            def write_section(ws, start, data_rows, title):
                cw = [len(x) for x in columns]
                ws.write(start, 0, title, wb.add_format({"bold": True}))
                rowcur = start + 1
                
                # Write headers
                for c_i, h in enumerate(columns):
                    ws.write(rowcur, c_i, h, hfmt)
                rowcur += 1
                
                # Write data
                for rowd in data_rows:
                    reorder = [rowd[0], rowd[1], rowd[2], rowd[3], rowd[4]]
                    for cc, vv in enumerate(reorder):
                        txt = str(vv)
                        if cc == 4:  # 80th Percentile column
                            try:
                                fv = float(txt)
                                ws.write_number(rowcur, cc, fv)
                                if limit_val is not None and fv > limit_val:
                                    ws.write(rowcur, cc, fv, redf)
                            except:
                                ws.write(rowcur, cc, txt)
                            cw[cc] = max(cw[cc], len(txt))
                        else:
                            ws.write(rowcur, cc, txt)
                            cw[cc] = max(cw[cc], len(txt))
                    rowcur += 1
                
                # Set column widths
                for c_col in range(len(columns)):
                    ws.set_column(c_col, c_col, cw[c_col] + 2)
                
                return rowcur

            # Write each batch to separate worksheet
            for sheet_name, rows in self.batches:
                p = []
                m = []
                for r in rows:
                    if r[5] == "Parent":
                        p.append(r)
                    else:
                        m.append(r)
                
                p.sort(key=lambda x: x[1].lower())
                m.sort(key=lambda x: x[1].lower())
                
                wsheet = wb.add_worksheet(sheet_name[:31])  # Excel sheet name limit
                next_r = write_section(wsheet, 0, p, "Parent Table")
                next_r += 1
                write_section(wsheet, next_r, m, "Metabolite Table")
            
            wb.close()
            output.seek(0)
            
            return True, output
            
        except Exception as e:
            return False, f"Export error: {str(e)}"

    def get_available_files(self):
        """Get list of available .sum files"""
        return [os.path.basename(path) for path in self.sum_filepaths]

    def get_batches(self):
        """Get list of current batches"""
        return [name for name, _ in self.batches] 