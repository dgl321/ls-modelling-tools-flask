import os
import re
import xlsxwriter
import tempfile

class TOXSWAExtractor:
    def __init__(self):
        self.all_data = {}
        self.project_shortcodes = {}
        self.main_dir = ""
        self.areic_comparison_enabled = False
        
    def extract_data(self, main_dir, selected_projects, selected_files=None, rac_value=None, areic_comparison=False):
        """Extract data from TOXSWA files"""
        try:
            self.main_dir = main_dir
            self.areic_comparison_enabled = areic_comparison
            self.all_data.clear()
            self.project_shortcodes.clear()
            
            errors = []
            
            for project in selected_projects:
                project_path = os.path.join(main_dir, project, "toxswa")
                if os.path.exists(project_path):
                    try:
                        self.process_files(project_path, project, selected_files)
                    except Exception as e:
                        errors.append(f"Error processing project {project}: {str(e)}")
                else:
                    errors.append(f"Project path not found: {project_path}")
            
            return self.all_data, errors
            
        except Exception as e:
            return {}, [f"Error extracting data: {str(e)}"]
    
    def process_files(self, folder_path, project_name, selected_files=None):
        """Process TOXSWA .sum files"""
        files = sorted([f for f in os.listdir(folder_path) if f.endswith(".sum")])
        if selected_files:
            files = [f for f in files if f in selected_files]

        project_root = os.path.dirname(folder_path)
        shortcode = self.extract_shortcode(project_root)
        self.project_shortcodes[project_name] = shortcode if shortcode else "Step 3"

        all_rows = []
        for filename in files:
            file_path = os.path.join(folder_path, filename)
            try:
                with open(file_path, "r", encoding="ISO-8859-1") as f:
                    content = f.read()
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
                continue

            # Extract Areic mean deposition value
            areic_value = self.extract_areic_mean_deposition(content)

            scenario = "Unknown"
            waterbody = "Unknown"
            scenario_match = re.search(r"\* Scenario\s*:\s*([^\r\n]+)", content)
            if scenario_match:
                scenario_raw = scenario_match.group(1).strip()
                if "_" in scenario_raw:
                    parts = scenario_raw.split("_", 1)
                    scenario = parts[0].strip()
                    waterbody = parts[1].strip().capitalize()
                else:
                    scenario = scenario_raw
                    wb_match = re.search(r"\* Water Body Type\s*:\s*(\S+)", content)
                    waterbody = wb_match.group(1).capitalize() if wb_match else "Unknown"

            app_dates = []
            app_section = re.search(r"Appl\.No\s+Date/Hour.*?\n(.*?)\n\n", content, re.DOTALL)
            if app_section:
                for line in app_section.group(1).split("\n"):
                    date_match = re.search(r"\d{2}-[A-Za-z]{3}-\d{4}-\d{2}h\d{2}", line)
                    if date_match:
                        app_dates.append(date_match.group().strip())

            max_date = ""
            max_match = re.search(r"Global max.*?(\d{2}-[A-Za-z]{3}-\d{4}-\d{2}h\d{2})", content, re.IGNORECASE | re.DOTALL)
            if max_match:
                max_date = max_match.group(1).strip()

            if max_date in app_dates:
                route = "Spray Drift"
            else:
                scenario_code = scenario[:1].upper() if scenario else ""
                if scenario_code == "D":
                    route = "Drainage"
                elif scenario_code == "R":
                    route = "Runoff"
                else:
                    route = "Spray Drift"

            parent_compound = self.extract_value(content, r"\* Substance\s*:\s*(\S+)", "Unknown")
            parent_max_sw = self.extract_float(content, r"Global max.*?([\d.]+)")
            pecsed_match = re.search(r"PEC in sediment of substance:\s*\S+.*?Global max\s+([<]?\s*\d+(?:\.\d+)?(?:e[+-]?\d+)?)", content, re.DOTALL)
            parent_max_sed_str = pecsed_match.group(1).strip() if pecsed_match else "0"
            parent_max_sed = self.parse_value(parent_max_sed_str)

            row_parent = {
                "Filename": filename,
                "Scenario": scenario,
                "Waterbody": waterbody,
                "Compound": parent_compound,
                "Max PECsw": self.format_for_display(parent_max_sw),
                "Max PECsed": self.format_for_display(parent_max_sed),
                "Areic mean deposition": areic_value,
                "Route": route,
                "Type": "Parent",
                "ApplicationDates": app_dates,
                "FilePath": file_path,
            }
            all_rows.append(row_parent)

            # Process metabolites
            subs = []
            for match2 in re.finditer(r"\* Substance\s+\d+:\s+(\S+)", content):
                sub = match2.group(1).strip()
                if sub != parent_compound:
                    subs.append(sub)
            if not subs:
                soil = self.extract_value(content, r"\* Soil metabolite:\s*(\S+)", "")
                if soil and soil != parent_compound:
                    subs.append(soil)
            for sub in subs:
                max_sw_str = self.extract_value(content, rf"\* Table:\s*PEC in water layer of substance:\s+{re.escape(sub)}.*?Global max\s+([<]?\s*\S+)", "0")
                max_sed_str = self.extract_value(content, rf"\* Table:\s*PEC in sediment of substance:\s+{re.escape(sub)}.*?Global max\s+([<]?\s*\S+)", "0")
                row_met = {
                    "Filename": filename,
                    "Scenario": scenario,
                    "Waterbody": waterbody,
                    "Compound": sub,
                    "Max PECsw": self.format_for_display(self.parse_value(max_sw_str)),
                    "Max PECsed": self.format_for_display(self.parse_value(max_sed_str)),
                    "Route": route,
                    "Type": "Metabolite",
                    "ApplicationDates": app_dates,
                    "FilePath": file_path,
                }
                all_rows.append(row_met)

        if all_rows:
            self.all_data[project_name] = all_rows
    
    def extract_shortcode(self, folder_path):
        """Extract shortcode from SWAN_log.txt"""
        swan_log_path = os.path.join(folder_path, "SWAN_log.txt")
        if not os.path.exists(swan_log_path):
            return ""
        try:
            with open(swan_log_path, "r", encoding="ISO-8859-1") as f:
                content = f.read()
        except Exception as e:
            print(f"Error reading {swan_log_path}: {e}")
            return ""

        buffer = ""
        nozzle = ""
        vfs = ""
        vfs_flag = ""

        spray_section = re.search(r"Spray drift mitigation.*?(?=Run-off mitigation)", content, re.DOTALL | re.IGNORECASE)
        spray_content = spray_section.group(0) if spray_section else content

        m = re.search(r"Buffer\s*width\s*\(m\)\s*:\s*(\d+)", spray_content, re.IGNORECASE)
        if m:
            buffer = f"{m.group(1)}b"

        m = re.search(r"Nozzle\s*reduction\s*\(\%\)\s*:\s*(\d+)", spray_content, re.IGNORECASE)
        if m:
            nozzle_val = int(m.group(1))
            if nozzle_val > 0:
                nozzle = f"{nozzle_val}%"

        runoff_section = re.search(r"Run-off mitigation.*?(?=Dry deposition)", content, re.DOTALL | re.IGNORECASE)
        if runoff_section:
            runoff_content = runoff_section.group(0)
            if re.search(r"Reduction\s*run-?off\s*mode:\s*VfsMod", runoff_content, re.IGNORECASE):
                m = re.search(r"Filter\s*strip\s*buffer\s*width\s*:\s*(\d+)", runoff_content, re.IGNORECASE)
                if m:
                    vfs = f"{m.group(1)}vfs"
                vfs_flag = " VFSMOD"

            elif re.search(r"Reduction\s*run-?off\s*mode:\s*ManualReduction", runoff_content, re.IGNORECASE):
                fr_volume_match = re.search(r"Fractional\s+reduction\s+in\s+run-off\s+volume\s*:\s*([\d.]+)", runoff_content, re.IGNORECASE)
                if fr_volume_match:
                    try:
                        vol_value = float(fr_volume_match.group(1))
                    except ValueError:
                        vol_value = None
                    if vol_value is not None:
                        if abs(vol_value - 0.6) < 1E-6:
                            vfs = "10vfs"
                        elif abs(vol_value - 0.8) < 1E-6:
                            vfs = "20vfs"

        return f"{buffer}{vfs}{nozzle}{vfs_flag}"
    
    def extract_areic_mean_deposition(self, content):
        """Extract areic mean deposition value"""
        m = re.search(r"Areic mean deposition\s*\(mg\.m-2\).*?\n\s*\d+\s+[^\n]*\s+([\d\.Ee-]+)", content, re.IGNORECASE)
        return m.group(1).strip() if m else "N/A"
    
    def extract_value(self, content, pattern, default_value):
        """Extract value using regex pattern"""
        m = re.search(pattern, content, re.DOTALL)
        return m.group(1).strip() if m else default_value
    
    def extract_float(self, content, pattern):
        """Extract float value using regex pattern"""
        m = re.search(pattern, content)
        if not m:
            return None
        val_str = m.group(1).strip()
        return self.parse_value(val_str)
    
    def parse_value(self, value_str):
        """Parse numeric value from string"""
        if not value_str:
            return None
        value_str = value_str.strip()
        if value_str.startswith("<"):
            raw = value_str.lstrip("<").strip()
            try:
                numeric = float(raw)
                if numeric < 1E-6:
                    return "<1E-06"
                return "<1E-06"
            except:
                return "<1E-06"
        try:
            numeric_val = float(value_str)
            if numeric_val < 1E-6:
                return "<1E-06"
            return numeric_val
        except:
            return None
    
    def format_for_display(self, val, compound_type="Parent"):
        """Format value for display"""
        PARENT_DECIMALS_SMALL = 4
        PARENT_DECIMALS_LARGE = 2
        METABOLITE_DECIMALS_SMALL = 6
        METABOLITE_DECIMALS_LARGE = 2

        try:
            num = float(val)
            if num <= 1E-6:
                return "<1E-06"
            
            if compound_type == "Parent":
                decimals = PARENT_DECIMALS_SMALL if num < 1 else PARENT_DECIMALS_LARGE
            else:
                decimals = METABOLITE_DECIMALS_SMALL if num < 1 else METABOLITE_DECIMALS_LARGE
                
            return f"{num:.{decimals}f}"
        except:
            return str(val)
    
    def get_table_data(self, compound_type="Parent", sort_by="Filename"):
        """Get table data for display"""
        if not self.all_data:
            return [], []
        
        # Collect all rows
        all_rows = []
        for project, rows in self.all_data.items():
            for row in rows:
                if row.get("Type", "") == compound_type:
                    all_rows.append({
                        "Project": project,
                        "Filename": row["Filename"],
                        "Compound": row["Compound"],
                        "Scenario": row["Scenario"],
                        "Waterbody": row["Waterbody"],
                        "Max PECsw": row.get("Max PECsw", "0"),
                        "Max PECsed": row.get("Max PECsed", "0"),
                        "Areic dep.": row.get("Areic mean deposition", "N/A") if self.areic_comparison_enabled else None,
                        "Route": row["Route"]
                    })
        
        if not all_rows:
            return [], []
        
        # Sort rows
        if sort_by == "Filename":
            all_rows.sort(key=lambda x: x["Filename"])
        elif sort_by == "Compound":
            all_rows.sort(key=lambda x: x["Compound"].upper())
        elif sort_by == "Scenario":
            all_rows.sort(key=lambda x: x["Scenario"].upper())
        elif sort_by == "File number":
            def get_sum_number(filename):
                m = re.search(r"(\d+)\.sum$", filename)
                return int(m.group(1)) if m else 999999999
            all_rows.sort(key=lambda x: get_sum_number(x["Filename"]))
        
        # Determine headers
        headers = ["Project", "Filename", "Compound", "Scenario", "Waterbody", "Max PECsw", "Max PECsed"]
        if self.areic_comparison_enabled:
            headers.append("Areic dep.")
        headers.append("Route")
        
        return all_rows, headers
    
    def export_to_excel(self, filepath):
        """Export data to Excel file"""
        if not self.all_data:
            return
        
        workbook = xlsxwriter.Workbook(filepath)
        
        # Create formats
        right_align = workbook.add_format({"align": "right"})
        header_format = workbook.add_format({
            "bg_color": "#82C940",
            "font_color": "#000000",
            "bold": True,
            "align": "left",
        })
        
        # Initialize sheet names tracking
        existing_sheet_names = set()
        
        for project, rows in self.all_data.items():
            # Create safe sheet name
            worksheet = workbook.add_worksheet(self.safe_sheet_name(project, existing_sheet_names))
            
            # Sort rows
            sorted_rows = sorted(rows, key=lambda r: (
                r["Compound"].upper(),
                int(re.search(r"(\d+)\.sum$", r["Filename"]).group(1)) if re.search(r"(\d+)\.sum$", r["Filename"]) else 999999999
            ))
            
            # Write headers
            headers = ["Filename", "Compound", "Scenario", "Waterbody", "Max PECsw (μg/L)", "Max PECsed (μg/L)", "Route of entry"]
            for col, header in enumerate(headers):
                worksheet.write(0, col, header, header_format)
            
            # Write data
            for row_idx, row in enumerate(sorted_rows, 1):
                data = [
                    row["Filename"],
                    row["Compound"],
                    row["Scenario"],
                    row["Waterbody"],
                    row.get("Max PECsw", "0"),
                    row.get("Max PECsed", "0"),
                    row["Route"]
                ]
                
                for col, value in enumerate(data):
                    if col >= 4:  # Numeric columns
                        worksheet.write(row_idx, col, value, right_align)
                    else:
                        worksheet.write(row_idx, col, value)
            
            worksheet.set_column(0, len(headers)-1, 15)
        
        workbook.close()
    
    def safe_sheet_name(self, name, existing_names):
        """Create safe Excel sheet name"""
        invalid_chars = "[]:*?/\\"
        for ch in invalid_chars:
            name = name.replace(ch, "")
        name = name.strip()
        if len(name) > 31:
            name = name[:31]
        base_name = name
        counter = 1
        while name in existing_names:
            suffix = str(counter)
            truncated_length = 31 - len(suffix)
            name = base_name[:truncated_length] + suffix
            counter += 1
        existing_names.add(name)
        return name 