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
        self.summary_mode = False
        self.project_order = []
        
    def extract_data(self, main_dir, selected_projects, selected_files=None, rac_value=None, areic_comparison=False, summary_mode=False, project_order=None):
        """Extract data from TOXSWA files"""
        try:
            self.main_dir = main_dir
            self.areic_comparison_enabled = areic_comparison
            self.summary_mode = summary_mode
            self.project_order = project_order or []
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
    
    def collect_step3_areic_map(self):
        """
        Build a dictionary mapping each unique (scenario, waterbody, compound, type)
        to the areic deposition value from files whose project shortcode includes "Step 3".
        """
        step3_map = {}
        for project, rows in self.all_data.items():
            # Check if this project is a "Step 3" project.
            if "Step 3" in self.project_shortcodes.get(project, ""):
                for row in rows:
                    key = (row["Scenario"], row["Waterbody"], row["Compound"], row["Type"])
                    areic_str = row.get("Areic mean deposition", "N/A")
                    if areic_str != "N/A":
                        try:
                            areic_val = float(areic_str)
                        except Exception:
                            areic_val = None
                        if areic_val and areic_val > 0:
                            # Only store the first valid areic value found per key.
                            if key not in step3_map:
                                step3_map[key] = areic_val
        return step3_map

    def _convert_to_float(self, value):
        try:
            if isinstance(value, str):
                if value.startswith("<"):
                    number_part = value.lstrip("<").strip()
                    return float(number_part) if number_part else 1e-06
                return float(value)
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    def create_summary_sheet(self, workbook):
        """Create summary sheet with project comparison"""
        try:
            summary_ws = workbook.add_worksheet("Summary")
            
            # Create formats.
            scientific_format = workbook.add_format({
                "num_format": '[<=0.000001]"<1E-06";0.000000',
                "align": "right"
            })
            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#DFF0D8",
                "border": 1,
                "align": "center",
                "valign": "vcenter"
            })
            right_align = workbook.add_format({"align": "right"})
            bold_format = workbook.add_format({"align": "right", "bold": True})
            
            # Determine project order.
            if self.project_order:
                project_order = self.project_order
            else:
                project_order = list(self.all_data.keys())
            
            # Build the baseline map from Step 3 files.
            step3_map = self.collect_step3_areic_map()
            
            # --- Set up the header rows.
            # Always include the first 3 fixed columns.
            summary_ws.merge_range("A1:A2", "Compound", header_format)
            summary_ws.merge_range("B1:B2", "Scenario", header_format)
            summary_ws.merge_range("C1:C2", "Waterbody", header_format)
            
            # For each project, determine how many columns to create.
            col_idx = 3
            for project in project_order:
                shortcode = self.project_shortcodes.get(project, "??")
                if self.areic_comparison_enabled:
                    summary_ws.merge_range(0, col_idx, 0, col_idx+2, shortcode, header_format)
                    summary_ws.write(1, col_idx,   "Max PECsw", header_format)
                    summary_ws.write(1, col_idx+1, "Max PECsed", header_format)
                    summary_ws.write(1, col_idx+2, "Areic dep. (mg/m²)", header_format)
                    col_idx += 3
                else:
                    summary_ws.merge_range(0, col_idx, 0, col_idx+1, shortcode, header_format)
                    summary_ws.write(1, col_idx,   "Max PECsw", header_format)
                    summary_ws.write(1, col_idx+1, "Max PECsed", header_format)
                    col_idx += 2
            
            # --- Collect all unique entries.
            all_entries = set()
            for rows in self.all_data.values():
                for r in rows:
                    key = (r["Scenario"], r["Waterbody"], r["Compound"], r["Type"])
                    all_entries.add(key)
            # Sort entries (Parents first, then Metabolites).
            sorted_entries = sorted(
                [e for e in all_entries if e[3] == "Parent"],
                key=lambda x: (x[0], x[1], x[2])  # Sort by scenario, waterbody, then compound for parents
            ) + sorted(
                [e for e in all_entries if e[3] == "Metabolite"],
                key=lambda x: (x[2], x[0], x[1])  # Sort by compound, then scenario, then waterbody for metabolites
            )
            
            # --- Write data rows.
            row_idx = 2
            for (scenario, waterbody, compound, ctype) in sorted_entries:
                summary_ws.write(row_idx, 0, compound)
                summary_ws.write(row_idx, 1, scenario)
                summary_ws.write(row_idx, 2, waterbody)
                
                col = 3
                for project in project_order:
                    sw_val = 0.0
                    sed_val = 0.0
                    areic_val = None
                    # Find the matching row for this project.
                    match = next(
                        (r for r in self.all_data[project]
                        if r["Scenario"] == scenario and
                            r["Waterbody"] == waterbody and
                            r["Compound"] == compound and
                            r["Type"] == ctype),
                        None
                    )
                    if match:
                        sw_val = self._convert_to_float(match.get("Max PECsw", 0))
                        sed_val = self._convert_to_float(match.get("Max PECsed", 0))
                        try:
                            areic_val = float(match.get("Areic mean deposition", "0"))
                        except Exception:
                            areic_val = None
                    
                    if self.areic_comparison_enabled:
                        # Write PECsw and PECsed.
                        if sw_val <= 1E-6:
                            summary_ws.write_number(row_idx, col, 0, scientific_format)
                        else:
                            summary_ws.write_number(row_idx, col, sw_val, right_align)
                        if sed_val <= 1E-6:
                            summary_ws.write_number(row_idx, col+1, 0, scientific_format)
                        else:
                            summary_ws.write_number(row_idx, col+1, sed_val, right_align)
                        # Write Areic deposition.
                        if areic_val is None or areic_val <= 1E-6:
                            summary_ws.write(row_idx, col+2, "<1E-06", scientific_format)
                        else:
                            summary_ws.write_number(row_idx, col+2, areic_val, right_align)
                        
                        # --- Ratio calculation.
                        # Only perform if areic_val is valid.
                        if areic_val and areic_val > 0:
                            key_for_step3 = (scenario, waterbody, compound, ctype)
                            step3_val = step3_map.get(key_for_step3, None)
                            if step3_val and step3_val > 0:
                                # Ratio = (current_areic / base_areic) * 100.
                                ratio = (areic_val / step3_val) * 100
                                # If ratio is below 5% and both PEC values are nonzero, bold them.
                                if ratio < 5 and sw_val > 1E-6 and sed_val > 1E-6:
                                    summary_ws.write_number(row_idx, col, sw_val, bold_format)
                                    summary_ws.write_number(row_idx, col+1, sed_val, bold_format)
                        col += 3
                    else:
                        # When areic comparison is not enabled, only write PECsw and PECsed.
                        if sw_val <= 1E-6:
                            summary_ws.write_number(row_idx, col, 0, scientific_format)
                        else:
                            summary_ws.write_number(row_idx, col, sw_val, right_align)
                        if sed_val <= 1E-6:
                            summary_ws.write_number(row_idx, col+1, 0, scientific_format)
                        else:
                            summary_ws.write_number(row_idx, col+1, sed_val, right_align)
                        col += 2
                row_idx += 1
            
            # Set fixed widths for the first three columns.
            summary_ws.set_column(0, 2, 20)
            col_offset = 3
            if self.areic_comparison_enabled:
                for _ in project_order:
                    summary_ws.set_column(col_offset, col_offset, 14)      # Max PECsw column
                    summary_ws.set_column(col_offset+1, col_offset+1, 14)    # Max PECsed column
                    summary_ws.set_column(col_offset+2, col_offset+2, 17)    # Areic dep. column
                    col_offset += 3
            else:
                for _ in project_order:
                    summary_ws.set_column(col_offset, col_offset, 14)      # Max PECsw column
                    summary_ws.set_column(col_offset+1, col_offset+1, 14)    # Max PECsed column
                    col_offset += 2
            
        except Exception as e:
            print(f"Summary Error: {str(e)}")

    def extract_daily_value(self, content, label, version, start_index=0):
        """Extract daily value from content"""
        pos = content.find(label, start_index)
        if pos == -1:
            return "N/A"
        if version == 3:
            offset = 13
            length = 22
        else:
            offset = 18
            length = 18
        start = pos + offset
        val_str = content[start : start + length].strip()
        if val_str.startswith("<"):
            try:
                return "< " + str(float(val_str.lstrip("<").strip()))
            except:
                return "< 1E-6"
        try:
            numeric_val = float(val_str)
            if numeric_val < 1E-6:
                return "<1E-06"
            return numeric_val
        except:
            return "N/A"

    def extract_date_only(self, date_str):
        """Extract date only from date string"""
        m = re.match(r"(\d{2}-[A-Za-z]{3}-\d{4})", date_str)
        return m.group(1) if m else date_str

    def format_for_excel(self, value):
        """Format value for Excel export"""
        try:
            num = float(value)
            if num <= 1E-6:
                return "<1E-06"
            return num
        except Exception:
            return value

    def export_to_excel(self, filepath):
        """Export data to Excel with identical formatting to original"""
        if not self.all_data:
            return False
            
        try:
            workbook = xlsxwriter.Workbook(filepath)
            
            # Create summary sheet if summary mode is enabled
            print(f"Summary mode enabled: {self.summary_mode}")
            if self.summary_mode:
                print("Creating summary sheet...")
                self.create_summary_sheet(workbook)
                print("Summary sheet created successfully")
            
            # Create formats
            right_align = workbook.add_format({"align": "right"})
            header_format = workbook.add_format({
                "bg_color": "#82C940",
                "font_color": "#000000",
                "bold": True,
                "align": "left",
            })

            # Daily search arrays (identical to original)
            sw_daily_search = [
                "PECsw_1_day",
                "PECsw_2 days",
                "PECsw_3_days",
                "PECsw_4_days",
                "PECsw_7_days",
                "PECsw_14_days",
                "PECsw_21_days",
                "PECsw_28_days",
                "PECsw_42_days",
                "PECsw_50_days",
                "PECsw_100_days",
            ]
            twaec_sw_search = [
                "TWAECsw_1_day",
                "TWAECsw_2_days",
                "TWAECsw_3_days",
                "TWAECsw_4_days",
                "TWAECsw_7_days",
                "TWAECsw_14_days",
                "TWAECsw_21_days",
                "TWAECsw_28_days",
                "TWAECsw_42_days",
                "TWAECsw_50_days",
                "TWAECsw_100_days",
            ]
            sed_daily_search = [
                "PECsed_1_day",
                "PECsed_2_days",
                "PECsed_3_days",
                "PECsed_4_days",
                "PECsed_7_days",
                "PECsed_14_days",
                "PECsed_21_days",
                "PECsed_28_days",
                "PECsed_42_days",
                "PECsed_50_days",
                "PECsed_100_days",
            ]
            twaec_sed_headers = [
                "TWAECsed_1_day",
                "TWAECsed_2_days",
                "TWAECsed_3_days",
                "TWAECsed_4_days",
                "TWAECsed_7_days",
                "TWAECsed_14_days",
                "TWAECsed_21_days",
                "TWAECsed_28_days",
                "TWAECsed_42_days",
                "TWAECsed_50_days",
                "TWAECsed_100_days",
            ]
            sw_daily_headers = [
                "PECsw 1 day",
                "PECsw 2 days",
                "PECsw 3 days",
                "PECsw 4 days",
                "PECsw 7 days",
                "PECsw 14 days",
                "PECsw 21 days",
                "PECsw 28 days",
                "PECsw 42 days",
                "PECsw 50 days",
                "PECsw 100 days",
            ]
            sed_daily_headers = [
                "PECsed 1 day",
                "PECsed 2 days",
                "PECsed 3 days",
                "PECsed 4 days",
                "PECsed 7 days",
                "PECsed 14 days",
                "PECsed 21 days",
                "PECsed 28 days",
                "PECsed 42 days",
                "PECsed 50 days",
                "PECsed 100 days",
            ]

            # Initialize sheet names tracking
            existing_sheet_names = set()

            for project, rows in self.all_data.items():
                # Create safe sheet name
                worksheet = workbook.add_worksheet(self.safe_sheet_name(project, existing_sheet_names))
                
                # Sort rows (identical to original)
                def get_sum_number(filename):
                    m = re.search(r"(\d+)\.sum$", filename)
                    return int(m.group(1)) if m else 999999999
                sorted_rows = sorted(rows, key=lambda r: (
                    r["Compound"].upper(),
                    get_sum_number(r["Filename"])
                ))

                # --- Water Sheet Header & Data Row ---
                sw_header = (
                    [
                        "Filename",
                        "Compound",
                        "Scenario",
                        "Waterbody",
                        "AppDate 1",
                        "AppDate 2",
                        "Max PECsw (μg/L)",
                        "Date of Max PECsw",
                        "Route of entry",
                    ]
                    + sw_daily_headers
                    + twaec_sw_search
                )
                for col, header in enumerate(sw_header):
                    worksheet.write(0, col, header, header_format)

                current_row = 1
                for r in sorted_rows:
                    try:
                        with open(r["FilePath"], "r", encoding="ISO-8859-1") as f:
                            content = f.read()
                    except:
                        continue

                    version = 4
                    if "FOCUS_TOXSWA v3.3.1" in content:
                        version = 3
                    elif "* FOCUS  TOXSWA version" in content:
                        version = 4

                    app_dates = r["ApplicationDates"][:2]
                    app_dates = [self.extract_date_only(d) for d in app_dates]
                    while len(app_dates) < 2:
                        app_dates.append("")

                    if r["Type"] == "Parent":
                        water_start = 0
                    else:
                        m_sw = re.search(
                            r"\* Table:\s*PEC in water layer of substance:\s+" + re.escape(r["Compound"]),
                            content,
                            re.IGNORECASE,
                        )
                        water_start = m_sw.start() if m_sw else 0

                    max_date_full = ""
                    max_match_full = re.search(
                        r"Global max.*?(\d{2}-[A-Za-z]{3}-\d{4}-\d{2}h\d{2})",
                        content,
                        re.IGNORECASE | re.DOTALL,
                    )
                    if max_match_full:
                        max_date_full = max_match_full.group(1).strip()

                    if max_date_full in r["ApplicationDates"]:
                        main_route = "Spray Drift"
                    else:
                        scenario_code = r["Scenario"][:1].upper() if r["Scenario"] else ""
                        if scenario_code == "D":
                            main_route = "Drainage"
                        elif scenario_code == "R":
                            main_route = "Runoff"
                        else:
                            main_route = "Spray Drift"

                    try:
                        pos_global = content.find("Global max", water_start)
                        max_sw_val = float(content[pos_global + 0x11 : pos_global + 0x11 + 0x13].strip())
                    except:
                        max_sw_val = 0.0

                    m_date = re.search(r"(\d{2}-[A-Za-z]{3}-\d{4})", content[pos_global:])
                    max_sw_date = m_date.group(1) if m_date else ""

                    if r["Type"] == "Parent":
                        pecsw_vals = [self.extract_daily_value(content, key, version, 0) for key in sw_daily_search]
                        twaecsw_vals = [self.extract_daily_value(content, key, version, 0) for key in twaec_sw_search]
                    else:
                        pecsw_vals = [self.extract_daily_value(content, key, version, water_start) for key in sw_daily_search]
                        twaecsw_vals = [self.extract_daily_value(content, key, version, water_start) for key in twaec_sw_search]

                    formatted_max_sw = self.format_for_excel(max_sw_val)

                    data_row = (
                        [
                            r["Filename"],
                            r["Compound"],
                            r["Scenario"],
                            r["Waterbody"],
                            app_dates[0],
                            app_dates[1],
                            formatted_max_sw,
                            max_sw_date,
                            main_route,
                        ]
                        + pecsw_vals
                        + twaecsw_vals
                    )

                    for col, cell in enumerate(data_row):
                        if col == 6 or col >= 9:
                            worksheet.write(current_row, col, cell, right_align)
                        else:
                            worksheet.write(current_row, col, cell)
                    current_row += 1

                worksheet.set_column(0, 50, 15)

                # --- Sediment Sheet Header & Data Row ---
                current_row += 1
                sed_header = (
                    [
                        "Filename",
                        "Compound",
                        "Scenario",
                        "Waterbody",
                        "AppDate 1",
                        "AppDate 2",
                        "Max PECsed (μg/L)",
                        "Date of Max PECSed",
                        "Route of entry",
                    ]
                    + sed_daily_headers
                    + twaec_sed_headers
                )

                for col, header in enumerate(sed_header):
                    worksheet.write(current_row, col, header, header_format)
                current_row += 1

                for r in sorted_rows:
                    try:
                        with open(r["FilePath"], "r", encoding="ISO-8859-1") as f:
                            content = f.read()
                    except:
                        continue

                    version = 4
                    if "FOCUS_TOXSWA v3.3.1" in content:
                        version = 3
                    elif "* FOCUS  TOXSWA version" in content:
                        version = 4

                    app_dates = r["ApplicationDates"][:2]
                    app_dates = [self.extract_date_only(d) for d in app_dates]
                    while len(app_dates) < 2:
                        app_dates.append("")

                    if r["Type"] == "Parent":
                        sed_start = 0
                    else:
                        m_sed = re.search(
                            r"\* Table:\s*PEC in sediment of substance:\s+" + re.escape(r["Compound"]),
                            content,
                            re.IGNORECASE,
                        )
                        sed_start = m_sed.start() if m_sed else 0

                    max_date_full = ""
                    max_match_full = re.search(
                        r"Global max.*?(\d{2}-[A-Za-z]{3}-\d{4}-\d{2}h\d{2})",
                        content,
                        re.IGNORECASE | re.DOTALL,
                    )
                    if max_match_full:
                        max_date_full = max_match_full.group(1).strip()

                    if max_date_full in r["ApplicationDates"]:
                        main_route = "Spray Drift"
                    else:
                        scenario_code = r["Scenario"][:1].upper() if r["Scenario"] else ""
                        if scenario_code == "D":
                            main_route = "Drainage"
                        elif scenario_code == "R":
                            main_route = "Runoff"
                        else:
                            main_route = "Spray Drift"

                    try:
                        pos_global_sed = content.rfind("Global max", sed_start)
                        max_sed_val = float(content[pos_global_sed + 0x11 : pos_global_sed + 0x11 + 0x13].strip())
                    except:
                        max_sed_val = 0.0

                    m_date_sed = re.search(r"(\d{2}-[A-Za-z]{3}-\d{4})", content[pos_global_sed:])
                    max_sed_date = m_date_sed.group(1) if m_date_sed else ""

                    if r["Type"] == "Parent":
                        pecsed_vals = [self.extract_daily_value(content, key, version, 0) for key in sed_daily_search]
                        twaecsed_vals = [self.extract_daily_value(content, key, version, 0) for key in twaec_sed_headers]
                    else:
                        pecsed_vals = [self.extract_daily_value(content, key, version, sed_start) for key in sed_daily_search]
                        twaecsed_vals = [self.extract_daily_value(content, key, version, sed_start) for key in twaec_sed_headers]

                    formatted_max_sed = self.format_for_excel(max_sed_val)

                    data_row = (
                        [
                            r["Filename"],
                            r["Compound"],
                            r["Scenario"],
                            r["Waterbody"],
                            app_dates[0],
                            app_dates[1],
                            formatted_max_sed,
                            max_sed_date,
                            main_route,
                        ]
                        + pecsed_vals
                        + twaecsed_vals
                    )

                    for col, cell in enumerate(data_row):
                        if col == 6 or col >= 9:
                            worksheet.write(current_row, col, cell, right_align)
                        else:
                            worksheet.write(current_row, col, cell)
                    current_row += 1

                worksheet.set_column(0, 50, 15)

            workbook.close()
            return True
            
        except Exception as e:
            print(f"Error exporting to Excel: {str(e)}")
            return False
    
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