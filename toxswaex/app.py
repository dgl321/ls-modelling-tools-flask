import os
import re
import json
import base64
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
import xlsxwriter
from xlsxwriter.utility import xl_range

app = Flask(__name__)
app.secret_key = 'toxswa_extractor_secret_key'

# Global variables to store session data
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

class TOXSWAExtractor:
    def __init__(self):
        self.main_dir = ""
        self.all_data = {}
        self.project_shortcodes = {}
        self.rac_value = None
        self.areic_comparison_enabled = False

    def extract_shortcode(self, folder_path):
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

        spray_section = re.search(
            r"Spray drift mitigation.*?(?=Run-off mitigation)",
            content,
            re.DOTALL | re.IGNORECASE,
        )
        spray_content = spray_section.group(0) if spray_section else content

        m = re.search(
            r"Buffer\s*width\s*\(m\)\s*:\s*(\d+)", spray_content, re.IGNORECASE
        )
        if m:
            buffer = f"{m.group(1)}b"

        m = re.search(
            r"Nozzle\s*reduction\s*\(\%\)\s*:\s*(\d+)", spray_content, re.IGNORECASE
        )
        if m:
            nozzle_val = int(m.group(1))
            if nozzle_val > 0:
                nozzle = f"{nozzle_val}%"

        runoff_section = re.search(
            r"Run-off mitigation.*?(?=Dry deposition)",
            content,
            re.DOTALL | re.IGNORECASE,
        )
        if runoff_section:
            runoff_content = runoff_section.group(0)
            if re.search(
                r"Reduction\s*run-?off\s*mode:\s*VfsMod", runoff_content, re.IGNORECASE
            ):
                m = re.search(
                    r"Filter\s*strip\s*buffer\s*width\s*:\s*(\d+)",
                    runoff_content,
                    re.IGNORECASE,
                )
                if m:
                    vfs = f"{m.group(1)}vfs"
                vfs_flag = " VFSMOD"

            elif re.search(
                r"Reduction\s*run-?off\s*mode:\s*ManualReduction",
                runoff_content,
                re.IGNORECASE,
            ):
                fr_volume_match = re.search(
                    r"Fractional\s+reduction\s+in\s+run-off\s+volume\s*:\s*([\d.]+)",
                    runoff_content,
                    re.IGNORECASE,
                )
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
        m = re.search(
            r"Areic mean deposition\s*\(mg\.m-2\).*?\n\s*\d+\s+[^\n]*\s+([\d\.Ee-]+)",
            content,
            re.IGNORECASE,
        )
        return m.group(1).strip() if m else "N/A"

    def extract_daily_value(self, content, label, version, start_index=0):
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
        m = re.match(r"(\d{2}-[A-Za-z]{3}-\d{4})", date_str)
        return m.group(1) if m else date_str

    def parse_value(self, value_str):
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
        PARENT_DECIMALS_SMALL = 4  # For values < 1
        PARENT_DECIMALS_LARGE = 2  # For values >= 1
        METABOLITE_DECIMALS_SMALL = 6  # For values < 1
        METABOLITE_DECIMALS_LARGE = 2  # For values >= 1

        try:
            num = float(val)
            if num <= 1E-6:
                return "<1E-06"
            
            # Choose decimals based on value magnitude and compound type
            if compound_type == "Parent":
                decimals = PARENT_DECIMALS_SMALL if num < 1 else PARENT_DECIMALS_LARGE
            else:
                decimals = METABOLITE_DECIMALS_SMALL if num < 1 else METABOLITE_DECIMALS_LARGE
                
            return f"{num:.{decimals}f}"
        except:
            return str(val)

    def process_files(self, folder_path, project_name, selected_files=None):
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
                    waterbody = (
                        wb_match.group(1).capitalize() if wb_match else "Unknown"
                    )

            app_dates = []
            app_section = re.search(
                r"Appl\.No\s+Date/Hour.*?\n(.*?)\n\n", content, re.DOTALL
            )
            if app_section:
                for line in app_section.group(1).split("\n"):
                    date_match = re.search(r"\d{2}-[A-Za-z]{3}-\d{4}-\d{2}h\d{2}", line)
                    if date_match:
                        app_dates.append(date_match.group().strip())

            max_date = ""
            max_match = re.search(
                r"Global max.*?(\d{2}-[A-Za-z]{3}-\d{4}-\d{2}h\d{2})",
                content,
                re.IGNORECASE | re.DOTALL,
            )
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

            parent_compound = self.extract_value(
                content, r"\* Substance\s*:\s*(\S+)", "Unknown"
            )
            parent_max_sw = self.extract_float(content, r"Global max.*?([\d.]+)")
            pecsed_match = re.search(
                r"PEC in sediment of substance:\s*\S+.*?Global max\s+([<]?\s*\d+(?:\.\d+)?(?:e[+-]?\d+)?)",
                content,
                re.DOTALL,
            )
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
                max_sw_str = self.extract_value(
                    content,
                    rf"\* Table: PEC in water layer of substance:\s+{re.escape(sub)}.*?Global max\s+([<]?\s*\S+)",
                    "0",
                )
                max_sed_str = self.extract_value(
                    content,
                    rf"\* Table: PEC in sediment of substance:\s+{re.escape(sub)}.*?Global max\s+([<]?\s*\S+)",
                    "0",
                )
                row_met = {
                    "Filename": filename,
                    "Scenario": scenario,
                    "Waterbody": waterbody,
                    "Compound": sub,
                    "Max PECsw": self.format_for_display(self.parse_value(max_sw_str)),
                    "Max PECsed": self.format_for_display(
                        self.parse_value(max_sed_str)
                    ),
                    "Route": route,
                    "Type": "Metabolite",
                    "ApplicationDates": app_dates,
                    "FilePath": file_path,
                }
                all_rows.append(row_met)

        if all_rows:
            self.all_data[project_name] = all_rows

    def extract_value(self, content, pattern, default_value):
        m = re.search(pattern, content, re.DOTALL)
        return m.group(1).strip() if m else default_value

    def extract_float(self, content, pattern):
        m = re.search(pattern, content)
        if not m:
            return None
        val_str = m.group(1).strip()
        return self.parse_value(val_str)

    def extract_data(self, main_dir, selected_projects, selected_files=None, rac_value=None, areic_comparison=False):
        """Extract data from TOXSWA directories"""
        self.main_dir = main_dir
        self.rac_value = rac_value
        self.areic_comparison_enabled = areic_comparison
        self.all_data.clear()
        self.project_shortcodes.clear()
        errors = []

        for project in selected_projects:
            project_path = os.path.join(self.main_dir, project, "toxswa")
            if os.path.exists(project_path):
                self.process_files(project_path, project, selected_files)
            else:
                errors.append(f"TOXSWA folder not found in project '{project}'")

        return self.all_data, errors

    def get_table_data(self, compound_type="Parent", sort_by="Filename"):
        """Get formatted table data for display"""
        if not self.all_data:
            return [], []

        filtered_rows = []
        for project, rows in self.all_data.items():
            for row in rows:
                if row.get("Type", "") == compound_type:
                    filtered_rows.append((project, row))

        if not filtered_rows:
            return [], []

        # Sort rows
        def get_sum_number(filename):
            m = re.search(r"(\d+)\.sum$", filename)
            return int(m.group(1)) if m else 999999999

        if sort_by == "File number":
            sort_key = lambda pr: get_sum_number(pr[1]["Filename"])
        elif sort_by == "Compound":
            sort_key = lambda pr: pr[1]["Compound"].upper()
        elif sort_by == "Scenario":
            sort_key = lambda pr: pr[1]["Scenario"].upper()
        else:
            sort_key = lambda pr: (pr[0].upper(), pr[1]["Filename"].upper())

        filtered_rows = sorted(filtered_rows, key=sort_key)

        # Build headers
        if self.areic_comparison_enabled:
            headers = [
                "Project", "Filename", "Compound", "Scenario", "Waterbody",
                "Max PECsw", "Max PECsed", "Areic dep.", "Route of entry"
            ]
        else:
            headers = [
                "Project", "Filename", "Compound", "Scenario", "Waterbody",
                "Max PECsw", "Max PECsed", "Route of entry"
            ]

        # Build table data
        table_data = []
        for project, row in filtered_rows:
            data_row = {
                "Project": project,
                "Filename": row["Filename"],
                "Compound": row["Compound"],
                "Scenario": row["Scenario"],
                "Waterbody": row["Waterbody"],
                "Max PECsw": row.get("Max PECsw", "0"),
                "Max PECsed": row.get("Max PECsed", "0"),
            }
            if self.areic_comparison_enabled:
                data_row["Areic dep."] = row.get("Areic mean deposition", "N/A")
            data_row["Route of entry"] = row["Route"]
            table_data.append(data_row)

        return table_data, headers

    def export_to_excel(self, filepath):
        """Export data to Excel file"""
        workbook = xlsxwriter.Workbook(filepath)
        
        # Create formats
        right_align = workbook.add_format({"align": "right"})
        header_format = workbook.add_format({
            "bg_color": "#82C940",
            "font_color": "#000000",
            "bold": True,
            "align": "left",
        })

        # Define search patterns for daily values
        sw_daily_search = [
            "PECsw_1_day", "PECsw_2 days", "PECsw_3_days", "PECsw_4_days",
            "PECsw_7_days", "PECsw_14_days", "PECsw_21_days", "PECsw_28_days",
            "PECsw_42_days", "PECsw_50_days", "PECsw_100_days",
        ]
        twaec_sw_search = [
            "TWAECsw_1_day", "TWAECsw_2_days", "TWAECsw_3_days", "TWAECsw_4_days",
            "TWAECsw_7_days", "TWAECsw_14_days", "TWAECsw_21_days", "TWAECsw_28_days",
            "TWAECsw_42_days", "TWAECsw_50_days", "TWAECsw_100_days",
        ]
        sed_daily_search = [
            "PECsed_1_day", "PECsed_2_days", "PECsed_3_days", "PECsed_4_days",
            "PECsed_7_days", "PECsed_14_days", "PECsed_21_days", "PECsed_28_days",
            "PECsed_42_days", "PECsed_50_days", "PECsed_100_days",
        ]
        twaec_sed_headers = [
            "TWAECsed_1_day", "TWAECsed_2_days", "TWAECsed_3_days", "TWAECsed_4_days",
            "TWAECsed_7_days", "TWAECsed_14_days", "TWAECsed_21_days", "TWAECsed_28_days",
            "TWAECsed_42_days", "TWAECsed_50_days", "TWAECsed_100_days",
        ]
        sw_daily_headers = [
            "PECsw 1 day", "PECsw 2 days", "PECsw 3 days", "PECsw 4 days",
            "PECsw 7 days", "PECsw 14 days", "PECsw 21 days", "PECsw 28 days",
            "PECsw 42 days", "PECsw 50 days", "PECsw 100 days",
        ]
        sed_daily_headers = [
            "PECsed 1 day", "PECsed 2 days", "PECsed 3 days", "PECsed 4 days",
            "PECsed 7 days", "PECsed 14 days", "PECsed 21 days", "PECsed 28 days",
            "PECsed 42 days", "PECsed 50 days", "PECsed 100 days",
        ]

        # Initialize a set to track existing worksheet names
        existing_sheet_names = set()

        def safe_sheet_name(name, existing_names):
            """Returns a version of the sheet name that is safe for Excel"""
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

        for project, rows in self.all_data.items():
            # Use the updated safe_sheet_name with the set
            worksheet = workbook.add_worksheet(safe_sheet_name(project, existing_sheet_names))
            sorted_rows = sorted(
                rows,
                key=lambda r: (
                    r["Compound"].upper(),
                    (
                        int(match.group(1))
                        if (match := re.search(r"(\d+)\.sum$", r["Filename"])) is not None
                        else 999999999
                    ),
                ),
            )

            # --- Water Sheet Header & Data Row ---
            sw_header = (
                [
                    "Filename", "Compound", "Scenario", "Waterbody", "AppDate 1",
                    "AppDate 2", "Max PECsw (μg/L)", "Date of Max PECsw", "Route of entry",
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
                        r["Filename"], r["Compound"], r["Scenario"], r["Waterbody"],
                        app_dates[0], app_dates[1], formatted_max_sw, max_sw_date, main_route,
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
                    "Filename", "Compound", "Scenario", "Waterbody", "AppDate 1",
                    "AppDate 2", "Max PECsed (μg/L)", "Date of Max PECSed", "Route of entry",
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
                        r["Filename"], r["Compound"], r["Scenario"], r["Waterbody"],
                        app_dates[0], app_dates[1], formatted_max_sed, max_sed_date, main_route,
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

            # Apply conditional formatting for RAC exceedance
            if self.rac_value is not None:
                sw_data_start = 1
                sw_data_end = sw_data_start + len(sorted_rows) - 1
                worksheet.conditional_format(
                    sw_data_start, 6, sw_data_end, 6,
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": self.rac_value,
                        "format": workbook.add_format({"font_color": "red"}),
                    },
                )
                sed_data_start = len(sorted_rows) + 3
                sed_data_end = sed_data_start + len(sorted_rows) - 1
                worksheet.conditional_format(
                    sed_data_start, 6, sed_data_end, 6,
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": self.rac_value,
                        "format": workbook.add_format({"font_color": "red"}),
                    },
                )

        workbook.close()

    def format_for_excel(self, value):
        """Returns "<1E-06" if the value is less than or equal to 0.000001; otherwise, returns the value as a float."""
        try:
            num = float(value)
            if num <= 1E-6:
                return "<1E-06"
            return num
        except Exception:
            return value

# Global extractor instance
extractor = TOXSWAExtractor()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/scan_directory', methods=['POST'])
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

@app.route('/extract_data', methods=['POST'])
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

@app.route('/export_excel', methods=['POST'])
def export_excel():
    try:
        if not extractor.all_data:
            return jsonify({'error': 'No data to export'})
        
        # Create temporary file
        import tempfile
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

@app.route('/get_table_data')
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

if __name__ == '__main__':
    app.run(debug=True, port=5001) 