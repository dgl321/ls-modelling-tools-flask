import os

os.environ["QT_API"] = "pyqt5"
print("QT_API is set to:", os.environ["QT_API"])

import sys
import re
import PyQt5.QtCore
import PyQt5.QtGui
import PyQt5.QtWidgets
import qtawesome as qta

from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTableWidget,
    QTableWidgetItem,
    QLabel,
    QCheckBox,
    QListWidget,
    QComboBox,
    QMessageBox,
    QMenu,
    QHeaderView,
)
from PyQt5.QtGui import QPixmap, QColor, QDesktopServices, QFont, QIcon
from PyQt5.QtCore import Qt, QUrl, QSize

import xlsxwriter
from xlsxwriter.utility import xl_range


class PELMOExtractor(QWidget):
    def __init__(self):
        super().__init__()
        self.logoLabel = QLabel(self)
        self.initUI()
        # For PELMO extraction, main_dir will store the FOCUS folder path.
        self.main_dir = ""
        # This will hold all extracted rows from all projects.
        self.all_rows = []
        # Default parametric limit value is not set until the user selects one.
        self.limit_value = None

        self.dark_stylesheet = """
        QWidget {
            background-color: #121212;
            color: #E0E0E0;
            font-family: 'Segoe UI', sans-serif;
            font-size: 16px;
        }
        QTableWidget {
            background-color: #1E1E1E;
            gridline-color: #2E2E2E;
        }
        QHeaderView::section {
            background-color: #1E1E1E;
            padding: 6px;
            border: 1px solid #2E2E2E;
        }
        QPushButton {
            background-color: #2D2D30;
            border: 1px solid #3E3A3F;
            padding: 6px;
            border-radius: 3px;
        }
        QPushButton:hover {
            background-color: #3A3A3F;
        }
        QLineEdit, QComboBox {
            background-color: #1E1E1E;
            border: 1px solid #3E3A3F;
            padding: 4px;
            border-radius: 3px;
        }
        QListWidget {
            background-color: #1E1E1E;
            border: 1px solid #3E3A3F;
        }
        QCheckBox { padding: 2px; }
        QLabel { color: #E0E0E0; }
        """
        self.light_stylesheet = """
        QWidget {
            background-color: #F0F0F0;
            color: #202020;
            font-family: 'Segoe UI', sans-serif;
            font-size: 16px;
        }
        QTableWidget {
            background-color: #FFFFFF;
            gridline-color: #C0C0C0;
        }
        QHeaderView::section {
            background-color: #E0E0E0;
            padding: 6px;
            border: 1px solid #C0C0C0;
        }
        QPushButton {
            background-color: #E0E0E0;
            border: 1px solid #A0A0A0;
            padding: 6px;
            border-radius: 3px;
        }
        QPushButton:hover { background-color: #D0D0D0; }
        QLineEdit, QComboBox {
            background-color: #FFFFFF;
            border: 1px solid #A0A0A0;
            padding: 4px;
            border-radius: 3px;
        }
        QListWidget { background-color: #FFFFFF; border: 1px solid #A0A0A0; }
        QCheckBox { padding: 2px; }
        QLabel { color: #202020; }
        """
        self.is_dark_mode = True
        QApplication.instance().setStyleSheet(self.dark_stylesheet)

    def initUI(self):
        self.setWindowTitle("PELMO Extractor")
        self.setGeometry(200, 200, 1150, 600)
        layout = QHBoxLayout()

        # Left Side: File/Project selection.
        file_select_layout = QVBoxLayout()
        script_dir = os.path.dirname(os.path.realpath(__file__))
        logo_path = os.path.join(script_dir, "ls.png")
        pixmap = QPixmap(logo_path)
        if pixmap.isNull():
            print("Logo file 'ls.png' not found at", logo_path)
        else:
            pixmap = pixmap.scaledToWidth(200, Qt.SmoothTransformation)
            self.logoLabel.setPixmap(pixmap)
        self.logoLabel.setAlignment(Qt.AlignCenter)
        file_select_layout.addWidget(self.logoLabel)

        self.subheadingLabel = QLabel("PELMO Extractor v1")
        subheading_font = QFont("Lato", 12)
        subheading_font.setWeight(500)
        self.subheadingLabel.setFont(subheading_font)
        self.subheadingLabel.setStyleSheet("font-size: 13pt; font-weight: 500; color: #bfb800;")
        self.subheadingLabel.setAlignment(Qt.AlignCenter)
        file_select_layout.addWidget(self.subheadingLabel)

        # Button to select PELMO directory.
        self.btnSelectDir = QPushButton("Select PELMO Directory")
        self.btnSelectDir.clicked.connect(self.selectDirectory)
        file_select_layout.addWidget(self.btnSelectDir)

        # Label and list to show project folders (from FOCUS).
        self.projectLabel = QLabel("Select Project(s):")
        file_select_layout.addWidget(self.projectLabel)
        self.fileList = QListWidget()
        # Allow multiple selection.
        self.fileList.setSelectionMode(QListWidget.MultiSelection)
        file_select_layout.addWidget(self.fileList)

        # (Batch mode checkbox is no longer needed; it is hidden.)
        self.batchCheckbox = QCheckBox("Batch Mode")
        self.batchCheckbox.setVisible(False)
        file_select_layout.addWidget(self.batchCheckbox)

        # Right Side: Data and controls.
        data_layout = QVBoxLayout()

        # Top controls: Parametric Limit dropdown with light/dark toggle and info buttons.
        top_controls_layout = QHBoxLayout()
        top_controls_layout.addWidget(QLabel("Parametric Limit:"))
        self.paramLimitComboBox = QComboBox()
        self.paramLimitComboBox.addItem("")  # Empty default.
        self.paramLimitComboBox.addItems(["0.1 µg/l", "0.001 µg/l"])
        self.paramLimitComboBox.setCurrentIndex(0)
        self.paramLimitComboBox.currentIndexChanged.connect(self.updateLimitValue)
        self.paramLimitComboBox.view().setMinimumWidth(90)
        self.paramLimitComboBox.setMinimumWidth(90)
        top_controls_layout.addWidget(self.paramLimitComboBox)
        self.themeToggleButton = QPushButton("Light Mode")
        self.themeToggleButton.setIcon(qta.icon("fa5s.sun", color="white"))
        self.themeToggleButton.clicked.connect(self.toggleTheme)
        top_controls_layout.addWidget(self.themeToggleButton)
        self.infoButton = QPushButton()
        self.infoButton.setIcon(qta.icon("fa5s.info-circle", color="white"))
        self.infoButton.setToolTip("Click for usage instructions")
        self.infoButton.clicked.connect(self.showInfoDialog)
        top_controls_layout.addWidget(self.infoButton)
        top_controls_layout.addStretch()
        data_layout.addLayout(top_controls_layout)

        # Bottom controls: Extract Data, Export to Excel, Copy Table, Reset, Open Excel checkbox.
        button_layout = QHBoxLayout()
        self.btnExtract = QPushButton("Extract Data")
        self.btnExtract.clicked.connect(self.extractData)
        button_layout.addWidget(self.btnExtract)

        self.btnExport = QPushButton("Export to Excel")
        self.btnExport.clicked.connect(self.exportToExcel)
        button_layout.addWidget(self.btnExport)

        self.btnCopy = QPushButton()
        self.btnCopy.setIcon(qta.icon("fa5.copy", color="white"))
        self.btnCopy.setToolTip("Copy table to clipboard")
        self.btnCopy.setFixedSize(30, 30)
        self.btnCopy.setIconSize(QSize(16, 16))
        self.btnCopy.setStyleSheet("padding: 0px;")
        self.btnCopy.clicked.connect(self.copyTableToClipboard)
        button_layout.addWidget(self.btnCopy)

        self.btnReset = QPushButton()
        self.btnReset.setIcon(qta.icon("fa5s.undo", color="white"))
        self.btnReset.setToolTip("Clear table")
        self.btnReset.setFixedSize(30, 30)
        self.btnReset.setIconSize(QSize(16, 16))
        self.btnReset.setStyleSheet("padding: 0px;")
        self.btnReset.clicked.connect(self.resetApplication)
        button_layout.addWidget(self.btnReset)

        self.chkOpenExcel = QCheckBox("Open Excel after export")
        button_layout.addWidget(self.chkOpenExcel)

        data_layout.addLayout(button_layout)

        # QTable to display extracted data.
        self.tableWidget = QTableWidget()
        self.tableWidget.setColumnCount(0)
        self.tableWidget.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # Enable right-click context menu for copying.
        self.tableWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableWidget.customContextMenuRequested.connect(self.showTableContextMenu)
        data_layout.addWidget(self.tableWidget)

        layout.addLayout(file_select_layout, 1)
        layout.addLayout(data_layout, 5)
        self.setLayout(layout)

    def selectDirectory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select PELMO Directory")
        if directory:
            # Look for the FOCUS folder inside the selected PELMO directory.
            focus_folder = None
            for item in os.listdir(directory):
                if os.path.isdir(os.path.join(directory, item)) and item.upper() == "FOCUS":
                    focus_folder = os.path.join(directory, item)
                    break
            if not focus_folder:
                QMessageBox.critical(self, "Error", "FOCUS folder not found in the selected PELMO directory.")
                return
            self.main_dir = focus_folder
            self.updateFileList()

    def updateFileList(self):
        self.fileList.clear()
        if not self.main_dir:
            return
        # List all project folders (directories ending with .run) in the FOCUS folder.
        project_folders = [
            f for f in os.listdir(self.main_dir)
            if os.path.isdir(os.path.join(self.main_dir, f)) and f.endswith(".run")
        ]
        if not project_folders:
            self.fileList.addItem("No project folders found.")
        else:
            for proj in sorted(project_folders):
                self.fileList.addItem(proj)

    def extractData(self):
        try:
            if not self.main_dir:
                QMessageBox.warning(self, "Error", "Please select a PELMO directory first!")
                return

            selected_items = self.fileList.selectedItems()
            if not selected_items:
                QMessageBox.warning(self, "Error", "Please select one or more project folders for extraction!")
                return

            all_rows = []
            all_extra_keys = set()  # To collect extra column names (active substance and metabolites)
            for item in selected_items:
                project_folder_name = item.text()
                project_path = os.path.join(self.main_dir, project_folder_name)
                # Get all crop folders (directories ending with .run) in the project folder.
                crop_folders = [
                    d for d in os.listdir(project_path)
                    if os.path.isdir(os.path.join(project_path, d)) and d.endswith(".run")
                ]
                if not crop_folders:
                    QMessageBox.critical(
                        self,
                        "Extraction Error",
                        f"No crop folders found in project '{project_folder_name}'.\nEnsure that a FOCUS Summary Report has been generated in the PELMO Evaluation window.",
                    )
                    continue

                for crop_folder in crop_folders:
                    crop_folder_path = os.path.join(project_path, crop_folder)
                    # Look for scenario folders within the crop folder (pattern: {Scenario}_-_{*}.run).
                    scenario_folders = [
                        d for d in os.listdir(crop_folder_path)
                        if os.path.isdir(os.path.join(crop_folder_path, d)) and "_-_" in d and d.endswith(".run")
                    ]
                    if not scenario_folders:
                        QMessageBox.critical(
                            self,
                            "Extraction Error",
                            f"No scenario folders found in crop folder '{crop_folder}'.\nEnsure that a FOCUS Summary Report has been generated in the PELMO Evaluation window.",
                        )
                        continue

                    scenario_found = False
                    for scenario_folder in scenario_folders:
                        scenario_folder_path = os.path.join(crop_folder_path, scenario_folder)
                        period_plm_path = os.path.join(scenario_folder_path, "period.plm")
                        if not os.path.exists(period_plm_path):
                            QMessageBox.critical(
                                self,
                                "Extraction Error",
                                f"'period.plm' not found in scenario folder '{scenario_folder}'.\nPlease generate a FOCUS Summary Report in the PELMO Evaluation window.",
                            )
                            continue
                        scenario_found = True
                        active_substance, active_pec_value, metabolites = self.extract_active_substance_and_metabolites(period_plm_path)
                        if not active_substance or not active_pec_value:
                            continue
                        row = {}
                        # Include the project name in the row.
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
                        all_rows.append(row)
                    if not scenario_found:
                        QMessageBox.critical(
                            self,
                            "Extraction Error",
                            f"No valid scenario folder with 'period.plm' found in crop folder '{crop_folder}'.",
                        )

            if not all_rows:
                QMessageBox.critical(
                    self,
                    "Extraction Error",
                    "No valid period.plm files were found in any crop/scenario folder.",
                )
                return

            # Build the table header: fixed columns plus extra keys.
            header = ["Project", "Crop", "Scenario"] + sorted(all_extra_keys)
            for row in all_rows:
                for key in header:
                    if key not in row:
                        row[key] = ""
            # Update the table widget.
            self.tableWidget.setColumnCount(len(header))
            self.tableWidget.setHorizontalHeaderLabels(header)
            self.tableWidget.setRowCount(len(all_rows))
            for row_idx, row in enumerate(all_rows):
                for col_idx, key in enumerate(header):
                    item = QTableWidgetItem(str(row[key]))
                    self.tableWidget.setItem(row_idx, col_idx, item)

            # Store the extracted rows for use in Excel export.
            self.all_rows = all_rows
            self.applyTableConditionalFormatting()

        except Exception as e:
            QMessageBox.critical(self, "Extraction Error", f"Failed to extract data: {str(e)}")

    def applyTableConditionalFormatting(self):
        if self.limit_value is not None:
            for row in range(self.tableWidget.rowCount()):
                # Assume numeric columns start after the first three fixed columns.
                for col in range(3, self.tableWidget.columnCount()):
                    item = self.tableWidget.item(row, col)
                    if item:
                        try:
                            val = float(item.text())
                            if val >= self.limit_value:
                                item.setForeground(QColor("red"))
                            else:
                                item.setForeground(QColor("white") if self.is_dark_mode else QColor("black"))
                        except ValueError:
                            pass

    def extract_active_substance_and_metabolites(self, file_path):
        active_substance = None
        active_pec_value = None
        metabolites = []
        metabolite = None

        with open(file_path, "r", encoding="ISO-8859-1") as file:
            for line in file:
                if "Results for ACTIVE SUBSTANCE" in line and "percolate at 1 m soil depth" in line:
                    match = re.search(r"Results for ACTIVE SUBSTANCE \((.*?)\)", line)
                    if match:
                        active_substance = match.group(1)
                if active_substance and "80 Perc." in line:
                    active_pec_value = line.split()[-1]
                if "Results for METABOLITE" in line and "percolate at 1 m soil depth" in line:
                    match = re.search(r"Results for METABOLITE.*?\((.*?)\)", line)
                    if match:
                        metabolite = match.group(1)
                if metabolite and "80 Perc." in line:
                    pec_value = line.split()[-1]
                    metabolites.append((metabolite, pec_value))
                    metabolite = None

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

    def updateLimitValue(self):
        current = self.paramLimitComboBox.currentText().strip()
        if current == "":
            self.limit_value = None
            print("No parametric limit selected.")
        elif current.startswith("0.1"):
            self.limit_value = 0.1
            print("Parametric limit set to:", self.limit_value)
        elif current.startswith("0.001"):
            self.limit_value = 0.001
            print("Parametric limit set to:", self.limit_value)
        self.applyTableConditionalFormatting()

    def toggleTheme(self):
        app = QApplication.instance()
        if self.is_dark_mode:
            app.setStyleSheet(self.light_stylesheet)
            self.themeToggleButton.setText("Dark Mode")
            self.themeToggleButton.setIcon(qta.icon("fa5s.moon", color="black"))
            self.infoButton.setIcon(qta.icon("fa5s.info-circle", color="black"))
            self.btnReset.setIcon(qta.icon("fa5s.undo", color="black"))
            self.btnCopy.setIcon(qta.icon("fa5.copy", color="black"))
        else:
            app.setStyleSheet(self.dark_stylesheet)
            self.themeToggleButton.setText("Light Mode")
            self.themeToggleButton.setIcon(qta.icon("fa5s.sun", color="white"))
            self.infoButton.setIcon(qta.icon("fa5s.info-circle", color="white"))
            self.btnReset.setIcon(qta.icon("fa5s.undo", color="white"))
            self.btnCopy.setIcon(qta.icon("fa5.copy", color="white"))
        self.is_dark_mode = not self.is_dark_mode
        self.applyTableConditionalFormatting()

    def showInfoDialog(self):
        info_text = (
            "<html>"
            "<body>"
            "<b>Instructions:</b><br><br>"
            "1. Select a PELMO directory containing the FOCUS folder.<br>"
            "2. The FOCUS folder will be scanned for project folders (ending with .run).<br>"
            "3. Multi-select one or more project folders and click 'Extract Data'.<br>"
            "4. Within each selected project, the scenario folder for a given crop should contain a 'period.plm' file.<br>"
            "5. If a scenario folder does not contain 'period.plm', you will be prompted to generate a 'FOCUS Summary Report' in the PELMO Evaluation window.<br>"
            "6. Extracted data is displayed in the table with columns for Project, Crop, Scenario and PEC values.<br>"
            "7. Use the parametric limit dropdown to highlight exceedances.<br>"
            "8. Right-click the table to copy selected data or click the copy icon to copy the entire table.<br>"
            "9. Click 'Export to Excel' to save the data; if 'Open Excel after export' is checked, the file will open automatically.<br>"
            "10. Click the clear icon to reset the table.<br>"
            "</body>"
            "</html>"
        )
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle("Usage Instructions")
        msgBox.setText(info_text)
        msgBox.setStyleSheet("QLabel { font-size: 12pt; }")
        msgBox.exec_()

    def exportToExcel(self):
        filePath, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
        if not filePath:
            return

        workbook = xlsxwriter.Workbook(filePath)
        # Group rows by project.
        projects = {}
        for row in self.all_rows:
            proj = row["Project"]
            projects.setdefault(proj, []).append(row)

        used_sheet_names = set()
        for project, rows in projects.items():
            # Create a sheet name from the project folder name.
            sheet_name = project[:-4] if project.lower().endswith(".run") else project
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            orig_sheet_name = sheet_name
            count = 1
            while sheet_name in used_sheet_names:
                sheet_name = f"{orig_sheet_name}_{count}"
                count += 1
            used_sheet_names.add(sheet_name)

            worksheet = workbook.add_worksheet(sheet_name)
            # Re-calculate header for this project.
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

            # Adjust column widths.
            for col in range(len(header)):
                max_width = len(header[col])
                for row in rows:
                    cell_text = str(row.get(header[col], ""))
                    max_width = max(max_width, len(cell_text))
                worksheet.set_column(col, col, max_width + 2)

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
        QMessageBox.information(self, "Export Successful", f"Data exported to {filePath}")
        if self.chkOpenExcel.isChecked():
            QDesktopServices.openUrl(QUrl.fromLocalFile(filePath))

    def resetApplication(self):
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(0)  # Clear column headings
        self.all_rows.clear()

    def copyTableToClipboard(self):
        clipboard_text = ""
        colCount = self.tableWidget.columnCount()
        rowCount = self.tableWidget.rowCount()
        headers = []
        for col in range(colCount):
            header_item = self.tableWidget.horizontalHeaderItem(col)
            header_text = header_item.text() if header_item else ""
            headers.append(header_text)
        clipboard_text += "\t".join(headers) + "\n"
        for row in range(rowCount):
            row_data = []
            for col in range(colCount):
                item = self.tableWidget.item(row, col)
                row_data.append(item.text() if item else "")
            clipboard_text += "\t".join(row_data) + "\n"
        QApplication.clipboard().setText(clipboard_text)
        QMessageBox.information(self, "Copied", "Table data copied to clipboard.")

    def showTableContextMenu(self, pos):
        menu = QMenu(self)
        copyAction = menu.addAction("Copy")
        action = menu.exec_(self.tableWidget.viewport().mapToGlobal(pos))
        if action == copyAction:
            self.copyTableToClipboard()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    script_dir = os.path.dirname(os.path.realpath(__file__))
    icon_path = os.path.join(script_dir, "logo.ico")
    app.setWindowIcon(QIcon(icon_path))
    extractor = PELMOExtractor()
    extractor.show()
    sys.exit(app.exec_())
