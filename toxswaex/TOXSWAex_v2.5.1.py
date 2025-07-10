import os

os.environ["QT_API"] = "pyqt5"

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
    QLineEdit,
    QMessageBox,
    QMenu,
    QHeaderView,
    QDialog,
)
from PyQt5.QtGui import QPixmap, QColor, QDesktopServices, QFont, QIcon
from PyQt5.QtCore import Qt, QUrl, QSize

import xlsxwriter

class SettingsDialog(QDialog):
    def __init__(self, current_value, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setMinimumWidth(300)
        
        self.checkbox = QCheckBox("Highlight mitigation exceeding 95%")
        self.checkbox.setChecked(current_value)
        
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Cancel")
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        
        layout = QVBoxLayout()
        layout.addWidget(self.checkbox)
        layout.addLayout(button_layout)
        self.setLayout(layout)
        
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)

def safe_sheet_name(name, existing_names):
    """
    Returns a version of the sheet name that is safe for Excel:
    - Removes invalid characters: []:*?/\
    - Trims leading/trailing spaces.
    - Truncates to 31 characters if needed.
    - Appends a numeric suffix if necessary to ensure uniqueness.
    """
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

def extract_date_only(date_str):
    m = re.match(r"(\d{2}-[A-Za-z]{3}-\d{4})", date_str)
    return m.group(1) if m else date_str


def extract_areic_mean_deposition(content):
    m = re.search(
        r"Areic mean deposition\s*\(mg\.m-2\).*?\n\s*\d+\s+[^\n]*\s+([\d\.Ee-]+)",
        content,
        re.IGNORECASE,
    )
    return m.group(1).strip() if m else "N/A"


def extract_daily_value(content, label, version, start_index=0):
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


class TOXSWAExtractor(QWidget):
    def __init__(self):
        super().__init__()
        self.logoLabel = QLabel(self)
        self.initUI()
        self.main_dir = "C:/SwashProjects"
        self.subfolder = ""
        self.all_data = {}
        self.project_shortcodes = {}
        self.batch_mode = False
        self.summary_mode = False

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

        if os.path.exists(self.main_dir):
            self.updateSubfolderList()
       
        self.areic_comparison_enabled = False  # Default value.
    

    def initUI(self):
        self.setWindowTitle("TOXSWA Extractor")
        self.setGeometry(200, 200, 1300, 600)

        # Create the main vertical layout (menubar at top, rest below)
        main_layout = QVBoxLayout(self)

        # Now create the rest of the UI in an HBoxLayout
        layout = QHBoxLayout()

        # LEFT SIDE: File Selection Panel
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

        self.subheadingLabel = QLabel("TOXSWA Extractor v1")
        subheading_font = QFont("Lato", 12)
        subheading_font.setWeight(500)
        self.subheadingLabel.setFont(subheading_font)
        self.subheadingLabel.setStyleSheet(
            "font-size: 13pt; font-weight: 500; color: #bfb800;"
        )
        self.subheadingLabel.setAlignment(Qt.AlignCenter)
        file_select_layout.addWidget(self.subheadingLabel)

        self.btnSelectDir = QPushButton("Select SWASH Directory")
        self.btnSelectDir.clicked.connect(self.selectDirectory)
        file_select_layout.addWidget(self.btnSelectDir)

        self.projectLabel = QLabel("Select Project:")
        file_select_layout.addWidget(self.projectLabel)
        self.subfolderDropdown = QComboBox()
        self.subfolderDropdown.setMinimumWidth(150)
        self.subfolderDropdown.currentIndexChanged.connect(self.updateFileList)
        file_select_layout.addWidget(self.subfolderDropdown)

        self.fileListLabel = QLabel(".sum Files:")
        file_select_layout.addWidget(self.fileListLabel)
        self.fileList = QListWidget()
        self.fileList.setSelectionMode(QListWidget.MultiSelection)
        file_select_layout.addWidget(self.fileList)

        # Batch mode and summary options
        self.batchCheckbox = QCheckBox("Batch Mode")
        self.batchCheckbox.stateChanged.connect(self.toggleBatchMode)
        file_select_layout.addWidget(self.batchCheckbox)

        self.summaryCheckbox = QCheckBox("Summary Sheet")
        self.summaryCheckbox.setVisible(False)
        file_select_layout.addWidget(self.summaryCheckbox)

        self.projectOrderLabel = QLabel("Order projects for summary:")
        self.projectOrderLabel.setVisible(False)
        file_select_layout.addWidget(self.projectOrderLabel)

        self.projectOrderList = QListWidget()
        self.projectOrderList.setDragDropMode(QListWidget.InternalMove)
        self.projectOrderList.setVisible(False)
        file_select_layout.addWidget(self.projectOrderList)

        self.summaryCheckbox.stateChanged.connect(self.toggleSummaryOrderList)

        # RIGHT SIDE: Data Display Panel
        data_layout = QVBoxLayout()
        top_controls_layout = QHBoxLayout()

        self.pnecInput = QLineEdit()
        self.pnecInput.setPlaceholderText("RAC (μg/L)")
        self.pnecInput.setMaximumWidth(100)
        self.pnecInput.setAlignment(Qt.AlignCenter)
        self.pnecInput.textChanged.connect(self.updateTable)
        top_controls_layout.addWidget(QLabel("RAC:"))
        top_controls_layout.addWidget(self.pnecInput)

        self.compoundTypeDropdown = QComboBox()
        self.compoundTypeDropdown.addItems(["Parent", "Metabolite"])
        self.compoundTypeDropdown.setMinimumWidth(150)
        self.compoundTypeDropdown.currentIndexChanged.connect(self.updateTable)
        top_controls_layout.addWidget(QLabel("Compound Type:"))
        top_controls_layout.addWidget(self.compoundTypeDropdown)

        self.sortDropdown = QComboBox()
        self.sortDropdown.addItems(["Filename", "Compound", "Scenario"])
        self.sortDropdown.setMinimumWidth(150)
        self.sortDropdown.currentIndexChanged.connect(self.updateTable)
        top_controls_layout.addWidget(QLabel("Sort by:"))
        top_controls_layout.addWidget(self.sortDropdown)

        self.themeToggleButton = QPushButton("Light Mode")
        self.themeToggleButton.setIcon(qta.icon("fa5s.sun", color="white"))
        self.themeToggleButton.clicked.connect(self.toggleTheme)
        top_controls_layout.addWidget(self.themeToggleButton)

        self.infoButton = QPushButton()
        self.infoButton.setIcon(qta.icon("fa5s.info-circle", color="white"))
        self.infoButton.setToolTip("Click for usage instructions")
        self.infoButton.clicked.connect(self.showInfoDialog)
        top_controls_layout.addWidget(self.infoButton)
        
        # Create the gear-icon button.
        self.settingsButton = QPushButton()
        self.settingsButton.setIcon(qta.icon("fa5s.cog", color="white"))
        self.settingsButton.setToolTip("Settings")
        self.settingsButton.setFlat(True)  # Optional: makes it look like an icon-only button.
        self.settingsButton.clicked.connect(self.openSettingsDialog)
        top_controls_layout.addWidget(self.settingsButton)

        top_controls_layout.addStretch()
        data_layout.addLayout(top_controls_layout)

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

        self.tableWidget = QTableWidget()
        self.tableWidget.setColumnCount(8)
        self.tableWidget.setHorizontalHeaderLabels(
            [
                "Project",
                "Filename",
                "Compound",
                "Scenario",
                "Waterbody",
                "Max PECsw",
                "Max PECsed",
                "Route of entry",
            ]
        )
        self.tableWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableWidget.customContextMenuRequested.connect(self.showContextMenu)
        self.tableWidget.setColumnWidth(0, 140)
        self.tableWidget.setColumnWidth(7, 135)
        self.tableWidget.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        data_layout.addWidget(self.tableWidget)

        layout.addLayout(file_select_layout, 1)
        layout.addLayout(data_layout, 5)

        main_layout.addLayout(layout)
        self.setLayout(main_layout)

    def toggleAreicComparison(self, checked):
        self.areic_comparison_enabled = checked
        
    def openSettingsDialog(self):
        dialog = SettingsDialog(self.areic_comparison_enabled, self)
        if dialog.exec_() == QDialog.Accepted:
            self.areic_comparison_enabled = dialog.checkbox.isChecked()        

    def resetApplication(self):
        try:
            self.all_data.clear()
            self.tableWidget.setRowCount(0)
            self.fileList.clearSelection()
            self.pnecInput.clear()
            self.projectOrderList.clear()
            self.projectOrderList.setVisible(False)
            self.projectOrderLabel.setVisible(False)
            if self.main_dir:
                self.updateSubfolderList()
                self.updateFileList()
            self.compoundTypeDropdown.setCurrentIndex(0)
            self.sortDropdown.setCurrentIndex(0)
        except Exception as e:
            QMessageBox.warning(self, "Reset Error", f"Error during reset: {str(e)}")

    def toggleSummaryOrderList(self, state):
        is_checked = state == Qt.Checked
        self.projectOrderList.setVisible(is_checked)
        self.projectOrderLabel.setVisible(is_checked)
        if is_checked:
            self.projectOrderList.clear()
            for project in self.all_data.keys():
                self.projectOrderList.addItem(project)

    def toggleTheme(self):
        """Toggle between dark and light mode."""
        app = QApplication.instance()
        if self.is_dark_mode:
            app.setStyleSheet(self.light_stylesheet)
            self.themeToggleButton.setText("Dark Mode")
            self.themeToggleButton.setIcon(qta.icon("fa5s.moon", color="black"))
            self.infoButton.setIcon(qta.icon("fa5s.info-circle", color="black"))
            self.btnReset.setIcon(qta.icon("fa5s.undo", color="black"))
            self.btnCopy.setIcon(qta.icon("fa5.copy", color="black"))
            self.settingsButton.setIcon(qta.icon("fa5s.cog", color="black"))
        else:
            app.setStyleSheet(self.dark_stylesheet)
            self.themeToggleButton.setText("Light Mode")
            self.themeToggleButton.setIcon(qta.icon("fa5s.sun", color="white"))
            self.infoButton.setIcon(qta.icon("fa5s.info-circle", color="white"))
            self.btnReset.setIcon(qta.icon("fa5s.undo", color="white"))
            self.btnCopy.setIcon(qta.icon("fa5.copy", color="white"))
            self.settingsButton.setIcon(qta.icon("fa5s.cog", color="white"))

        self.is_dark_mode = not self.is_dark_mode

    def showInfoDialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Usage Instructions")
        dialog.setMaximumHeight(500)
        dialog.setMinimumWidth(800) 

        layout = QVBoxLayout(dialog)

        info_text = """<html>
        <body style="font-size:12pt; font-family:'Segoe UI';">
            <b>Instructions:</b><br><br>
            1. Select a Directory: Click 'Select SWASH Directory' to choose the directory containing TOXSWA `.sum` files.<br><br>
            2. Selecting Files: Choose a project and select `.sum` files, or enable 'Batch Mode' to process multiple projects.<br><br>
            3. Optional: Add summary sheet with PECs for each batch mode project. Select the desired order in the output table.<br><br>
            4. Configuring Filters:<br>
            &nbsp;&nbsp;&nbsp;- Enter a RAC value to highlight exceedance.<br>
            &nbsp;&nbsp;&nbsp;- Choose Parent/Metabolite compounds.<br>
            &nbsp;&nbsp;&nbsp;- Sort files by filename, compound, or scenario.<br><br>
            5. Extract Data: Click 'Extract Data' to process selected files.<br><br>
            6. Export to Excel: Click 'Export to Excel', select a save location, and optionally open the file.<br><br>
            <b>Settings:</b>
            <ul>
                <li><b>Highlight mitigation exceeding 95%:</b> This applies bold formatting to Summary sheet PECs where scenarios require mitigation above the FOCUS trigger (Note: Ensure only one set of Step 3 data is included per batch).</li><br>
            </ul>
        </body>
        </html>"""
        label = QLabel(dialog)
        label.setText(info_text)
        label.setWordWrap(True)
        layout.addWidget(label)

        ok_button = QPushButton("OK", dialog)
        ok_button.clicked.connect(dialog.accept)
        layout.addWidget(ok_button)

        dialog.exec_()

    def selectDirectory(self):
        directory = QFileDialog.getExistingDirectory(
            self, "Select SWASH Directory", "C:/SwashProjects"
        )
        if directory:
            self.main_dir = directory
            self.updateSubfolderList()

    def updateSubfolderList(self):
        self.subfolderDropdown.clear()
        if self.main_dir and not self.batch_mode:
            subfolders = [
                f
                for f in os.listdir(self.main_dir)
                if os.path.isdir(os.path.join(self.main_dir, f))
            ]
            self.subfolderDropdown.addItems(subfolders)
            if subfolders:
                self.subfolderDropdown.setCurrentIndex(0)
        self.updateFileList()

    def updateFileList(self):
        self.fileList.clear()
        if not self.main_dir:
            return
        if self.batch_mode:
            self.fileListLabel.setText("Select projects:")
            self.fileList.addItems(
                [
                    f
                    for f in os.listdir(self.main_dir)
                    if os.path.isdir(os.path.join(self.main_dir, f))
                ]
            )
        else:
            self.fileListLabel.setText(".sum Files:")
            self.subfolder = self.subfolderDropdown.currentText()
            subfolder_path = os.path.join(self.main_dir, self.subfolder, "toxswa")
            if os.path.exists(subfolder_path):
                files = sorted(
                    [f for f in os.listdir(subfolder_path) if f.endswith(".sum")]
                )
                self.fileList.addItems(files)

    def toggleBatchMode(self, state):
        if not self.main_dir:
            QMessageBox.warning(
                self, "No Directory Selected", "Please select a SWASH directory first."
            )
            self.batchCheckbox.setChecked(False)
            return
        self.batch_mode = state == Qt.Checked
        self.summaryCheckbox.setVisible(self.batch_mode)
        if self.batch_mode:
            self.projectLabel.hide()
            self.subfolderDropdown.hide()
        else:
            self.projectLabel.show()
            self.subfolderDropdown.show()
            self.updateSubfolderList()
        self.updateFileList()

    def extractData(self):
        try:
            if not self.main_dir:
                QMessageBox.warning(self, "Error", "Please select a directory first!")
                return

            selected_files = [item.text() for item in self.fileList.selectedItems()]
            if not selected_files:
                QMessageBox.warning(self, "Error", "Please select files to process!")
                return

            self.all_data.clear()
            self.project_shortcodes.clear()

            if self.batch_mode:
                for project in selected_files:
                    project_path = os.path.join(self.main_dir, project, "toxswa")
                    if os.path.exists(project_path):
                        self.processFiles(project_path, project)
            else:
                subfolder_path = os.path.join(self.main_dir, self.subfolder, "toxswa")
                if os.path.exists(subfolder_path):
                    self.processFiles(subfolder_path, self.subfolder, selected_files)
                else:
                    QMessageBox.critical(
                        self, "Error", f"Path not found: {subfolder_path}"
                    )
                    return

            self.updateTable()

            if self.summaryCheckbox.isChecked():
                self.toggleSummaryOrderList(Qt.Checked)

        except Exception as e:
            QMessageBox.critical(
                self, "Extraction Error", f"Failed to extract data: {str(e)}"
            )

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

    def processFiles(self, folder_path, project_name, selected_files=None):
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

            # Extract Areic mean deposition value and print it for verification
            areic_value = extract_areic_mean_deposition(content)
            print(f"File {filename}: Areic mean deposition = {areic_value}")

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

            parent_compound = self.extractValue(
                content, r"\* Substance\s*:\s*(\S+)", "Unknown"
            )
            parent_max_sw = self.extractFloat(content, r"Global max.*?([\d.]+)")
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
                soil = self.extractValue(content, r"\* Soil metabolite:\s*(\S+)", "")
                if soil and soil != parent_compound:
                    subs.append(soil)
            for sub in subs:
                max_sw_str = self.extractValue(
                    content,
                    rf"\* Table: PEC in water layer of substance:\s+{re.escape(sub)}.*?Global max\s+([<]?\s*\S+)",
                    "0",
                )
                max_sed_str = self.extractValue(
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

    def updateTable(self):
        if self.areic_comparison_enabled:
            headers = [
                "Project", "Filename", "Compound", "Scenario", "Waterbody",
                "Max PECsw", "Max PECsed", "Areic dep.", "Route of entry"
            ]
            col_count = 9
        else:
            headers = [
                "Project", "Filename", "Compound", "Scenario", "Waterbody",
                "Max PECsw", "Max PECsed", "Route of entry"
            ]
            col_count = 8

        self.tableWidget.setColumnCount(col_count)
        for i, header_text in enumerate(headers):
            header_item = QTableWidgetItem(header_text)

            if header_text == "Areic dep.":
                header_item.setToolTip("mg/m²")
            self.tableWidget.setHorizontalHeaderItem(i, header_item)

        self.tableWidget.setRowCount(0)  

        if not self.all_data:
            return

        try:
            pnec_value = float(self.pnecInput.text()) if self.pnecInput.text() else None
        except ValueError:
            pnec_value = None

        filtered_rows = []
        for project, rows in self.all_data.items():
            for row in rows:
                if row.get("Type", "") == self.compoundTypeDropdown.currentText():
                    filtered_rows.append((project, row))

        if not filtered_rows:
            self.tableWidget.setRowCount(1)
            placeholder = QTableWidgetItem(f"No {self.compoundTypeDropdown.currentText()} data available")
            placeholder.setFlags(placeholder.flags() & ~Qt.ItemIsEditable)
            self.tableWidget.setItem(0, 0, placeholder)
            return

        def get_sum_number(filename):
            m = re.search(r"(\d+)\.sum$", filename)
            return int(m.group(1)) if m else 999999999

        sort_option = self.sortDropdown.currentText()
        if sort_option == "File number":
            sort_key = lambda r: get_sum_number(r["Filename"])
        elif sort_option == "Compound":
            sort_key = lambda r: r["Compound"].upper()
        elif sort_option == "Scenario":
            sort_key = lambda r: r["Scenario"].upper()
        else:
            sort_key = lambda r: r["Filename"]

        filtered_rows = sorted(filtered_rows, key=lambda pr: sort_key(pr[1]))
        self.tableWidget.setRowCount(len(filtered_rows))

        row_index = 0
        for project, row in filtered_rows:
            data = [
                project,
                row["Filename"],
                row["Compound"],
                row["Scenario"],
                row["Waterbody"],
                self.format_for_display(row.get("Max PECsw"), row.get("Type")) if row.get("Max PECsw") else "0",
                self.format_for_display(row.get("Max PECsed"), row.get("Type")) if row.get("Max PECsed") else "0",
            ]
            if self.areic_comparison_enabled:
                data.append(row.get("Areic mean deposition", "N/A"))
            data.append(row["Route"])

            # Insert each data item into the table.
            for col, value in enumerate(data):
                item = QTableWidgetItem(str(value))
                
                if pnec_value is not None and col == 5:
                    try:
                        if float(str(value).replace("<1E-06", "0.000001")) > pnec_value:
                            item.setForeground(QColor(255, 0, 0))
                    except ValueError:
                        pass
                self.tableWidget.setItem(row_index, col, item)
            row_index += 1

        self.tableWidget.setColumnWidth(0, 170)
        self.tableWidget.setColumnWidth(col_count - 1, 170)
        
    def updateTable(self):
        # Decide on headers based on whether areic comparison is enabled.
        if self.areic_comparison_enabled:
            headers = [
                "Project", "Filename", "Compound", "Scenario", "Waterbody",
                "Max PECsw", "Max PECsed", "Areic dep.", "Route of entry"
            ]
            col_count = 9
        else:
            headers = [
                "Project", "Filename", "Compound", "Scenario", "Waterbody",
                "Max PECsw", "Max PECsed", "Route of entry"
            ]
            col_count = 8

        # Set column count and create header items with tooltips.
        self.tableWidget.setColumnCount(col_count)
        for i, header_text in enumerate(headers):
            header_item = QTableWidgetItem(header_text)
            # Set header tooltips for units.
            if header_text in ["Max PECsw", "Max PECsed"]:
                header_item.setToolTip("µg/L")
            elif header_text == "Areic dep.":
                header_item.setToolTip("mg/m²")
            self.tableWidget.setHorizontalHeaderItem(i, header_item)

        # Clear any existing rows.
        self.tableWidget.setRowCount(0)
        if not self.all_data:
            return

        # Get RAC filter value, if any.
        try:
            pnec_value = float(self.pnecInput.text()) if self.pnecInput.text() else None
        except ValueError:
            pnec_value = None

        # Filter rows based on selected compound type.
        filtered_rows = []
        for project, rows in self.all_data.items():
            for row in rows:
                if row.get("Type", "") == self.compoundTypeDropdown.currentText():
                    filtered_rows.append((project, row))
        if not filtered_rows:
            self.tableWidget.setRowCount(1)
            placeholder = QTableWidgetItem(
                f"No {self.compoundTypeDropdown.currentText()} data available"
            )
            placeholder.setFlags(placeholder.flags() & ~Qt.ItemIsEditable)
            self.tableWidget.setItem(0, 0, placeholder)
            return

        # Sort rows (example: by file number)
        def get_sum_number(filename):
            m = re.search(r"(\d+)\.sum$", filename)
            return int(m.group(1)) if m else 999999999

        sort_option = self.sortDropdown.currentText()
        if sort_option == "File number":
            sort_key = lambda pr: get_sum_number(pr[1]["Filename"])
        elif sort_option == "Compound":
            sort_key = lambda pr: pr[1]["Compound"].upper()
        elif sort_option == "Scenario":
            sort_key = lambda pr: pr[1]["Scenario"].upper()
        elif sort_option == "Filename":
            # Group by project first, then by filename.
            sort_key = lambda pr: (pr[0].upper(), pr[1]["Filename"].upper())
        else:
            sort_key = lambda pr: pr[1]["Filename"]

        filtered_rows = sorted(filtered_rows, key=sort_key)
        self.tableWidget.setRowCount(len(filtered_rows))

        # Populate rows.
        row_index = 0
        for project, row in filtered_rows:
            # Build the row data list.
            data = [
                project,
                row["Filename"],
                row["Compound"],
                row["Scenario"],
                row["Waterbody"],
                self.format_for_display(row.get("Max PECsw"), row.get("Type"))
                if row.get("Max PECsw")
                else "0",
                self.format_for_display(row.get("Max PECsed"), row.get("Type"))
                if row.get("Max PECsed")
                else "0",
            ]
            if self.areic_comparison_enabled:
                data.append(row.get("Areic mean deposition", "N/A"))
            data.append(row["Route"])  # "Route of entry" always comes last.

            # Insert each data item into the table.
            for col, value in enumerate(data):
                item = QTableWidgetItem(str(value))
                # Set tooltip on the "Project" column (assumed column 0).
                if col == 0:
                    item.setToolTip(str(value))
                if pnec_value is not None and col == 5:
                    try:
                        if float(str(value).replace("<1E-06", "0.000001")) > pnec_value:
                            item.setForeground(QColor(255, 0, 0))
                    except ValueError:
                        pass
                self.tableWidget.setItem(row_index, col, item)
            row_index += 1

        # Set fixed column widths.
        self.tableWidget.setColumnWidth(0, 170)   # Project
        self.tableWidget.setColumnWidth(1, 200)   # Filename
        self.tableWidget.setColumnWidth(2, 150)   # Compound
        self.tableWidget.setColumnWidth(3, 100)   # Scenario
        self.tableWidget.setColumnWidth(4, 100)   # Waterbody
        self.tableWidget.setColumnWidth(5, 120)   # Max PECsw
        self.tableWidget.setColumnWidth(6, 120)   # Max PECsed
        if self.areic_comparison_enabled:
            self.tableWidget.setColumnWidth(7, 120)   # Areic dep.
            self.tableWidget.setColumnWidth(8, 180)   # Route of entry
        else:
            self.tableWidget.setColumnWidth(7, 180)   # Route of entry

    def exportToExcel(self):
        if not self.all_data:
            return
        filePath, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "", "Excel Files (*.xlsx)"
        )
        if not filePath:
            return

        workbook = xlsxwriter.Workbook(filePath)

        if self.batch_mode and self.summaryCheckbox.isChecked():
            self.createSummarySheet(workbook)

        right_align = workbook.add_format({"align": "right"})
        header_format = workbook.add_format({
            "bg_color": "#82C940",
            "font_color": "#000000",
            "bold": True,
            "align": "left",
        })

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

        # Initialize a set to track existing worksheet names
        existing_sheet_names = set()

        for project, rows in self.all_data.items():
            # Use the updated safe_sheet_name with the set
            worksheet = workbook.add_worksheet(safe_sheet_name(project, existing_sheet_names))
            sorted_rows = sorted(
                rows,
                key=lambda r: (
                    r["Compound"].upper(),
                    (
                        int(re.search(r"(\d+)\.sum$", r["Filename"]).group(1))
                        if re.search(r"(\d+)\.sum$", r["Filename"])
                        else 999999999
                    ),
                ),
            )

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
                app_dates = [extract_date_only(d) for d in app_dates]
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
                    pecsw_vals = [extract_daily_value(content, key, version, 0) for key in sw_daily_search]
                    twaecsw_vals = [extract_daily_value(content, key, version, 0) for key in twaec_sw_search]
                else:
                    pecsw_vals = [extract_daily_value(content, key, version, water_start) for key in sw_daily_search]
                    twaecsw_vals = [extract_daily_value(content, key, version, water_start) for key in twaec_sw_search]

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
                app_dates = [extract_date_only(d) for d in app_dates]
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
                    pecsed_vals = [extract_daily_value(content, key, version, 0) for key in sed_daily_search]
                    twaecsed_vals = [extract_daily_value(content, key, version, 0) for key in twaec_sed_headers]
                else:
                    pecsed_vals = [extract_daily_value(content, key, version, sed_start) for key in sed_daily_search]
                    twaecsed_vals = [extract_daily_value(content, key, version, sed_start) for key in twaec_sed_headers]

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

            try:
                pnec_value = float(self.pnecInput.text())
            except:
                pnec_value = None
            if pnec_value is not None:
                sw_data_start = 1
                sw_data_end = sw_data_start + len(sorted_rows) - 1
                worksheet.conditional_format(
                    sw_data_start,
                    6,
                    sw_data_end,
                    6,
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": pnec_value,
                        "format": workbook.add_format({"font_color": "red"}),
                    },
                )
                sed_data_start = len(sorted_rows) + 3
                sed_data_end = sed_data_start + len(sorted_rows) - 1
                worksheet.conditional_format(
                    sed_data_start,
                    6,
                    sed_data_end,
                    6,
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": pnec_value,
                        "format": workbook.add_format({"font_color": "red"}),
                    },
                )

        workbook.close()
        if self.chkOpenExcel.isChecked():
            QDesktopServices.openUrl(QUrl.fromLocalFile(filePath))

    def format_for_excel(self, value):
        """
        Returns "<1E-06" if the value is less than or equal to 0.000001;
        otherwise, returns the value as a float.
        """
        try:
            num = float(value)
            if num <= 1E-6:
                return "<1E-06"
            return num
        except Exception:
            return value

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

    def createSummarySheet(self, workbook):
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
            if hasattr(self, "projectOrderList") and self.projectOrderList.isVisible():
                project_order = [self.projectOrderList.item(i).text()
                                for i in range(self.projectOrderList.count())]
            else:
                project_order = list(self.all_data.keys())
            
            # Build the baseline map from Step 3 files.
            # This map uses a key (scenario, waterbody, compound, type) and
            # stores the areic deposition value from files whose project shortcode contains "Step 3".
            step3_map = self.collect_step3_areic_map()
            
            # --- Set up the header rows.
            # Always include the first 3 fixed columns.
            summary_ws.merge_range("A1:A2", "Compound", header_format)
            summary_ws.merge_range("B1:B2", "Scenario", header_format)
            summary_ws.merge_range("C1:C2", "Waterbody", header_format)
            
            # For each project, determine how many columns to create.
            # If areic comparison is enabled, add 3 columns (PECsw, PECsed, Areic dep.)
            # Otherwise, add only 2 columns (PECsw, PECsed).
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
            # Parents are sorted by scenario and waterbody first
            # Metabolites are sorted by compound name first, then scenario and waterbody
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
                                print(f"[{project}] {scenario}/{compound} ({ctype}): ratio = {ratio:.2f}%")
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
            QMessageBox.critical(self, "Summary Error", f"Error: {str(e)}")

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

    def extractValue(self, content, pattern, default_value):
        m = re.search(pattern, content, re.DOTALL)
        return m.group(1).strip() if m else default_value

    def extractFloat(self, content, pattern):
        m = re.search(pattern, content)
        if not m:
            return None
        val_str = m.group(1).strip()
        return self.parse_value(val_str)

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
                return "<1E-06"  # Changed from "< " + str(numeric) to be consistent
            except:
                return "<1E-06"
        try:
            numeric_val = float(value_str)
            if numeric_val < 1E-6:
                return "<1E-06"
            return numeric_val
        except:
            return None

    def showContextMenu(self, pos):
        menu = QMenu()
        copyAction = menu.addAction("Copy")
        action = menu.exec_(self.tableWidget.mapToGlobal(pos))
        if action == copyAction:
            self.copySelection()

    def copySelection(self):
        ranges = self.tableWidget.selectedRanges()
        if not ranges:
            return
        text = ""
        for r in ranges:
            for row in range(r.topRow(), r.bottomRow() + 1):
                row_text = []
                for col in range(r.leftColumn(), r.rightColumn() + 1):
                    item = self.tableWidget.item(row, col)
                    row_text.append(item.text() if item else "")
                text += "\t".join(row_text) + "\n"
        QApplication.clipboard().setText(text)

    def copyTableToClipboard(self):
        """Copy entire table content to clipboard in tab-separated format."""
        clipboard_text = ""

        colCount = self.tableWidget.columnCount()
        rowCount = self.tableWidget.rowCount()

        headers = [
            self.tableWidget.horizontalHeaderItem(col).text() for col in range(colCount)
        ]
        clipboard_text += "\t".join(headers) + "\n"

        for row in range(rowCount):
            row_data = []
            for col in range(colCount):
                item = self.tableWidget.item(row, col)
                row_data.append(item.text() if item else "")
            clipboard_text += "\t".join(row_data) + "\n"

        QApplication.clipboard().setText(clipboard_text)
        QMessageBox.information(self, "Copied", "Table data copied to clipboard.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    script_dir = os.path.dirname(os.path.realpath(__file__))
    icon_path = os.path.join(script_dir, "logo.ico")
    app.setWindowIcon(QIcon(icon_path))
    ex = TOXSWAExtractor()
    ex.show()
    sys.exit(app.exec_())
