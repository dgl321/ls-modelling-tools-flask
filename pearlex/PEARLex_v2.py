import os
import sys
import re

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
    QListWidget,
    QComboBox,
    QMessageBox,
    QHeaderView,
    QMenu,
    QCheckBox,
    QInputDialog,
)
from PyQt5.QtGui import QPixmap, QColor, QFont, QIcon, QDesktopServices
from PyQt5.QtCore import Qt, QUrl, QSize, QTimer
import qtawesome as qta
import xlsxwriter

DESKTOP_SERVICES_AVAILABLE = True


class PearlGroundwaterExtractor(QWidget):
    def __init__(self):
        super().__init__()
        self.main_dir = ""
        self.sum_filepaths = []
        self.all_data = []
        self.is_dark_mode = True
        self.batches = []

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
        QPushButton:hover {
            background-color: #D0D0D0;
        }
        QLineEdit, QComboBox {
            background-color: #FFFFFF;
            border: 1px solid #A0A0A0;
            padding: 4px;
            border-radius: 3px;
        }
        QListWidget {
            background-color: #FFFFFF;
            border: 1px solid #A0A0A0;
        }
        QLabel { color: #202020; }
        """
        QApplication.instance().setStyleSheet(self.dark_stylesheet)
        self.initUI()

    def initUI(self):
        self.setWindowTitle("PEARLex")
        self.setGeometry(200, 200, 1175, 600)
        main_layout = QHBoxLayout()

        # LEFT LAYOUT: File selection and logo
        left_layout = QVBoxLayout()
        self.logoLabel = QLabel()
        script_dir = os.path.dirname(os.path.realpath(__file__))
        logo_path = os.path.join(script_dir, "ls.png")
        pixmap = QPixmap(logo_path)
        if not pixmap.isNull():
            pixmap = pixmap.scaledToWidth(200, Qt.SmoothTransformation)
            self.logoLabel.setPixmap(pixmap)
        self.logoLabel.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(self.logoLabel)

        self.subheadingLabel = QLabel("PEARLex v1")
        sh_font = QFont("Lato", 14)
        sh_font.setWeight(500)
        self.subheadingLabel.setFont(sh_font)
        # Note: "colour" here is in the stylesheet as per UK grammar
        self.subheadingLabel.setStyleSheet(
            "font-size: 13pt; font-weight: 500; color: #bfb800;"
        )
        self.subheadingLabel.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(self.subheadingLabel)

        self.btnSelectDir = QPushButton("Select PearlDB Directory")
        self.btnSelectDir.clicked.connect(self.selectDirectory)
        left_layout.addWidget(self.btnSelectDir)

        self.fileListLabel = QLabel(".sum Files:")
        left_layout.addWidget(self.fileListLabel)

        self.fileList = QListWidget()
        self.fileList.setSelectionMode(QListWidget.MultiSelection)
        left_layout.addWidget(self.fileList)

        # RIGHT LAYOUT: Filters, buttons, table
        right_layout = QVBoxLayout()

        # TOP FILTERS
        top_filters = QHBoxLayout()
        self.compoundTypeDropdown = QComboBox()
        self.compoundTypeDropdown.addItems(["Parent", "Metabolite"])
        self.compoundTypeDropdown.currentIndexChanged.connect(self.updateTable)
        self.compoundTypeDropdown.view().setMinimumWidth(110)
        self.compoundTypeDropdown.setMinimumWidth(110)
        top_filters.addWidget(QLabel("Compound Type:"))
        top_filters.addWidget(self.compoundTypeDropdown)

        self.sortDropdown = QComboBox()
        self.sortDropdown.addItems(["Filename", "Compound", "Scenario"])
        self.sortDropdown.currentIndexChanged.connect(self.updateTable)
        self.sortDropdown.view().setMinimumWidth(110)
        self.sortDropdown.setMinimumWidth(110)
        top_filters.addWidget(QLabel("Sort by:"))
        top_filters.addWidget(self.sortDropdown)

        self.limitDropdown = QComboBox()
        self.limitDropdown.addItems(["", "0.1 µg/L", "0.001 µg/L"])
        self.limitDropdown.currentIndexChanged.connect(self.updateTable)
        self.limitDropdown.setMinimumWidth(100)
        self.limitDropdown.view().setMinimumWidth(100)
        top_filters.addWidget(QLabel("Parametric Limit:"))
        top_filters.addWidget(self.limitDropdown)

        self.themeToggleButton = QPushButton("Light Mode")
        self.themeToggleButton.setIcon(qta.icon("fa5s.sun", color="white"))
        self.themeToggleButton.clicked.connect(self.toggleTheme)
        top_filters.addWidget(self.themeToggleButton)

        self.infoButton = QPushButton()
        self.infoButton.setIcon(qta.icon("fa5s.info-circle", color="white"))
        self.infoButton.setToolTip("Click for usage instructions")
        self.infoButton.clicked.connect(self.showInfoDialog)
        top_filters.addWidget(self.infoButton)

        self.chkBatchMode = QCheckBox("Batch Mode")
        self.chkBatchMode.stateChanged.connect(self.toggleBatchMode)
        top_filters.addWidget(self.chkBatchMode)

        top_filters.addStretch()
        right_layout.addLayout(top_filters)

        # SINGLE MODE WIDGET (non-batch)
        self.singleWidget = QWidget()
        single_layout = QHBoxLayout(self.singleWidget)

        self.btnExtractSingle = QPushButton("Extract Data")
        self.btnExtractSingle.clicked.connect(self.extractData)
        single_layout.addWidget(self.btnExtractSingle)

        self.btnExportSingle = QPushButton("Export to Excel")
        self.btnExportSingle.clicked.connect(self.exportToExcelSingle)
        single_layout.addWidget(self.btnExportSingle)

        self.btnCopySingle = QPushButton()
        self.btnCopySingle.setIcon(qta.icon("fa5.copy", color="white"))
        self.btnCopySingle.setToolTip("Copy table to clipboard")
        self.btnCopySingle.setFixedSize(36, 36)
        self.btnCopySingle.setIconSize(QSize(19, 19))
        self.btnCopySingle.setStyleSheet("padding: 0px;")
        self.btnCopySingle.clicked.connect(self.copyTableToClipboard)
        single_layout.addWidget(self.btnCopySingle)

        self.btnResetSingle = QPushButton()
        self.btnResetSingle.setIcon(qta.icon("fa5s.undo", color="white"))
        self.btnResetSingle.setToolTip("Clear Table")
        self.btnResetSingle.setFixedSize(36, 36)
        self.btnResetSingle.setIconSize(QSize(19, 19))
        self.btnResetSingle.setStyleSheet("padding: 0px;")
        self.btnResetSingle.clicked.connect(self.clearData)
        single_layout.addWidget(self.btnResetSingle)

        self.chkOpenExcelSingle = QCheckBox("Open Excel after export")
        single_layout.addWidget(self.chkOpenExcelSingle)

        right_layout.addWidget(self.singleWidget)

        # BATCH MODE WIDGET
        self.batchWidget = QWidget()
        batch_layout = QHBoxLayout(self.batchWidget)

        self.btnExtractBatch = QPushButton("Extract Data")
        self.btnExtractBatch.clicked.connect(self.extractData)
        batch_layout.addWidget(self.btnExtractBatch)

        self.btnClearTable = QPushButton("Clear Table")
        self.btnClearTable.clicked.connect(self.clearData)
        batch_layout.addWidget(self.btnClearTable)

        self.btnAddBatch = QPushButton("Add to Batch")
        self.btnAddBatch.clicked.connect(self.addToBatch)
        batch_layout.addWidget(self.btnAddBatch)

        self.btnExportBatches = QPushButton("Export Batches")
        self.btnExportBatches.clicked.connect(self.exportBatches)
        batch_layout.addWidget(self.btnExportBatches)

        self.chkOpenExcelBatch = QCheckBox("Open Excel after export")
        batch_layout.addWidget(self.chkOpenExcelBatch)

        right_layout.addWidget(self.batchWidget)
        self.batchWidget.hide()  # Hidden by default

        # TABLE WIDGET
        self.tableWidget = QTableWidget()
        self.tableWidget.setColumnCount(6)
        self.tableWidget.setHorizontalHeaderLabels(
            [
                "Project",
                "Filename",
                "Compound Type",
                "Scenario",
                "Compound",
                "80th Percentile (µg/L)",
            ]
        )
        header = self.tableWidget.horizontalHeader()
        header.setStretchLastSection(False)
        for i in range(self.tableWidget.columnCount()):
            if i == 5:
                header.setSectionResizeMode(i, QHeaderView.Fixed)
                self.tableWidget.setColumnWidth(
                    i, 170
                )  # Force column 6 to be 285 pixels wide
            else:
                header.setSectionResizeMode(i, QHeaderView.Stretch)
        # Set a size hint on the header item for column 6
        header_item = self.tableWidget.horizontalHeaderItem(5)
        header_item.setSizeHint(QSize(285, header_item.sizeHint().height()))
        self.tableWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tableWidget.customContextMenuRequested.connect(self.showContextMenu)
        right_layout.addWidget(self.tableWidget)

        main_layout.addLayout(left_layout, 1)
        main_layout.addLayout(right_layout, 5)
        self.setLayout(main_layout)

        # Use QTimer to delay the adjustment until after the UI is shown
        QTimer.singleShot(0, self.adjustColumnWidth)

    def adjustColumnWidth(self):
        header = self.tableWidget.horizontalHeader()
        header.setStretchLastSection(False)
        self.tableWidget.setColumnWidth(5, 170)

    def toggleBatchMode(self, state):
        is_checked = state == Qt.Checked
        if is_checked:
            self.singleWidget.hide()
            self.batchWidget.show()
        else:
            self.batchWidget.hide()
            self.singleWidget.show()

    def showInfoDialog(self):
        info_text = (
            "<html><body>"
            "<b>Usage Instructions:</b><br><br>"
            "1. Select a directory containing .sum files.<br><br>"
            "2. If not in Batch Mode: Extract Data, then Export to Excel for Parent+Metabolite tables in one sheet.<br><br>"
            "3. If Batch Mode is checked: Extract Data → Add to Batch → Clear Table → repeat if needed. Then Export Batches (Parent+Met in separate sheets).<br>"
            "</body></html>"
        )
        msg = QMessageBox(self)
        msg.setWindowTitle("Usage Instructions")
        msg.setText(info_text)
        msg.setStyleSheet("QLabel { font-size: 11pt; }")
        msg.exec_()

    def showContextMenu(self, pos):
        menu = QMenu()
        actCopy = menu.addAction("Copy")
        action = menu.exec_(self.tableWidget.mapToGlobal(pos))
        if action == actCopy:
            self.copySelection()

    def copySelection(self):
        rngs = self.tableWidget.selectedRanges()
        if not rngs:
            return
        lines = []
        for r in rngs:
            for row in range(r.topRow(), r.bottomRow() + 1):
                row_items = []
                for col in range(r.leftColumn(), r.rightColumn() + 1):
                    it = self.tableWidget.item(row, col)
                    row_items.append(it.text() if it else "")
                lines.append("\t".join(row_items))
        QApplication.clipboard().setText("\n".join(lines))

    def copyTableToClipboard(self):
        row_count = self.tableWidget.rowCount()
        col_count = self.tableWidget.columnCount()
        if row_count == 0:
            QMessageBox.information(self, "Copy Table", "No data available to copy.")
            return
        copied_text = ""
        headers = [
            self.tableWidget.horizontalHeaderItem(col).text()
            for col in range(col_count)
        ]
        copied_text += "\t".join(headers) + "\n"
        for row in range(row_count):
            row_data = []
            for col in range(col_count):
                item = self.tableWidget.item(row, col)
                row_data.append(item.text() if item else "")
            copied_text += "\t".join(row_data) + "\n"
        QApplication.clipboard().setText(copied_text)
        QMessageBox.information(
            self, "Copy Table", "Table copied to clipboard successfully."
        )

    def toggleTheme(self):
        app = QApplication.instance()
        if self.is_dark_mode:
            app.setStyleSheet(self.light_stylesheet)
            self.themeToggleButton.setText("Dark Mode")
            self.themeToggleButton.setIcon(qta.icon("fa5s.moon", color="black"))
            self.infoButton.setIcon(qta.icon("fa5s.info-circle", color="black"))
            self.btnCopySingle.setIcon(qta.icon("fa5.copy", color="black"))
            self.btnResetSingle.setIcon(qta.icon("fa5s.undo", color="black"))
        else:
            app.setStyleSheet(self.dark_stylesheet)
            self.themeToggleButton.setText("Light Mode")
            self.themeToggleButton.setIcon(qta.icon("fa5s.sun", color="white"))
            self.infoButton.setIcon(qta.icon("fa5s.info-circle", color="white"))
            self.btnCopySingle.setIcon(qta.icon("fa5.copy", color="white"))
            self.btnResetSingle.setIcon(qta.icon("fa5s.undo", color="white"))
        self.is_dark_mode = not self.is_dark_mode

    def selectDirectory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select PearlDB Directory")
        if directory:
            self.main_dir = directory
            self.updateFileList()

    def updateFileList(self):
        self.fileList.clear()
        self.sum_filepaths.clear()
        if not self.main_dir:
            return
        for f in os.listdir(self.main_dir):
            if f.endswith(".sum"):
                self.sum_filepaths.append(os.path.join(self.main_dir, f))
        for sub in os.listdir(self.main_dir):
            subp = os.path.join(self.main_dir, sub)
            if os.path.isdir(subp):
                for ff in os.listdir(subp):
                    if ff.endswith(".sum"):
                        self.sum_filepaths.append(os.path.join(subp, ff))
        for path in self.sum_filepaths:
            self.fileList.addItem(os.path.basename(path))

    def extractData(self):
        sel = self.fileList.selectedIndexes()
        if not sel:
            QMessageBox.warning(
                self, "No Files", "Please select one or more .sum files."
            )
            return
        self.all_data.clear()
        for idx in sel:
            fp = self.sum_filepaths[idx.row()]
            try:
                with open(fp, "r", encoding="ISO-8859-1") as f:
                    content = f.read()
            except Exception as e:
                QMessageBox.warning(self, "File Error", f"Cannot read {fp}\n{str(e)}")
                continue
            p = re.search(r"Application_scheme\s+(\S+)", content)
            project = p.group(1) if p else "Unknown"
            s = re.search(r"Location\s*[:]*\s*(.*)", content)
            scenario_raw = s.group(1).strip() if s else "Unknown"
            scenario = scenario_raw.capitalize()
            comp_list = re.findall(r"Result_(\S+)\s+([\d.]+)", content)
            for i, (comp, val_str) in enumerate(comp_list):
                ctype = "Parent" if i == 0 else "Metabolite"
                try:
                    val = float(val_str)
                except:
                    val = 0.0
                self.all_data.append(
                    [project, os.path.basename(fp), scenario, comp, val, ctype]
                )
        self.updateTable()

    def updateTable(self):
        target_type = self.compoundTypeDropdown.currentText()
        sort_field = self.sortDropdown.currentText()
        limit_val = None
        if not self.limitDropdown.currentText().startswith("("):
            try:
                limit_val = float(self.limitDropdown.currentText().split()[0])
            except:
                pass
        data_for_table = [r for r in self.all_data if r[5] == target_type]
        if sort_field == "Filename":
            data_for_table.sort(key=lambda x: x[1].lower())
        elif sort_field == "Compound":
            data_for_table.sort(key=lambda x: x[3].lower())
        elif sort_field == "Scenario":
            data_for_table.sort(key=lambda x: x[2].lower())
        self.tableWidget.setRowCount(len(data_for_table))
        for row_i, row_data in enumerate(data_for_table):
            display = [
                row_data[0],
                row_data[1],
                row_data[5],
                row_data[2],
                row_data[3],
                row_data[4],
            ]
            for col_i, val in enumerate(display):
                item = QTableWidgetItem(str(val))
                if col_i == 5 and limit_val is not None:
                    try:
                        if float(val) > limit_val:
                            item.setForeground(QColor(255, 0, 0))
                    except:
                        pass
                self.tableWidget.setItem(row_i, col_i, item)

    def clearData(self):
        self.all_data.clear()
        self.tableWidget.setRowCount(0)

    def addToBatch(self):
        if not self.all_data:
            QMessageBox.warning(self, "No Data", "Extract data first.")
            return
        batch_name, ok = QInputDialog.getText(self, "Batch Name", "Enter batch name:")
        if not ok or not batch_name.strip():
            batch_name = f"Batch_{len(self.batches) + 1}"
        new_copy = [row[:] for row in self.all_data]
        self.batches.append((batch_name, new_copy))
        QMessageBox.information(self, "Batch Added", "Batch added successfully.")

    def exportToExcelSingle(self):
        if not self.all_data:
            QMessageBox.warning(self, "No Data", "No extracted data.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        try:
            limit_val = None
            ls = self.limitDropdown.currentText()
            if not ls.startswith("("):
                try:
                    limit_val = float(ls.split()[0])
                except:
                    pass
            parents = []
            mets = []
            for row in self.all_data:
                if row[5] == "Parent":
                    parents.append(row)
                else:
                    mets.append(row)
            parents.sort(key=lambda x: x[1].lower())
            mets.sort(key=lambda x: x[1].lower())
            wb = xlsxwriter.Workbook(path)
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
                for col_i, h in enumerate(columns):
                    ws.write(row_cursor, col_i, h, header_fmt)
                row_cursor += 1
                for rdat in data_rows:
                    reorder = [rdat[0], rdat[1], rdat[2], rdat[3], rdat[4]]
                    for cc, valx in enumerate(reorder):
                        txt = str(valx)
                        if cc == 4:
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
                for cidx in range(len(columns)):
                    ws.set_column(cidx, cidx, col_width[cidx] + 2)
                return row_cursor

            nextrow = write_table(0, parents, "Parent Table")
            nextrow += 1
            write_table(nextrow, mets, "Metabolite Table")
            wb.close()
            if self.chkOpenExcelSingle.isChecked():
                QDesktopServices.openUrl(QUrl.fromLocalFile(path))
            QMessageBox.information(self, "Exported", "Exported successfully.")
        except Exception as e:
            QMessageBox.warning(self, "Export Error", str(e))

    def exportBatches(self):
        if not self.batches:
            QMessageBox.warning(self, "No Batches", "No batches to export.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        try:
            limit_val = None
            ls = self.limitDropdown.currentText()
            if not ls.startswith("("):
                try:
                    limit_val = float(ls.split()[0])
                except:
                    pass
            wb = xlsxwriter.Workbook(path)
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
                for c_i, h in enumerate(columns):
                    ws.write(rowcur, c_i, h, hfmt)
                rowcur += 1
                for rowd in data_rows:
                    reorder = [rowd[0], rowd[1], rowd[2], rowd[3], rowd[4]]
                    for cc, vv in enumerate(reorder):
                        txt = str(vv)
                        if cc == 4:
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
                for c_col in range(len(columns)):
                    ws.set_column(c_col, c_col, cw[c_col] + 2)
                return rowcur

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
                wsheet = wb.add_worksheet(sheet_name[:31])
                next_r = write_section(wsheet, 0, p, "Parent Table")
                next_r += 1
                write_section(wsheet, next_r, m, "Metabolite Table")
            wb.close()
            if self.chkBatchMode.isChecked():
                if self.chkOpenExcelBatch.isChecked():
                    QDesktopServices.openUrl(QUrl.fromLocalFile(path))
            else:
                if self.chkOpenExcelSingle.isChecked():
                    QDesktopServices.openUrl(QUrl.fromLocalFile(path))
            QMessageBox.information(self, "Exported", "Batches exported successfully.")
        except Exception as e:
            QMessageBox.warning(self, "Export Error", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    script_dir = os.path.dirname(os.path.realpath(__file__))
    icon_path = os.path.join(script_dir, "logo.ico")
    app.setWindowIcon(QIcon(icon_path))
    ex = PearlGroundwaterExtractor()
    ex.show()
    sys.exit(app.exec_())
