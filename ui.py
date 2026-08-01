"""
UI Module
PyQt5 GUI for VBOX File Renamer application.
"""

import os
import sys
from typing import List, Optional
import pandas as pd
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QLineEdit, QFileDialog, QTableWidget, QTableWidgetItem, QComboBox, 
    QSpinBox, QCheckBox, QMessageBox, QProgressBar, QTabWidget, QGroupBox,
    QHeaderView, QApplication, QMenu, QAction, QDialogButtonBox, QDialog
)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QIcon, QColor, QFont, QDrag, QPixmap
from PyQt5.QtCore import QMimeData

from excel_reader import ExcelReader
from renamer import VBOXRenamer, FileRenameInfo
from matcher import FileMatcher


class ColumnMappingDialog(QDialog):
    """Dialog for manually mapping Excel columns."""

    def __init__(self, parent, available_columns: List[str], current_mapping: dict):
        """
        Initialize column mapping dialog.

        Args:
            parent: Parent widget
            available_columns: List of available column names
            current_mapping: Current mapping dict
        """
        super().__init__(parent)
        self.setWindowTitle("Map Excel Columns")
        self.setGeometry(100, 100, 500, 300)
        self.available_columns = available_columns
        self.mapping = current_mapping.copy()
        self.init_ui()

    def init_ui(self):
        """Initialize UI elements."""
        layout = QVBoxLayout()

        # Required fields
        required_fields = ['start_point', 'end_point', 'road_type']
        optional_fields = ['serial', 'distance']

        for field in required_fields + optional_fields:
            row_layout = QHBoxLayout()
            field_label = QLabel(f"{field.replace('_', ' ').title()}:")
            field_label.setMinimumWidth(100)

            combo = QComboBox()
            combo.addItem("-- Not Found --")
            combo.addItems(self.available_columns)

            if field in self.mapping:
                combo.setCurrentText(self.mapping[field])

            combo.currentTextChanged.connect(lambda text, f=field: self._on_mapping_changed(f, text))
            combo.setMinimumWidth(300)

            row_layout.addWidget(field_label)
            row_layout.addWidget(combo)
            layout.addLayout(row_layout)

        # Buttons
        button_layout = QHBoxLayout()
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def _on_mapping_changed(self, field: str, column: str):
        """Handle mapping change."""
        if column != "-- Not Found --":
            self.mapping[field] = column
        else:
            self.mapping.pop(field, None)

    def get_mapping(self) -> dict:
        """Get the current mapping."""
        return self.mapping


class RenameWorkerThread(QThread):
    """Worker thread for rename operations."""

    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)
    status_update = pyqtSignal(str)

    def __init__(self, renamer: VBOXRenamer, rename_list: List[FileRenameInfo]):
        """
        Initialize worker thread.

        Args:
            renamer: VBOXRenamer instance
            rename_list: List of files to rename
        """
        super().__init__()
        self.renamer = renamer
        self.rename_list = rename_list

    def run(self):
        """Run the rename operation."""
        total = len(self.rename_list)

        for i, file_info in enumerate(self.rename_list):
            self.status_update.emit(f"Processing {file_info.filename}...")
            self.progress.emit(int((i / total) * 100))
            # Actual rename happens in main thread in batches

        successful, summary = self.renamer.rename_files(self.rename_list)
        self.progress.emit(100)
        self.finished.emit(successful > 0, summary)


class VBOXRenamerUI(QMainWindow):
    """Main application window."""

    def __init__(self):
        """Initialize the application."""
        super().__init__()
        self.setWindowTitle("VBOX Test Suite File Renamer")
        self.setGeometry(100, 100, 1200, 800)

        # Application state
        self.folder_path = None
        self.excel_path = None
        self.excel_reader = ExcelReader()
        self.renamer = None
        self.preview_data = []

        self.setAcceptDrops(True)
        self.init_ui()
        self.apply_styles()

    def init_ui(self):
        """Initialize UI elements."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout()

        # Create tab widget
        tabs = QTabWidget()

        # Tab 1: File Selection
        tab1 = self._create_file_selection_tab()
        tabs.addTab(tab1, "Select Files")

        # Tab 2: Column Mapping
        tab2 = self._create_column_mapping_tab()
        tabs.addTab(tab2, "Map Columns")

        # Tab 3: Preview
        tab3 = self._create_preview_tab()
        tabs.addTab(tab3, "Preview")

        # Tab 4: Settings
        tab4 = self._create_settings_tab()
        tabs.addTab(tab4, "Settings")

        main_layout.addWidget(tabs)

        # Status bar
        self.status_label = QLabel("Ready")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.status_label)
        main_layout.addWidget(self.progress_bar)

        central_widget.setLayout(main_layout)

    def _create_file_selection_tab(self) -> QWidget:
        """Create file selection tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        # Folder selection
        folder_group = QGroupBox("VBOX Files Folder")
        folder_layout = QHBoxLayout()

        self.folder_input = QLineEdit()
        self.folder_input.setReadOnly(True)
        self.folder_input.setPlaceholderText("Select folder containing VBOX files...")

        folder_btn = QPushButton("Select Folder")
        folder_btn.clicked.connect(self._select_folder)
        folder_btn.setMinimumWidth(120)

        folder_layout.addWidget(self.folder_input)
        folder_layout.addWidget(folder_btn)
        folder_group.setLayout(folder_layout)
        layout.addWidget(folder_group)

        # Excel selection
        excel_group = QGroupBox("Excel File with Sequence Data")
        excel_layout = QHBoxLayout()

        self.excel_input = QLineEdit()
        self.excel_input.setReadOnly(True)
        self.excel_input.setPlaceholderText("Select Excel file with route data...")

        excel_btn = QPushButton("Select Excel")
        excel_btn.clicked.connect(self._select_excel)
        excel_btn.setMinimumWidth(120)

        excel_layout.addWidget(self.excel_input)
        excel_layout.addWidget(excel_btn)
        excel_group.setLayout(excel_layout)
        layout.addWidget(excel_group)

        # File info
        info_group = QGroupBox("File Information")
        info_layout = QHBoxLayout()

        self.file_count_label = QLabel("Files: 0")
        self.excel_rows_label = QLabel("Excel rows: 0")

        info_layout.addWidget(self.file_count_label)
        info_layout.addStretch()
        info_layout.addWidget(self.excel_rows_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # Action buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        load_preview_btn = QPushButton("Load Preview")
        load_preview_btn.clicked.connect(self._generate_preview)
        load_preview_btn.setMinimumWidth(150)

        button_layout.addWidget(load_preview_btn)
        layout.addLayout(button_layout)

        layout.addStretch()
        widget.setLayout(layout)
        return widget

    def _create_column_mapping_tab(self) -> QWidget:
        """Create column mapping tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        info_label = QLabel(
            "The application will attempt to auto-detect Excel columns.\n"
            "If auto-detection fails or is incorrect, manually map columns below."
        )
        layout.addWidget(info_label)

        # Mapping fields
        fields_group = QGroupBox("Column Mapping")
        fields_layout = QVBoxLayout()

        self.column_combos = {}
        required_fields = ['start_point', 'end_point', 'road_type']
        optional_fields = ['serial', 'distance']

        for field in required_fields + optional_fields:
            row_layout = QHBoxLayout()
            field_label = QLabel(f"{field.replace('_', ' ').title()}:")
            field_label.setMinimumWidth(120)

            combo = QComboBox()
            combo.addItem("-- Not Detected --")
            combo.setMinimumWidth(300)

            self.column_combos[field] = combo

            row_layout.addWidget(field_label)
            row_layout.addWidget(combo)
            row_layout.addStretch()
            fields_layout.addLayout(row_layout)

        fields_group.setLayout(fields_layout)
        layout.addWidget(fields_group)

        # Auto-detect button
        button_layout = QHBoxLayout()
        auto_btn = QPushButton("Auto-Detect Columns")
        auto_btn.clicked.connect(self._auto_detect_columns)

        button_layout.addStretch()
        button_layout.addWidget(auto_btn)
        layout.addLayout(button_layout)

        layout.addStretch()
        widget.setLayout(layout)
        return widget

    def _create_preview_tab(self) -> QWidget:
        """Create preview tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        # Summary
        summary_group = QGroupBox("Preview Summary")
        summary_layout = QHBoxLayout()

        self.matched_label = QLabel("Matched: 0")
        self.duplicate_label = QLabel("Duplicates: 0")
        self.conflict_label = QLabel("Conflicts: 0")
        self.not_found_label = QLabel("Not Found: 0")
        self.unchanged_label = QLabel("Already Correct: 0")

        summary_layout.addWidget(self.matched_label)
        summary_layout.addWidget(self.duplicate_label)
        summary_layout.addWidget(self.conflict_label)
        summary_layout.addWidget(self.not_found_label)
        summary_layout.addWidget(self.unchanged_label)
        summary_layout.addStretch()

        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Search / Filter:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Filter by filename, status, row, route, or reason...")
        self.search_input.textChanged.connect(self._filter_preview_table)
        filter_layout.addWidget(self.search_input)
        layout.addLayout(filter_layout)

        # Preview table
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(10)
        self.preview_table.setHorizontalHeaderLabels([
            'Status', 'Old Filename', 'New Filename', 'Serial Number', 'Excel Row', 'Direction',
            'Start Point', 'End Point', 'Score', 'Reason'
        ])

        self.preview_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.preview_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.preview_table.setAlternatingRowColors(True)

        layout.addWidget(self.preview_table)

        # Action buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        export_btn = QPushButton("Export Preview")
        export_btn.clicked.connect(self._export_preview)
        export_btn.setMinimumWidth(130)

        rename_btn = QPushButton("Rename Files")
        rename_btn.clicked.connect(self._rename_files)
        rename_btn.setMinimumWidth(130)
        rename_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; }")

        button_layout.addWidget(export_btn)
        button_layout.addWidget(rename_btn)

        layout.addLayout(button_layout)
        widget.setLayout(layout)
        return widget

    def _create_settings_tab(self) -> QWidget:
        """Create settings tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        # Serial format
        serial_group = QGroupBox("Serial Number Format")
        serial_layout = QVBoxLayout()

        self.serial_format_combo = QComboBox()
        self.serial_format_combo.addItem("1 (integer only)", "integer")
        self.serial_format_combo.addItem("1.0 (with decimal)", "decimal")
        self.serial_format_combo.addItem("001 (zero-padded)", "padded")
        self.serial_format_combo.currentIndexChanged.connect(self._on_serial_format_changed)

        serial_layout.addWidget(QLabel("Select format for serial numbers:"))
        serial_layout.addWidget(self.serial_format_combo)
        serial_layout.addStretch()

        serial_group.setLayout(serial_layout)
        layout.addWidget(serial_group)

        # File extensions
        ext_group = QGroupBox("File Extensions")
        ext_layout = QVBoxLayout()

        self.extensions_input = QLineEdit()
        self.extensions_input.setText(".vbo,.vbox,.vts")
        self.extensions_input.setPlaceholderText("Comma-separated: .vbo,.vbox,.vts")

        ext_layout.addWidget(QLabel("Extensions to include:"))
        ext_layout.addWidget(self.extensions_input)

        ext_group.setLayout(ext_layout)
        layout.addWidget(ext_group)

        # Options
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()

        self.subfolders_check = QCheckBox("Include subfolders")
        options_layout.addWidget(self.subfolders_check)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Undo section
        undo_group = QGroupBox("Undo Operations")
        undo_layout = QVBoxLayout()

        undo_btn = QPushButton("Undo Last Rename")
        undo_btn.clicked.connect(self._undo_rename)
        undo_btn.setMinimumWidth(130)

        view_log_btn = QPushButton("View Rename Log")
        view_log_btn.clicked.connect(self._view_log)
        view_log_btn.setMinimumWidth(130)

        undo_layout.addWidget(QLabel("Revert previous rename operations:"))
        undo_btn_layout = QHBoxLayout()
        undo_btn_layout.addWidget(undo_btn)
        undo_btn_layout.addWidget(view_log_btn)
        undo_btn_layout.addStretch()
        undo_layout.addLayout(undo_btn_layout)

        undo_group.setLayout(undo_layout)
        layout.addWidget(undo_group)

        layout.addStretch()
        widget.setLayout(layout)
        return widget

    def _select_folder(self):
        """Select folder containing VBOX files."""
        folder = QFileDialog.getExistingDirectory(self, "Select VBOX Files Folder")
        if folder:
            self.folder_path = folder
            self.folder_input.setText(folder)
            self._update_file_count()
            self.status_label.setText(f"Folder selected: {folder}")

    def _select_excel(self):
        """Select Excel file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", 
            "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )
        if file_path:
            self.excel_path = file_path
            self.excel_input.setText(file_path)

            # Try to read Excel
            success, message = self.excel_reader.read_file(file_path)
            if success:
                self.status_label.setText(message)
                self._update_excel_count()
                self._update_column_options()
                self._auto_detect_columns()
            else:
                QMessageBox.critical(self, "Error", f"Failed to read Excel: {message}")

    def _update_file_count(self):
        """Update file count display."""
        if self.folder_path:
            extensions = self._get_extensions()
            include_subfolders = self.subfolders_check.isChecked()
            files = []
            if include_subfolders:
                for root, _, filenames in os.walk(self.folder_path):
                    for f in filenames:
                        if any(f.lower().endswith(ext.lower()) for ext in extensions):
                            files.append(f)
            else:
                for f in os.listdir(self.folder_path):
                    if os.path.isfile(os.path.join(self.folder_path, f)):
                        if any(f.lower().endswith(ext.lower()) for ext in extensions):
                            files.append(f)
            self.file_count_label.setText(f"Files: {len(files)}")

    def _update_excel_count(self):
        """Update Excel row count display."""
        df = self.excel_reader.get_data()
        self.excel_rows_label.setText(f"Excel rows: {len(df)}")

    def _update_column_options(self):
        """Update column dropdown options."""
        columns = self.excel_reader.get_column_options()
        for combo in self.column_combos.values():
            combo.clear()
            combo.addItem("-- Not Detected --")
            combo.addItems(columns)

    def _auto_detect_columns(self):
        """Auto-detect and set Excel columns."""
        mapping = self.excel_reader._auto_detect_columns()

        for field, combo in self.column_combos.items():
            if field in mapping:
                combo.setCurrentText(mapping[field])
            else:
                combo.setCurrentIndex(0)

        # Also update the mapping
        if mapping:
            self.excel_reader.set_column_mapping(mapping)
            self.status_label.setText(f"Auto-detected columns: {', '.join(mapping.values())}")

    def _on_serial_format_changed(self):
        """Handle serial format change."""
        if self.renamer:
            format_key = self.serial_format_combo.currentData()
            self.renamer.set_serial_format(format_key)

    def _generate_preview(self):
        """Generate preview of rename operations."""
        # Validate inputs
        if not self.folder_path:
            QMessageBox.warning(self, "Warning", "Please select a folder with VBOX files")
            return

        if not self.excel_path:
            QMessageBox.warning(self, "Warning", "Please select an Excel file")
            return

        if not self.excel_reader.is_configured():
            QMessageBox.warning(self, "Warning", "Excel columns are not properly configured")
            return

        # Update mapping from UI
        mapping = {}
        for field, combo in self.column_combos.items():
            text = combo.currentText()
            if text != "-- Not Detected --":
                mapping[field] = text

        if not all(field in mapping for field in ['start_point', 'end_point', 'road_type']):
            QMessageBox.warning(self, "Warning", "Missing required column mappings")
            return

        self.excel_reader.set_column_mapping(mapping)

        # Create renamer
        self.renamer = VBOXRenamer(self.folder_path, self.excel_reader)
        format_key = self.serial_format_combo.currentData()
        self.renamer.set_serial_format(format_key)

        # Get extensions
        extensions = self._get_extensions()

        # Generate preview
        self.progress_bar.setValue(10)
        include_subfolders = self.subfolders_check.isChecked()
        self.preview_data = self.renamer.generate_preview(extensions, include_subfolders)
        self.progress_bar.setValue(80)

        # Display preview
        self._display_preview()

        # Update summary
        summary = self.renamer.get_preview_summary()
        self.matched_label.setText(f"Matched: {summary['matched']}")
        self.duplicate_label.setText(f"Duplicates: {summary['duplicate']}")
        self.conflict_label.setText(f"Conflicts: {summary['conflict']}")
        self.not_found_label.setText(f"Not Found: {summary['not_found']}")
        self.unchanged_label.setText(f"Already Correct: {summary.get('unchanged', 0)}")

        self.status_label.setText(f"Preview generated: {summary['total']} files")
        self.progress_bar.setValue(100)

    def _display_preview(self):
        """Display preview table."""
        self.preview_table.setRowCount(0)

        status_colors = {
            'matched': QColor(144, 238, 144),      # Light green
            'duplicate': QColor(255, 192, 203),    # Light pink
            'conflict': QColor(255, 99, 71),       # Red
            'not_found': QColor(255, 255, 153),    # Light yellow
            'unchanged': QColor(217, 225, 242),    # Light blue
        }

        for row_idx, file_info in enumerate(self.preview_data):
            self.preview_table.insertRow(row_idx)

            # Status
            status_item = QTableWidgetItem(file_info.status.upper())
            status_item.setBackground(status_colors.get(file_info.status, QColor(255, 255, 255)))
            self.preview_table.setItem(row_idx, 0, status_item)

            # Old filename
            self.preview_table.setItem(row_idx, 1, QTableWidgetItem(file_info.filename))

            # New filename
            self.preview_table.setItem(row_idx, 2, QTableWidgetItem(file_info.new_filename))

            # Serial number from matched Excel row
            self.preview_table.setItem(row_idx, 3, QTableWidgetItem(str(file_info.matched_data.get('serial', ''))))

            # Excel row
            source_row = file_info.matched_data.get('source_excel_row', '')
            excel_row_text = str(source_row or file_info.excel_row + 1) if file_info.excel_row >= 0 else "-"
            self.preview_table.setItem(row_idx, 4, QTableWidgetItem(excel_row_text))

            # Direction
            self.preview_table.setItem(row_idx, 5, QTableWidgetItem(str(file_info.matched_data.get('direction', ''))))

            # Start and end from matched Excel row
            self.preview_table.setItem(row_idx, 6, QTableWidgetItem(str(file_info.matched_data.get('start_point', ''))))
            self.preview_table.setItem(row_idx, 7, QTableWidgetItem(str(file_info.matched_data.get('end_point', ''))))

            # Score
            score_text = f"{file_info.score:.0%}" if file_info.score > 0 else "-"
            self.preview_table.setItem(row_idx, 8, QTableWidgetItem(score_text))

            # Reason
            self.preview_table.setItem(row_idx, 9, QTableWidgetItem(file_info.reason))

        self._filter_preview_table()

    def _rename_files(self):
        """Rename files after confirmation."""
        if not self.preview_data:
            QMessageBox.warning(self, "Warning", "No preview generated")
            return

        # Count files to rename
        to_rename = [f for f in self.preview_data if f.status == 'matched']

        if not to_rename:
            QMessageBox.warning(self, "Warning", "No files to rename (no matched files)")
            return

        # Confirm
        reply = QMessageBox.question(
            self, "Confirm Rename",
            f"Are you sure you want to rename {len(to_rename)} file(s)?\n\n"
            "A log file will be created for undo operations.",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.No:
            return

        # Perform rename
        try:
            self.progress_bar.setValue(20)
            successful, summary = self.renamer.rename_files(to_rename)
            self.progress_bar.setValue(100)

            if successful > 0:
                QMessageBox.information(
                    self, "Success",
                    f"Successfully renamed {successful} file(s)\n\n"
                    f"Log files created in folder:\n"
                    f"- rename_log.json\n"
                    f"- rename_log.xlsx"
                )
                self.status_label.setText(f"Renamed {successful} files")
            else:
                QMessageBox.warning(self, "Error", summary)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Rename failed: {str(e)}")

    def _export_preview(self):
        """Export preview to Excel."""
        if not self.preview_data:
            QMessageBox.warning(self, "Warning", "No preview to export")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export Preview", "", "Excel Files (*.xlsx)"
        )

        if file_path:
            try:
                data = []
                for file_info in self.preview_data:
                    data.append({
                        'Status': file_info.status,
                        'Old Filename': file_info.filename,
                        'New Filename': file_info.new_filename,
                        'Serial Number': file_info.matched_data.get('serial', ''),
                        'Excel Row': file_info.matched_data.get('source_excel_row', file_info.excel_row + 1 if file_info.excel_row >= 0 else ''),
                        'Direction': file_info.matched_data.get('direction', ''),
                        'Start Point': file_info.matched_data.get('start_point', ''),
                        'End Point': file_info.matched_data.get('end_point', ''),
                        'Road Type': file_info.matched_data.get('road_type', ''),
                        'Distance': file_info.matched_data.get('distance', ''),
                        'Match Score': f"{file_info.score:.0%}" if file_info.score > 0 else '',
                        'Reason': file_info.reason,
                    })

                df = pd.DataFrame(data)
                df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Success", f"Preview exported to:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed: {str(e)}")

    def _undo_rename(self):
        """Undo last rename operation."""
        if not self.folder_path:
            QMessageBox.warning(self, "Warning", "Please select the folder with VBOX files")
            return

        if not self.renamer:
            self.renamer = VBOXRenamer(self.folder_path, self.excel_reader)

        reply = QMessageBox.question(
            self, "Confirm Undo",
            "This will undo the last rename operation using rename_log.json\n\n"
            "Are you sure?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            success, message = self.renamer.undo_last_rename()

            if success:
                QMessageBox.information(self, "Success", message)
                self.status_label.setText("Undo completed")
            else:
                QMessageBox.warning(self, "Info", message)

    def _view_log(self):
        """View rename log file."""
        if not self.folder_path:
            QMessageBox.warning(self, "Warning", "Please select the folder with VBOX files")
            return

        log_file = os.path.join(self.folder_path, 'rename_log.xlsx')
        if os.path.exists(log_file):
            os.startfile(log_file)  # Windows
        else:
            QMessageBox.information(
                self, "Info",
                "No rename log found. Rename files first to create a log."
            )

    def _get_extensions(self) -> List[str]:
        """Read and normalize extension settings."""
        extensions = [ext.strip() for ext in self.extensions_input.text().split(',') if ext.strip()]
        if not extensions:
            extensions = ['.vbo']
        return [ext if ext.startswith('.') else f'.{ext}' for ext in extensions]

    def _filter_preview_table(self):
        """Filter preview rows based on the search box."""
        if not hasattr(self, 'search_input'):
            return
        query = self.search_input.text().strip().lower()
        for row in range(self.preview_table.rowCount()):
            row_text = []
            for column in range(self.preview_table.columnCount()):
                item = self.preview_table.item(row, column)
                if item:
                    row_text.append(item.text().lower())
            self.preview_table.setRowHidden(row, query not in " ".join(row_text))

    def dragEnterEvent(self, event):
        """Accept dropped folders or Excel files."""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        """Support drag-and-drop selection for folder and Excel path."""
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isdir(path):
                self.folder_path = path
                self.folder_input.setText(path)
                self._update_file_count()
                self.status_label.setText(f"Folder selected: {path}")
            elif path.lower().endswith(('.xlsx', '.xls')):
                self.excel_path = path
                self.excel_input.setText(path)
                success, message = self.excel_reader.read_file(path)
                if success:
                    self.status_label.setText(message)
                    self._update_excel_count()
                    self._update_column_options()
                    self._auto_detect_columns()
                else:
                    QMessageBox.critical(self, "Error", f"Failed to read Excel: {message}")

    def apply_styles(self):
        """Apply application styles."""
        style = """
            QMainWindow {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 6px 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
            QPushButton:pressed {
                background-color: #004578;
            }
            QTableWidget {
                border: 1px solid #cccccc;
                gridline-color: #cccccc;
            }
        """
        self.setStyleSheet(style)


def main():
    """Main entry point."""
    app = QApplication(sys.argv)
    window = VBOXRenamerUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
