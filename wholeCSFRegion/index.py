import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QFileDialog, QProgressBar, QMessageBox, QListWidget, QSplitter, QGroupBox, QComboBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
import math
import subprocess

# Columns to extract
columns_to_extract = [
    'DPdeniro', 'S_SP', 'S_Total', 'Com Date',
    'DP/NAP LAT', 'DP/NAP LONG', 'BRGY_NAME',
    'CFS Cluster', 'Tech', 'Location Type'
]

# All valid clusters
valid_clusters = [
    "AGUSAN", "BUKIDNON", "CAGAYAN EAST", "CAGAYAN WEST",
    "COTABATO", "DAVAO NORTH", "DAVAO SOUTH", "DIGOS", "GENSAN",
    "KORONADAL", "LANAO", "MARAWI", "OZAMIS", "SURIGAO",
    "TAGUM 1", "TAGUM 2", "ZAMBOANGA CITY", "ZAMBOANGA DEL NORTE",
    "ZAMBOANGA DEL SUR", "ZAMBOANGA SIBUGAY"
]


class ExcelProcessor(QThread):
    """Thread for processing Excel files"""
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    processing_finished = pyqtSignal(list)
    error_occurred = pyqtSignal(str)

    def __init__(self, input_filepath, output_dir, selected_cluster):
        super().__init__()
        self.input_filepath = input_filepath
        self.output_dir = output_dir
        self.selected_cluster = selected_cluster

    def run(self):
        try:
            self.status_updated.emit("Loading Excel file...")
            df = pd.read_excel(self.input_filepath)

            if self.selected_cluster not in df['CFS Cluster'].unique():
                self.error_occurred.emit(f"Cluster '{self.selected_cluster}' not found in file.")
                return

            self.status_updated.emit(f"Filtering for {self.selected_cluster}...")
            df = df[df['CFS Cluster'] == self.selected_cluster]

            df = df[columns_to_extract].copy()

            # Add coordinates
            self.status_updated.emit("Adding coordinates...")
            df['coordinates'] = df['DP/NAP LAT'].astype(str) + ", " + df['DP/NAP LONG'].astype(str)

            output_files = []

            # Save compiled file (full filtered dataset)
            compiled_filename = f"{self.selected_cluster}_compiled.xlsx"
            compiled_filepath = os.path.join(self.output_dir, compiled_filename)
            df.to_excel(compiled_filepath, index=False)
            output_files.append(compiled_filepath)

            # Split into chunks of 2000
            num_chunks = math.ceil(len(df) / 2000)
            for i in range(num_chunks):
                chunk = df.iloc[i*2000:(i+1)*2000]
                filename = f"{self.selected_cluster}_part{i+1}.xlsx"
                filepath = os.path.join(self.output_dir, filename)
                chunk.to_excel(filepath, index=False)
                output_files.append(filepath)

            self.progress_updated.emit(100)
            self.status_updated.emit("Processing complete!")
            self.processing_finished.emit(output_files)

        except Exception as e:
            self.error_occurred.emit(str(e))


class ExcelProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_filepath = None
        self.output_dir = os.path.join(os.path.expanduser("~"), "ExcelProcessorOutput")
        os.makedirs(self.output_dir, exist_ok=True)
        self.processor = None
        self.selected_cluster = None
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("Excel Cluster Processor")
        self.setGeometry(200, 150, 800, 600)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        title = QLabel("Excel File Processor")
        title.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        splitter = QSplitter(Qt.Vertical)
        layout.addWidget(splitter)

        # File selection
        file_group = QGroupBox("File Selection")
        file_layout = QVBoxLayout(file_group)
        btn_layout = QHBoxLayout()
        self.select_file_btn = QPushButton("Select Excel File")
        self.select_file_btn.clicked.connect(self.select_file)
        btn_layout.addWidget(self.select_file_btn)

        self.file_label = QLabel("No file selected")
        btn_layout.addWidget(self.file_label)
        file_layout.addLayout(btn_layout)

        # ⚠ Warning note
        warning_label = QLabel("⚠ Please close all related Excel files before processing.\n"
                               "This ensures smooth performance and faster completion.")
        warning_font = QFont()
        warning_font.setPointSize(10)
        warning_font.setBold(True)
        warning_label.setFont(warning_font)
        warning_label.setStyleSheet("color: red;")
        file_layout.addWidget(warning_label)

        # Cluster dropdown
        self.cluster_dropdown = QComboBox()
        self.cluster_dropdown.addItems(valid_clusters)
        file_layout.addWidget(QLabel("Select CFS Cluster:"))
        file_layout.addWidget(self.cluster_dropdown)

        splitter.addWidget(file_group)

        # Processing group
        progress_group = QGroupBox("Processing")
        progress_layout = QVBoxLayout(progress_group)

        self.process_btn = QPushButton("Process File")
        self.process_btn.setEnabled(False)
        self.process_btn.clicked.connect(self.process_file)
        progress_layout.addWidget(self.process_btn)

        self.status_label = QLabel("Ready")
        progress_layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        progress_layout.addWidget(self.progress_bar)

        splitter.addWidget(progress_group)

        # Output group
        output_group = QGroupBox("Output Files")
        output_layout = QVBoxLayout(output_group)
        self.output_list = QListWidget()
        output_layout.addWidget(self.output_list)

        # Button to open output folder manually
        self.open_folder_btn = QPushButton("Open Output Folder")
        self.open_folder_btn.setEnabled(False)
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        output_layout.addWidget(self.open_folder_btn)

        splitter.addWidget(output_group)

    def select_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if filepath:
            self.input_filepath = filepath
            self.file_label.setText(filepath)
            self.process_btn.setEnabled(True)

    def process_file(self):
        if not self.input_filepath:
            QMessageBox.warning(self, "Warning", "Please select an Excel file first.")
            return

        self.selected_cluster = self.cluster_dropdown.currentText()
        self.output_list.clear()
        self.status_label.setText("Processing...")
        self.progress_bar.setValue(0)

        # 🔒 Disable process button while running
        self.process_btn.setEnabled(False)

        self.processor = ExcelProcessor(self.input_filepath, self.output_dir, self.selected_cluster)
        self.processor.progress_updated.connect(self.progress_bar.setValue)
        self.processor.status_updated.connect(self.status_label.setText)
        self.processor.processing_finished.connect(self.show_results)
        self.processor.error_occurred.connect(self.show_error)
        self.processor.start()

    def show_results(self, files):
        self.status_label.setText("Done")
        for f in files:
            self.output_list.addItem(os.path.basename(f))
        self.open_folder_btn.setEnabled(True)

        # ✅ Re-enable process button
        self.process_btn.setEnabled(True)

        # Auto open the folder after processing
        self.open_output_folder()

        QMessageBox.information(self, "Success", f"Processing complete! {len(files)} file(s) generated.")

    def open_output_folder(self):
        if os.name == 'nt':  # Windows
            os.startfile(self.output_dir)
        elif sys.platform == "darwin":  # macOS
            subprocess.Popen(["open", self.output_dir])
        else:  # Linux
            subprocess.Popen(["xdg-open", self.output_dir])

    def show_error(self, msg):
        self.status_label.setText("Error")
        QMessageBox.critical(self, "Error", msg)

        # ✅ Re-enable process button after error
        self.process_btn.setEnabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessorApp()
    window.show()
    sys.exit(app.exec_())
