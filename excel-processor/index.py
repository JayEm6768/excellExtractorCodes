import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QProgressBar, QMessageBox,
                             QListWidget, QSplitter, QTextEdit, QGroupBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont


columns_to_extract = [
    'DPdeniro', 'S_SP', 'S_Total', 'Com Date',
    'DP/NAP LAT', 'DP/NAP LONG', 'BRGY_NAME',
    'CFS Cluster', 'Tech', 'Location Type'
]

valid_clusters = ["DAVAO NORTH", "DAVAO SOUTH", "TAGUM 1", "TAGUM 2"]

group1_brgy = [
    'Agdao', 'Bago Gallera', 'Baliok', 'Bangkas Heights', 'Barangay 1-A', 
    'Barangay 2-A', 'Barangay 3-A', 'Barangay 4-A', 'Barangay 5-A', 'Barangay 6-A', 
    'Barangay 7-A', 'Barangay 8-A', 'Barangay 9-A', 'Barangay 10-A', 'Barangay 11-B', 
    'Barangay 12-B', 'Barangay 13-B', 'Barangay 14-B', 'Barangay 15-B', 'Barangay 16-B', 
    'Barangay 17-B', 'Barangay 18-B', 'Barangay 19-B', 'Barangay 20-B', 'Barangay 21-C', 
    'Barangay 22-C', 'Barangay 23-C', 'Barangay 24-C', 'Barangay 26-C', 'Barangay 27-C', 
    'Barangay 28-C', 'Barangay 29-C', 'Barangay 30-C', 'Barangay 31-D', 'Barangay 32-D', 
    'Barangay 33-D', 'Barangay 34-D', 'Barangay 35-D', 'Barangay 36-D', 'Barangay 37-D', 
    'Barangay 38-D', 'Barangay 39-D', 'Barangay 40-D', 'Bucana', 'Centro', 
    'Gov. Vicente Duterte', 'Gov. Paciano Bangoy', 'Lapu-lapu', 'Leon Garcia Sr.', 
    'San Antonio', 'Tres De Mayo', 'Zone 1',
    'Matina Crossing', 'Kap. Tomas Monteverde Sr.'
]

group2_brgy = [
    'Rafael Castillo', 'Sasa', 'Vicente Hizon Sr.', 
    'Ubalde', 'Wilfredo Aquino', 'Pampanga',
    'Buhangin', 'Alfonso Angliongto Sr.'
]

group3_brgy = [
    'Cabantian', 'Mandug', 'Panacan', 'Bunawan', 'Indangan', 
    'Alejandra Navarro', 'Tagpore',
    'Tibungco', 'Communal', 'San Isidro', 'Acacia','Tigatto','Ilang'
]

group_mapping = {
    "South": group1_brgy,
    "Central": group2_brgy,
    "North": group3_brgy
}


class ExcelProcessor(QThread):
    """Thread for processing Excel files to prevent UI freezing"""
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    processing_finished = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)
    debug_info = pyqtSignal(str)

    def __init__(self, input_filepath, output_dir):
        super().__init__()
        self.input_filepath = input_filepath
        self.output_dir = output_dir

    def clean_barangay_name(self, name):
        """Clean barangay name by removing extra text like (POB.) and trimming whitespace"""
        if pd.isna(name):
            return name
        name = str(name).strip()
        # Remove (POB.) and similar suffixes
        name = name.split('(')[0].strip()
        # Remove any trailing special characters
        name = name.rstrip('.) ')
        return name

    def match_barangay(self, actual_name, target_names):
        """Flexible matching of barangay names"""
        if pd.isna(actual_name):
            return False
        
        actual_clean = self.clean_barangay_name(actual_name)
        
        for target_name in target_names:
            target_clean = self.clean_barangay_name(target_name)
            # Check if the cleaned names match (case insensitive)
            if actual_clean.lower() == target_clean.lower():
                return True
            # Also check if the target is contained in the actual name
            if target_clean.lower() in actual_clean.lower():
                return True
        return False

    def run(self):
        try:
            self.status_updated.emit("Loading Excel file...")
            self.progress_updated.emit(10)
            
            # Load the Excel file
            df = pd.read_excel(self.input_filepath)
            
            # DEBUG: Show column names and first few rows
            self.debug_info.emit(f"Original DataFrame shape: {df.shape}")
            self.debug_info.emit(f"Columns in file: {list(df.columns)}")
            
            self.status_updated.emit("Filtering data...")
            self.progress_updated.emit(30)
            
            # Check if all required columns exist
            missing_columns = [col for col in columns_to_extract if col not in df.columns]
            if missing_columns:
                self.debug_info.emit(f"WARNING: Missing columns: {missing_columns}")
                # Try to find similar column names
                for missing_col in missing_columns:
                    similar_cols = [col for col in df.columns if missing_col.lower() in col.lower()]
                    if similar_cols:
                        self.debug_info.emit(f"  Similar columns found for '{missing_col}': {similar_cols}")
            
            # Filter columns and clusters
            df = df[columns_to_extract]
            
            # Clean barangay names in the dataframe
            df['BRGY_NAME_CLEAN'] = df['BRGY_NAME'].apply(self.clean_barangay_name)
            self.debug_info.emit(f"After column filtering shape: {df.shape}")
            
            # Check cluster values
            cluster_values = df['CFS Cluster'].unique() if 'CFS Cluster' in df.columns else []
            self.debug_info.emit(f"Unique CFS Cluster values: {cluster_values}")
            
            df = df[df['CFS Cluster'].isin(valid_clusters)]
            self.debug_info.emit(f"After cluster filtering shape: {df.shape}")
            
            # Check barangay values
            barangay_values = df['BRGY_NAME'].unique() if 'BRGY_NAME' in df.columns else []
            self.debug_info.emit(f"Unique BRGY_NAME values: {barangay_values[:20]}")  # First 20 only
            
            # Check cleaned barangay values
            cleaned_barangay_values = df['BRGY_NAME_CLEAN'].unique() if 'BRGY_NAME_CLEAN' in df.columns else []
            self.debug_info.emit(f"Unique cleaned BRGY_NAME values: {cleaned_barangay_values[:20]}")
            
            # Check tech values
            tech_values = df['Tech'].unique() if 'Tech' in df.columns else []
            self.debug_info.emit(f"Unique Tech values: {tech_values}")
            
            self.status_updated.emit("Processing barangay groups...")
            self.progress_updated.emit(50)
            
            # Process barangay groups
            output_files = {}
            
            for name, brgys in group_mapping.items():
                self.debug_info.emit(f"\nProcessing {name} group...")
                
                # Filter by barangay using flexible matching
                filtered = df[df['BRGY_NAME_CLEAN'].apply(lambda x: self.match_barangay(x, brgys))]
                self.debug_info.emit(f"After barangay filtering: {filtered.shape[0]} rows")
                
                # Apply cluster-specific filtering
                if name == "South":
                    # For South group, only include Davao North entries
                    filtered = filtered[filtered['CFS Cluster'] == "DAVAO NORTH"]
                    self.debug_info.emit(f"After DAVAO NORTH filter: {filtered.shape[0]} rows")
                elif name in ["Central", "North"]:
                    # For Central and North, exclude Davao South
                    filtered = filtered[filtered['CFS Cluster'] != "DAVAO SOUTH"]
                    self.debug_info.emit(f"After excluding DAVAO SOUTH: {filtered.shape[0]} rows")
                
                # Remove the temporary clean column before saving
                filtered_to_save = filtered.drop(columns=['BRGY_NAME_CLEAN'], errors='ignore')
                
                # Create main file even if empty to ensure all files are generated
                main_filename = f"{name}.xlsx"
                main_filepath = os.path.join(self.output_dir, main_filename)
                filtered_to_save.to_excel(main_filepath, index=False)
                output_files[main_filename] = main_filepath
                self.debug_info.emit(f"Created {main_filename} with {filtered.shape[0]} rows")
                
                # Create spare file
                spare_data = filtered.copy()
                spare_data['Tech'] = spare_data['Tech'].fillna("GPON")
                spare_data['Tech'] = spare_data['Tech'].replace(" ", "GPON")
                spare_data = spare_data[~spare_data['Tech'].isin(["VDSL", "ADSL", "ADSL/VDSL"])]
                spare_data = spare_data.drop(columns=['BRGY_NAME_CLEAN'], errors='ignore')
                
                spare_filename = f"{name} Spare.xlsx"
                spare_filepath = os.path.join(self.output_dir, spare_filename)
                spare_data.to_excel(spare_filepath, index=False)
                output_files[spare_filename] = spare_filepath
                self.debug_info.emit(f"Created {spare_filename} with {spare_data.shape[0]} rows")
            
            self.status_updated.emit("Creating DSL file...")
            self.progress_updated.emit(70)
            
            # Create DSL file - only from Davao North cluster with DSL technologies
            dsl_data = df[df['Tech'].isin(["VDSL", 'ADSL', 'ADSL/VDSL'])]
            self.debug_info.emit(f"DSL technologies found: {dsl_data.shape[0]} rows")
            
            # Apply the Davao North constraint
            dsl_data = dsl_data[dsl_data['CFS Cluster'] == "DAVAO NORTH"]
            self.debug_info.emit(f"DSL from DAVAO NORTH: {dsl_data.shape[0]} rows")
            
            # Remove the temporary clean column before saving
            dsl_data = dsl_data.drop(columns=['BRGY_NAME_CLEAN'], errors='ignore')
            
            # Create DSL file even if empty
            dsl_filepath = os.path.join(self.output_dir, "DSL.xlsx")
            dsl_data.to_excel(dsl_filepath, index=False)
            output_files["DSL.xlsx"] = dsl_filepath
            self.debug_info.emit(f"Created DSL.xlsx with {dsl_data.shape[0]} rows")
            
            self.status_updated.emit("Adding coordinates...")
            self.progress_updated.emit(90)
            
            # Add coordinates to all files
            for filename in os.listdir(self.output_dir):
                if filename.endswith(".xlsx"):
                    filepath = os.path.join(self.output_dir, filename)
                    data = pd.read_excel(filepath)
                    if 'DP/NAP LAT' in data.columns and 'DP/NAP LONG' in data.columns:
                        # Handle NaN values in coordinates
                        data['DP/NAP LAT'] = data['DP/NAP LAT'].fillna('')
                        data['DP/NAP LONG'] = data['DP/NAP LONG'].fillna('')
                        data['coordinates'] = data['DP/NAP LAT'].astype(str) + ", " + data['DP/NAP LONG'].astype(str)
                    data.to_excel(filepath, index=False)
            
            self.progress_updated.emit(100)
            self.status_updated.emit("Processing complete!")
            self.debug_info.emit(f"Final output: {len(output_files)} files created")
            self.processing_finished.emit(output_files)
            
        except Exception as e:
            self.error_occurred.emit(str(e))
            import traceback
            self.debug_info.emit(f"Error details: {traceback.format_exc()}")


class ExcelProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_filepath = None
        self.output_dir = os.path.join(os.path.expanduser("~"), "ExcelProcessorOutput")
        os.makedirs(self.output_dir, exist_ok=True)
        self.processor = None
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("Excel File Processor")
        self.setGeometry(100, 100, 900, 700)
        
        # Set application style
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f7fa;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 8px;
                margin-top: 1ex;
                padding-top: 10px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #2c3e50;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
            QPushButton#browseButton {
                background-color: #2ecc71;
            }
            QPushButton#browseButton:hover {
                background-color: #27ae60;
            }
            QPushButton#openFolderButton {
                background-color: #9b59b6;
            }
            QPushButton#openFolderButton:hover {
                background-color: #8e44ad;
            }
            QListWidget {
                border: 1px solid #cccccc;
                border-radius: 5px;
                background-color: white;
            }
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 5px;
                text-align: center;
                background-color: #ecf0f1;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                width: 10px;
            }
        """)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Title
        title = QLabel("Excel File Processor")
        title.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(20)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setStyleSheet("color: #2c3e50; padding: 15px;")
        layout.addWidget(title)
        
        # Description
        description = QLabel(
            "This tool processes Excel files by filtering data based on CFS Cluster and Barangay, "
            "then categorizes the data into regional groups and technology types. "
            "South file only includes entries from Davao North CFS cluster."
        )
        description.setWordWrap(True)
        description.setAlignment(Qt.AlignCenter)
        description.setStyleSheet("color: #7f8c8d; padding: 0 0 15px 0;")
        layout.addWidget(description)
        
        # Requirements box
        requirements = QLabel(
            "Requirements:\n"
            "- All files: Filtered by valid CFS clusters (Davao North, Davao South, Tagum 1, Tagum 2)\n"
            "- South file: Only includes Davao North entries from Group 1 barangays\n"
            "- Central/North files: Include Tagum 1 and Tagum 2 clusters (exclude Davao South)\n"
            "- DSL file: Only includes VDSL/ADSL technologies from Davao North cluster"
        )
        requirements.setWordWrap(True)
        requirements.setStyleSheet("background-color: #e8f4fd; padding: 10px; border-radius: 5px;")
        layout.addWidget(requirements)
        
        # Splitter for main content
        splitter = QSplitter(Qt.Vertical)
        layout.addWidget(splitter)
        
        # File selection group
        file_group = QGroupBox("File Selection")
        file_layout = QVBoxLayout(file_group)
        
        file_btn_layout = QHBoxLayout()
        self.select_file_btn = QPushButton("Select Excel File")
        self.select_file_btn.setObjectName("browseButton")
        self.select_file_btn.clicked.connect(self.select_file)
        file_btn_layout.addWidget(self.select_file_btn)
        
        self.file_label = QLabel("No file selected")
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet("padding: 5px; background-color: #e8f4fd; border-radius: 5px;")
        file_btn_layout.addWidget(self.file_label, 1)
        
        file_layout.addLayout(file_btn_layout)
        splitter.addWidget(file_group)
        
        # Progress group
        progress_group = QGroupBox("Processing")
        progress_layout = QVBoxLayout(progress_group)
        
        self.process_btn = QPushButton("Process File")
        self.process_btn.clicked.connect(self.process_file)
        self.process_btn.setEnabled(False)
        progress_layout.addWidget(self.process_btn)
        
        self.status_label = QLabel("Ready to process")
        self.status_label.setStyleSheet("padding: 5px; background-color: #f0f0f0; border-radius: 5px;")
        progress_layout.addWidget(self.status_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        splitter.addWidget(progress_group)
        
        # Output group
        output_group = QGroupBox("Output Files")
        output_layout = QVBoxLayout(output_group)
        
        self.output_list = QListWidget()
        output_layout.addWidget(self.output_list)
        
        open_folder_btn = QPushButton("Open Output Folder")
        open_folder_btn.setObjectName("openFolderButton")
        open_folder_btn.clicked.connect(self.open_output_folder)
        output_layout.addWidget(open_folder_btn)
        
        splitter.addWidget(output_group)
        
        # Set splitter sizes
        splitter.setSizes([150, 150, 300])
        
        # Log area
        log_group = QGroupBox("Processing Log")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_group)
        
        self.log("Application started")
        self.log(f"Output directory: {self.output_dir}")
        
    def log(self, message):
        """Add a message to the log"""
        self.log_text.append(f"{pd.Timestamp.now().strftime('%H:%M:%S')} - {message}")
        
    def select_file(self):
        """Open a file dialog to select an Excel file"""
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        
        if filepath:
            self.input_filepath = filepath
            self.file_label.setText(filepath)
            self.process_btn.setEnabled(True)
            self.log(f"Selected file: {filepath}")
            
    def process_file(self):
        """Start processing the selected file"""
        if not self.input_filepath:
            QMessageBox.warning(self, "Warning", "Please select an Excel file first.")
            return
            
        # Disable the process button during processing
        self.process_btn.setEnabled(False)
        self.select_file_btn.setEnabled(False)
        self.output_list.clear()
        self.progress_bar.setValue(0)
        self.status_label.setText("Processing...")
        
        # Create and start the processor thread
        self.processor = ExcelProcessor(self.input_filepath, self.output_dir)
        self.processor.progress_updated.connect(self.update_progress)
        self.processor.status_updated.connect(self.update_status)
        self.processor.processing_finished.connect(self.processing_finished)
        self.processor.error_occurred.connect(self.processing_error)
        self.processor.debug_info.connect(self.log)
        self.processor.start()
        
        self.log("Started processing file")
        
    def update_progress(self, value):
        """Update the progress bar"""
        self.progress_bar.setValue(value)
        
    def update_status(self, status):
        """Update the status label"""
        self.status_label.setText(status)
        self.log(status)
        
    def processing_finished(self, output_files):
        """Handle completion of processing"""
        self.process_btn.setEnabled(True)
        self.select_file_btn.setEnabled(True)
        self.status_label.setText("Processing complete!")
        
        # Add output files to the list
        for filename in output_files:
            self.output_list.addItem(filename)
            
        self.log(f"Processing complete. Generated {len(output_files)} files.")
        QMessageBox.information(self, "Success", 
                               f"Processing complete! Generated {len(output_files)} files.")
        
    def processing_error(self, error_message):
        """Handle processing errors"""
        self.process_btn.setEnabled(True)
        self.select_file_btn.setEnabled(True)
        self.status_label.setText("Error occurred!")
        
        self.log(f"Error: {error_message}")
        QMessageBox.critical(self, "Error", f"An error occurred during processing:\n{error_message}")
        
    def open_output_folder(self):
        """Open the output folder in the system file explorer"""
        os.startfile(self.output_dir) if os.name == 'nt' else \
        os.system(f'open "{self.output_dir}"') if os.name == 'posix' and sys.platform == 'darwin' else \
        os.system(f'xdg-open "{self.output_dir}"')
        self.log("Opened output folder")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle("Fusion")
    
    window = ExcelProcessorApp()
    window.show()
    
    sys.exit(app.exec_())