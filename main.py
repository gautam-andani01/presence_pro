import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
from datetime import datetime
from pathlib import Path
from processsing import transformation

class ModernApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # Set window properties
        self.setWindowTitle('Modern File Processor')
        self.setGeometry(100, 100, 500, 400)
        
        # Set overall layout
        self.layout = QVBoxLayout()

        # Create labels and buttons with modern style
        self.label_biometric = QLabel('Upload Biometric File', self)
        self.label_biometric.setAlignment(Qt.AlignCenter)
        self.label_biometric.setFont(QFont('Arial', 10))
        
        self.button_biometric = QPushButton('Upload Biometric File', self)
        self.button_biometric.setObjectName('uploadButton')
        self.button_biometric.clicked.connect(self.upload_biometric)

        self.label_keka = QLabel('Upload Keka File', self)
        self.label_keka.setAlignment(Qt.AlignCenter)
        self.label_keka.setFont(QFont('Arial', 10))

        self.button_keka = QPushButton('Upload Keka File', self)
        self.button_keka.setObjectName('uploadButton')
        self.button_keka.clicked.connect(self.upload_keka)

        self.button_process = QPushButton('Process Files', self)
        self.button_process.setObjectName('processButton')
        self.button_process.clicked.connect(self.process_files)
        
        self.button_download = QPushButton('Download Processed File')
        self.button_download.setObjectName('downloadButton')
        self.button_download.setEnabled(False)
        self.button_download.clicked.connect(self.download_file)

        self.label_status = QLabel('', self)
        self.label_status.setAlignment(Qt.AlignCenter)
        self.label_status.setFont(QFont('Arial', 10))

        # Add widgets to the layout
        self.layout.addWidget(self.label_biometric)
        self.layout.addWidget(self.button_biometric)
        self.layout.addWidget(self.label_keka)
        self.layout.addWidget(self.button_keka)
        self.layout.addWidget(self.button_process)
        self.layout.addWidget(self.button_download)
        self.layout.addWidget(self.label_status)
        
        self.setLayout(self.layout)

        # Variables to hold file paths
        self.biometric_file = None
        self.keka_file = None
        self.processed_file = None

        # Apply custom styles (QSS)
        self.apply_styles()

    def apply_styles(self):
        # Add a modern flat look with QSS (CSS-like styles)
        self.setStyleSheet("""
            QWidget {
                background-color: #2C2F33;
                color: #FFFFFF;
            }
            QLabel {
                color: #99AAB5;
                font-size: 14px;
            }
            QPushButton {
                font-size: 14px;
                padding: 10px;
                border-radius: 5px;
                border: 1px solid #7289DA;
                background-color: #5865F2;
                color: #FFFFFF;
            }
            QPushButton#uploadButton:hover {
                background-color: #7289DA;
            }
            QPushButton#processButton:hover {
                background-color: #43B581;
            }
            QPushButton#downloadButton:disabled {
                background-color: #99AAB5;
                color: #23272A;
            }
            QPushButton#downloadButton:enabled:hover {
                background-color: #43B581;
            }
        """)

    def upload_biometric(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Biometric File", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.biometric_file = file_name
            self.label_biometric.setText(f"Selected: {file_name.split('/')[-1]}")

    def upload_keka(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Keka File", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.keka_file = file_name
            self.label_keka.setText(f"Selected: {file_name.split('/')[-1]}")

    def process_files(self):
        if not self.biometric_file:
            QMessageBox.warning(self, "Warning", "Please select the biometric file!")
            return
        if not self.keka_file:
            QMessageBox.warning(self, "Warning", "Please select the keka file!")
            return
        
        try:
            # Process only the biometric file for now
            # self.processed_file = self.process_biometric(self.biometric_file)
            self.processed_file = transformation(biometric_path=self.biometric_file, keka_path=self.keka_file)
            self.label_status.setText(f"Biometric file processed successfully.")
            
            # Enable the download button
            self.button_download.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred during processing:\n{str(e)}")

    def process_biometric(self, biometric_file):
        # Just read the biometric file and save it back for now
        biometric_df = pd.read_excel(biometric_file)
        
        # Perform any transformations here if needed. For now, we just save it as-is.
        processed_file_path = "temp_processed_biometric_output.xlsx"
        biometric_df.to_excel(processed_file_path, index=False)
        return processed_file_path

    def download_file(self):
        try:
            if not self.processed_file:
                QMessageBox.warning(self, "Warning", "No processed file found. Please process the files first.")
                return

            # Get the user's Downloads directory
            downloads_dir = str(Path.home() / "Downloads")

            # Create a timestamped filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name = f"processed_biometric_{timestamp}.xlsx"
            output_path = os.path.join(downloads_dir, file_name)

            # Save the processed file to the Downloads folder
            os.rename(self.processed_file, output_path)

            # Notify user about the download
            QMessageBox.information(self, "Download Complete", f"File downloaded to: {output_path}")
            self.label_status.setText(f"File downloaded to: {output_path}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while downloading:\n{str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ModernApp()
    window.show()
    sys.exit(app.exec_())
