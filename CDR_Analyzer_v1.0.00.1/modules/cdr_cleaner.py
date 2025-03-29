import os
import pandas as pd
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QPushButton, QFileDialog, QMessageBox,
                            QProgressBar, QFrame, QStackedWidget)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon, QPixmap

class CDRFileCleaner(QWidget):
    def __init__(self, back_to_home_callback):
        super().__init__()
        self.back_to_home_callback = back_to_home_callback
        self.setStyleSheet("""
            background-color: #1E1E2F;
            color: #FFFFFF;
            font-family: Arial;
        """)
        
        self.create_ui()
        
    def create_ui(self):
        """Create the user interface for the cleaner app"""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        self.setLayout(main_layout)
        
        # Title frame
        title_frame = QFrame()
        title_frame.setStyleSheet("background-color: #2C2C3A; border-radius: 10px;")
        title_layout = QVBoxLayout(title_frame)
        
        # Title label
        title_label = QLabel("CDR File Cleaner")
        title_label.setFont(QFont("Arial", 18, QFont.Bold))
        title_label.setStyleSheet("color: #FFCC08;")
        title_label.setAlignment(Qt.AlignCenter)
        title_layout.addWidget(title_label)
        
        # Description label
        desc_label = QLabel("Clean and standardize CDR Excel files")
        desc_label.setFont(QFont("Arial", 12))
        desc_label.setStyleSheet("color: #AAAAAA;")
        desc_label.setAlignment(Qt.AlignCenter)
        title_layout.addWidget(desc_label)
        
        main_layout.addWidget(title_frame)
        
        # Content frame
        content_frame = QFrame()
        content_frame.setStyleSheet("background-color: #2C2C3A; border-radius: 10px;")
        content_layout = QVBoxLayout(content_frame)
        content_layout.setContentsMargins(20, 20, 20, 20)
        
        # Instruction label
        instruction_label = QLabel(
            "This tool will:\n"
            "1. Remove newlines from date/time fields\n"
            "2. Standardize datetime format\n"
            "3. Generate cleaned Excel and CSV versions"
        )
        instruction_label.setFont(QFont("Arial", 11))
        instruction_label.setWordWrap(True)
        content_layout.addWidget(instruction_label)
        
        # Select file button
        self.select_btn = QPushButton("📂 Select CDR File")
        self.select_btn.setFont(QFont("Arial", 12, QFont.Bold))
        self.select_btn.setStyleSheet("""
            QPushButton {
                background-color: #27AE60;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                padding: 12px;
                min-width: 200px;
            }
            QPushButton:hover {
                background-color: #219653;
            }
        """)
        self.select_btn.clicked.connect(self.select_file)
        content_layout.addWidget(self.select_btn, alignment=Qt.AlignCenter)
        
        # Status label
        self.status_label = QLabel("Ready to process files")
        self.status_label.setFont(QFont("Arial", 10))
        self.status_label.setStyleSheet("color: #AAAAAA;")
        self.status_label.setAlignment(Qt.AlignCenter)
        content_layout.addWidget(self.status_label)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #444;
                border-radius: 4px;
                background: #2C2C3A;
            }
            QProgressBar::chunk {
                background-color: #27AE60;
            }
        """)
        self.progress_bar.hide()
        content_layout.addWidget(self.progress_bar)
        
        main_layout.addWidget(content_frame)
        
        # Back button
        back_btn = QPushButton("⬅ Back to Home")
        back_btn.setFont(QFont("Arial", 12))
        back_btn.setStyleSheet("""
            QPushButton {
                background-color: #E74C3C;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #C0392B;
            }
        """)
        back_btn.clicked.connect(self.back_to_home_callback)
        main_layout.addWidget(back_btn)
        
        main_layout.addStretch()
        
    def clean_cdr_file(self, input_file):
        """Clean the CDR file and save it in the same folder"""
        try:
            self.status_label.setText("Processing file...")
            self.status_label.setStyleSheet("color: #FFCC08;")
            self.progress_bar.show()
            self.progress_bar.setValue(10)
            QApplication.processEvents()
            
            # Read Excel file
            self.progress_bar.setValue(30)
            xls = pd.ExcelFile(input_file)
            sheet_name = xls.sheet_names[0]
            df = pd.read_excel(xls, sheet_name=sheet_name)

            # Clean the event_date_time column
            self.progress_bar.setValue(50)
            df["event_date_time"] = df["event_date_time"].astype(str).str.replace("\n", "").str.strip()
            df["event_date_time"] = pd.to_datetime(
                df["event_date_time"], 
                format="%Y-%m-%d %H:%M:%S", 
                errors="coerce"
            )

            # Prepare output paths
            self.progress_bar.setValue(70)
            base_folder = os.path.dirname(input_file)
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            output_excel = os.path.join(base_folder, f"{base_name}_cleaned.xlsx")
            output_csv = os.path.join(base_folder, f"{base_name}_cleaned.csv")

            # Save cleaned files
            self.progress_bar.setValue(90)
            df.to_excel(output_excel, index=False)
            df.to_csv(output_csv, index=False)

            self.progress_bar.setValue(100)
            return True, f"Files successfully saved to:\n{output_excel}\n{output_csv}"
            
        except Exception as e:
            return False, f"Error processing file:\n{str(e)}"
        finally:
            self.progress_bar.hide()

    def select_file(self):
        """Open file dialog to select an Excel file and process it"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select CDR File",
            "",
            "Excel Files (*.xlsx);;All Files (*)"
        )
        
        if file_path:
            success, message = self.clean_cdr_file(file_path)
            if success:
                QMessageBox.information(
                    self, 
                    "Success", 
                    message,
                    QMessageBox.Ok
                )
                self.status_label.setText("File cleaned successfully!")
                self.status_label.setStyleSheet("color: #27AE60;")
            else:
                QMessageBox.critical(
                    self, 
                    "Error", 
                    message,
                    QMessageBox.Ok
                )
                self.status_label.setText("Error cleaning file!")
                self.status_label.setStyleSheet("color: #E74C3C;")