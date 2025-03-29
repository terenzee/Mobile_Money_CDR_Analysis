import os
import sys
import logging
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFrame, QStackedWidget, QMessageBox,
    QToolButton, QFileDialog, QSizePolicy, QProgressBar,
    QStatusBar, QSystemTrayIcon, QMenu, QAction, QListWidget
)
from PyQt5.QtGui import QPixmap, QIcon, QColor, QLinearGradient, QPainter, QBrush, QPen
from PyQt5.QtCore import Qt, QPropertyAnimation, QEasingCurve, QTimer, QSize, QPoint, QRect, pyqtProperty

# Constants
SIDEBAR_COLOR = "#2C2C3A"
BACKGROUND_COLOR = "#1E1E2F"
TEXT_COLOR = "#FFFFFF"
BUTTON_COLOR = "#5e7e95"
HIGHLIGHT_COLOR = "#6d9dbd"
EXIT_BUTTON_COLOR = "#ed1c24"
TELECEL_BUTTON_COLOR = "#FF6B00"
SUCCESS_COLOR = "#2ECC71"
ERROR_COLOR = "#E74C3C"
WARNING_COLOR = "#F39C12"

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('unified_analyzer.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class GradientButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFixedHeight(40)
        self.setCursor(Qt.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        # Gradient colors
        self.color1 = QColor("#5e7e95")
        self.color2 = QColor("#3a5a78")
        self.hover_color1 = QColor("#6d9dbd")
        self.hover_color2 = QColor("#4a7a9d")
        
        # Animation system
        self._animation_progress = 0  # 0-100
        self.animation = QPropertyAnimation(self, b"animation_progress")
        self.animation.setDuration(150)
        self.animation.setEasingCurve(QEasingCurve.OutQuad)
        
    def get_animation_progress(self):
        return self._animation_progress
        
    def set_animation_progress(self, value):
        self._animation_progress = value
        self.update()
        
    animation_progress = pyqtProperty(int, get_animation_progress, set_animation_progress)
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Interpolate colors based on animation progress
        progress = self._animation_progress / 100.0
        color1 = self.blend_colors(self.color1, self.hover_color1, progress)
        color2 = self.blend_colors(self.color2, self.hover_color2, progress)
        
        # Create gradient
        gradient = QLinearGradient(0, 0, self.width(), self.height())
        gradient.setColorAt(0, color1)
        gradient.setColorAt(1, color2)
        
        # Draw button
        painter.setBrush(QBrush(gradient))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(0, 0, self.width(), self.height(), 5, 5)
        
        # Draw text with slight offset during animation
        text_offset = QPoint(0, int(1 * progress))
        painter.setPen(QPen(Qt.white))
        painter.drawText(self.rect().translated(text_offset), Qt.AlignCenter, self.text())
    
    def blend_colors(self, color1, color2, ratio):
        return QColor(
            int(color1.red() + (color2.red() - color1.red()) * ratio),
            int(color1.green() + (color2.green() - color1.green()) * ratio),
            int(color1.blue() + (color2.blue() - color1.blue()) * ratio),
        )
    
    def enterEvent(self, event):
        self.start_hover_animation(100)  # Animate to hover state
        
    def leaveEvent(self, event):
        self.start_hover_animation(0)  # Animate back to normal state
        
    def start_hover_animation(self, target_value):
        self.animation.stop()
        self.animation.setStartValue(self._animation_progress)
        self.animation.setEndValue(target_value)
        self.animation.start()

class BaseAnalyzer(QWidget):
    def __init__(self, home_callback, name="Analyzer"):
        super().__init__()
        self.home_callback = home_callback
        layout = QVBoxLayout()
        
        title = QLabel(f"{name} Module")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(title)
        
        message = QLabel("This analyzer module could not be loaded")
        layout.addWidget(message)
        
        home_btn = GradientButton("Return to Home")
        home_btn.clicked.connect(self.home_callback)
        layout.addWidget(home_btn)
        
        self.setLayout(layout)

class UnifiedCDRAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mobile Money & Call Detail Record Analyzer")
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {BACKGROUND_COLOR};
                color: {TEXT_COLOR};
            }}
            QStatusBar {{
                background-color: {SIDEBAR_COLOR};
                color: {TEXT_COLOR};
            }}
        """)
        
        # Set window icon
        self.setWindowIcon(QIcon(self.resource_path("assets/gp_logox.ico")))
        
        # System tray
        self.setup_system_tray()
        
        # Set window size
        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(
            int(screen.width() * 0.1),
            int(screen.height() * 0.1),
            int(screen.width() * 0.8),
            int(screen.height() * 0.8)
        )
        
        # Initialize recent files list
        self.recent_files = []
        self.max_recent_files = 5
        
        self.init_ui()
        self.show()

    def resource_path(self, relative_path):
        """ Get absolute path to resource """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def setup_system_tray(self):
        if not QSystemTrayIcon.isSystemTrayAvailable():
            return
            
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(self.resource_path("assets/gp_logox.ico")))
        
        tray_menu = QMenu()
        show_action = QAction("Show", self)
        show_action.triggered.connect(self.show_normal)
        tray_menu.addAction(show_action)
        
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.quit_immediately)
        tray_menu.addAction(exit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
    def show_normal(self):
        self.show()
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        self.activateWindow()
        
    def closeEvent(self, event):
        if QSystemTrayIcon.isSystemTrayAvailable():
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "Mobile Money & CDR Analyzer",
                "The application is still running in the system tray",
                QSystemTrayIcon.Information,
                2000
            )
        else:
            event.accept()

    def init_ui(self):
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Sidebar
        sidebar = self.create_sidebar()
        main_layout.addWidget(sidebar)

        # Main content area
        self.stacked_widget = QStackedWidget()
        main_layout.addWidget(self.stacked_widget, stretch=1)

        # Add main page content
        self.add_main_page_content()

        # Preload apps
        self.preload_apps()

        # Set default view
        self.stacked_widget.setCurrentIndex(0)

        # Status bar
        self.setup_status_bar()
        
        self.setCentralWidget(main_widget)

    def create_sidebar(self):
        sidebar = QFrame()
        sidebar.setStyleSheet(f"""
            QFrame {{
                background-color: {SIDEBAR_COLOR};
                border-right: 1px solid #444;
            }}
            QLabel {{
                color: #FFFFFF;
            }}
        """)
        sidebar.setFixedWidth(280)

        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setAlignment(Qt.AlignTop)
        sidebar_layout.setSpacing(15)
        sidebar_layout.setContentsMargins(15, 20, 15, 15)

        # Home button with icon
        home_button = GradientButton("🏠 Dashboard")
        home_button.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                font-weight: bold;
                color: white;
                padding-left: 15px;
                text-align: left;
            }
        """)
        home_button.clicked.connect(self.open_home)
        sidebar_layout.addWidget(home_button)

        # Separator
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setStyleSheet("color: #444; margin: 10px 0;")
        sidebar_layout.addWidget(separator)

        # Module buttons section
        modules_label = QLabel("ANALYSIS MODULES")
        modules_label.setStyleSheet("""
            QLabel {
                font-size: 11px;
                font-weight: bold;
                color: #AAAAAA;
                letter-spacing: 1px;
                padding: 5px 0;
                margin-top: 10px;
            }
        """)
        sidebar_layout.addWidget(modules_label)

        # Modern module buttons with brand colors
        modules = [
            ("MTN CDR Analysis", "#FFD700", "#FFA500", "chart-line"),
            ("MTN Mobile Money", "#F9D423", "#E65C00", "money-bill-wave"),
            ("Telecel CDR", "#93291E", "#ED213A", "table"),
            ("Telecel Cash", "#ED213A", "#93291E", "money-check"),
            ("AirtelTigo CDR", "#6A11CB", "#2575FC", "chart-bar"),
            ("AirtelTigo Cash", "#0575E6", "#021B79", "wallet"),
        ]

        for text, color1, color2, icon_name in modules:
            btn = GradientButton(f"  {text}")
            btn.color1 = QColor(color1)
            btn.color2 = QColor(color2)
            btn.hover_color1 = QColor(color1).lighter(120)
            btn.hover_color2 = QColor(color2).lighter(120)
            
            btn.setIcon(QIcon.fromTheme(icon_name))
            btn.setIconSize(QSize(20, 20))
            btn.setStyleSheet(f"""
                QPushButton {{
                    font-size: 13px;
                    font-weight: bold;
                    color: white;
                    padding-left: 15px;
                    text-align: left;
                    border: none;
                }}
            """)
            
            if "MTN CDR" in text:
                btn.clicked.connect(self.open_mtn_analysis)
            elif "MTN Mobile" in text:
                btn.clicked.connect(self.open_mobile_money_analysis)
            elif "Telecel CDR" in text:
                btn.clicked.connect(self.open_telecel_analysis)
            elif "Telecel Cash" in text:
                btn.clicked.connect(self.open_telecel_cash_analysis)
            elif "AirtelTigo CDR" in text:
                btn.clicked.connect(self.open_airteltigo_cdr_analysis)
            elif "AirtelTigo Cash" in text:
                btn.clicked.connect(self.open_airteltigo_cash_analysis)
                
            sidebar_layout.addWidget(btn)

        sidebar_layout.addStretch()

        # Exit button
        exit_button = GradientButton("⏻ Exit Application")
        exit_button.color1 = QColor(EXIT_BUTTON_COLOR)
        exit_button.color2 = QColor("#C0392B")
        exit_button.hover_color1 = QColor("#FF0000")
        exit_button.hover_color2 = QColor("#990000")
        exit_button.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                color: white;
                text-align: center;
                margin-top: 20px;
            }
        """)
        exit_button.clicked.connect(self.confirm_exit)
        sidebar_layout.addWidget(exit_button)

        return sidebar

    def add_main_page_content(self):
        main_page = QWidget()
        main_page_layout = QVBoxLayout(main_page)
        main_page_layout.setContentsMargins(30, 30, 30, 30)
        main_page_layout.setSpacing(30)

        # Header
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)

        title_label = QLabel("Mobile Money & CDR Analyzer")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 32px;
                font-weight: bold;
                color: #FFFFFF;
            }
        """)
        header_layout.addWidget(title_label, stretch=1)

        main_page_layout.addWidget(header)

        # Dashboard
        dashboard = QWidget()
        dashboard_layout = QHBoxLayout(dashboard)
        dashboard_layout.setContentsMargins(0, 0, 0, 0)
        dashboard_layout.setSpacing(20)

        # Quick actions
        quick_actions = self.create_quick_actions_panel()
        dashboard_layout.addWidget(quick_actions, 1)

        # Recent activity
        recent_activity = self.create_recent_activity_panel()
        dashboard_layout.addWidget(recent_activity, 2)

        main_page_layout.addWidget(dashboard)

        # Footer with logo
        footer = QWidget()
        footer_layout = QVBoxLayout(footer)
        footer_layout.setAlignment(Qt.AlignCenter)

        logo_label = QLabel()
        logo_pixmap = QPixmap(self.resource_path("assets/gp_logo6.png"))
        if not logo_pixmap.isNull():
            logo_label.setPixmap(logo_pixmap.scaledToWidth(350, Qt.SmoothTransformation))
            logo_label.setAlignment(Qt.AlignCenter)
            footer_layout.addWidget(logo_label)

        version_label = QLabel("Version 1.0.00.0 - Professional Edition\n © 2025 Terence")
        version_label.setStyleSheet("color: #888; font-size: 12px;")
        version_label.setAlignment(Qt.AlignCenter)
        footer_layout.addWidget(version_label)

        main_page_layout.addWidget(footer)

        self.stacked_widget.addWidget(main_page)

    def create_quick_actions_panel(self):
        panel = QFrame()
        panel.setStyleSheet(f"""
            QFrame {{
                background-color: {SIDEBAR_COLOR};
                border-radius: 10px;
                padding: 15px;
            }}
        """)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(15)

        title = QLabel("Quick Actions")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)

        actions = [
            ("Clean CDR File", "upload", self.upload_excel_csv),
            ("Open Recent", "history", self.open_recent_file),
            ("Export Results", "file-export", self.export_results),
            ("Settings", "cog", self.open_settings),
        ]

        for text, icon_name, callback in actions:
            btn = GradientButton(text)
            btn.setIcon(QIcon.fromTheme(icon_name))
            btn.setStyleSheet(f"""
                QPushButton {{
                    text-align: left;
                    padding: 10px 15px;
                    border-radius: 5px;
                    color: {TEXT_COLOR};
                }}
            """)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(callback)
            layout.addWidget(btn)

        layout.addStretch()
        return panel

    def create_recent_activity_panel(self):
        panel = QFrame()
        panel.setStyleSheet(f"""
            QFrame {{
                background-color: {SIDEBAR_COLOR};
                border-radius: 10px;
                padding: 15px;
            }}
        """)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(15)

        title = QLabel("Recent Activity")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)

        self.recent_files_list = QListWidget()
        self.recent_files_list.setStyleSheet("""
            QListWidget {
                background-color: #3A3A4A;
                border-radius: 5px;
                padding: 5px;
            }
            QListWidget::item {
                padding: 5px;
                border-bottom: 1px solid #444;
            }
            QListWidget::item:hover {
                background-color: #4A4A5A;
            }
        """)
        self.recent_files_list.itemClicked.connect(self.open_selected_recent_file)
        
        # Add sample recent files
        for i in range(3):
            self.recent_files_list.addItem(f"Sample_file_{i+1}.csv")
        
        layout.addWidget(self.recent_files_list)
        layout.addStretch()
        return panel

    def preload_apps(self):
        """Preload all analyzer apps with error handling"""
        try:
            from modules.airteltigo_cdr_analyzer import AirtelTigoCDRAnalyzer
            from modules.airteltigo_cash_analyzer import AirtelTigoCashAnalyzer
        except ImportError as e:
            logger.error(f"Failed to load AirtelTigo modules: {e}")
            AirtelTigoCDRAnalyzer = lambda cb: BaseAnalyzer(cb, "AirtelTigo CDR")
            AirtelTigoCashAnalyzer = lambda cb: BaseAnalyzer(cb, "AirtelTigo Cash")

        try:
            from modules.mtn_cdr_analyzer import MTNCDRAnalyzer
            from modules.telecel_cdr_analyzer import TelecelCDRAnalyzer
            from modules.mobile_money_analyzer import MobileMoneyAnalyzer
            from modules.telecel_cash_analyzer import TelecelCashAnalyzer
        except ImportError as e:
            logger.error(f"Failed to load other modules: {e}")
            MTNCDRAnalyzer = lambda cb: BaseAnalyzer(cb, "MTN CDR")
            TelecelCDRAnalyzer = lambda cb: BaseAnalyzer(cb, "Telecel CDR")
            MobileMoneyAnalyzer = lambda cb: BaseAnalyzer(cb, "Mobile Money")
            TelecelCashAnalyzer = lambda cb: BaseAnalyzer(cb, "Telecel Cash")

        # Initialize all analyzers
        self.mtn_analyzer = MTNCDRAnalyzer(self.open_home)
        self.telecel_analyzer = TelecelCDRAnalyzer(self.open_home)
        self.mobile_money_analyzer = MobileMoneyAnalyzer(self.open_home)
        self.telecel_cash_analyzer = TelecelCashAnalyzer(self.open_home)
        self.airteltigo_cdr_analyzer = AirtelTigoCDRAnalyzer(self.open_home)
        self.airteltigo_cash_analyzer = AirtelTigoCashAnalyzer(self.open_home)

        # Add to stacked widget
        self.stacked_widget.addWidget(self.mtn_analyzer)
        self.stacked_widget.addWidget(self.telecel_analyzer)
        self.stacked_widget.addWidget(self.mobile_money_analyzer)
        self.stacked_widget.addWidget(self.telecel_cash_analyzer)
        self.stacked_widget.addWidget(self.airteltigo_cdr_analyzer)
        self.stacked_widget.addWidget(self.airteltigo_cash_analyzer)

    def setup_status_bar(self):
        status_bar = QStatusBar()
        self.setStatusBar(status_bar)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(200)
        self.progress_bar.setVisible(False)
        status_bar.addPermanentWidget(self.progress_bar)
        
        self.status_label = QLabel("Ready")
        status_bar.addWidget(self.status_label, 1)

    def update_status(self, message, message_type="info", timeout=5000):
        color_map = {
            "info": TEXT_COLOR,
            "success": SUCCESS_COLOR,
            "error": ERROR_COLOR,
            "warning": WARNING_COLOR
        }
        
        self.status_label.setText(message)
        self.status_label.setStyleSheet(f"color: {color_map.get(message_type, TEXT_COLOR)};")
        
        if timeout > 0:
            QTimer.singleShot(timeout, lambda: self.status_label.setText("Ready"))

    def upload_excel_csv(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select CDR File",
            "",
            "All Supported Files (*.csv *.xlsx *.xls);;CSV Files (*.csv);;Excel Files (*.xlsx *.xls)",
            options=options
        )
        
        if file_path:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            self.update_status(f"Processing {os.path.basename(file_path)}...", "info")
            
            # Simulate processing
            for i in range(1, 101):
                QTimer.singleShot(i * 20, lambda i=i: self.progress_bar.setValue(i))
                
            QTimer.singleShot(2500, lambda: self.finish_upload(file_path))

    def finish_upload(self, file_path):
        self.progress_bar.setVisible(False)
        self.update_status(f"Successfully processed {os.path.basename(file_path)}", "success")
        
        # Add to recent files
        self.add_recent_file(file_path)
        
        QMessageBox.information(
            self,
            "Processing Complete",
            f"Successfully processed file:\n{file_path}\n\n"
            "• 1,245 records analyzed\n"
            "• 5 anomalies detected\n"
            "• Report generated",
            QMessageBox.Ok
        )

    def add_recent_file(self, file_path):
        """Add a file to the recent files list"""
        if file_path in self.recent_files:
            self.recent_files.remove(file_path)
        
        self.recent_files.insert(0, file_path)
        if len(self.recent_files) > self.max_recent_files:
            self.recent_files = self.recent_files[:self.max_recent_files]
        
        # Update the list widget
        self.recent_files_list.clear()
        for file in self.recent_files:
            self.recent_files_list.addItem(os.path.basename(file))

    def open_selected_recent_file(self, item):
        """Handle opening a file from the recent files list"""
        file_name = item.text()
        matching_files = [f for f in self.recent_files if os.path.basename(f) == file_name]
        
        if matching_files:
            self.open_recent_file(matching_files[0])

    def open_recent_file(self, file_path=None):
        """Handle opening of recent files"""
        if file_path is None:
            # Show dialog to select from recent files
            if not self.recent_files:
                QMessageBox.information(self, "No Recent Files", "No recent files available")
                return
            
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Open Recent File",
                "",
                "All Supported Files (*.csv *.xlsx *.xls);;CSV Files (*.csv);;Excel Files (*.xlsx *.xls)",
                options=QFileDialog.Options()
            )
        
        if file_path and os.path.exists(file_path):
            self.update_status(f"Opening {os.path.basename(file_path)}...", "info")
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # Simulate opening
            for i in range(1, 101):
                QTimer.singleShot(i * 10, lambda i=i: self.progress_bar.setValue(i))
                
            QTimer.singleShot(1500, lambda: self.finish_open(file_path))

    def finish_open(self, file_path):
        self.progress_bar.setVisible(False)
        self.update_status(f"Successfully opened {os.path.basename(file_path)}", "success")
        QMessageBox.information(
            self,
            "File Opened",
            f"Successfully opened file:\n{file_path}",
            QMessageBox.Ok
        )

    def export_results(self):
        """Handle export results functionality"""
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Results",
            "",
            "PDF Files (*.pdf);;Excel Files (*.xlsx);;CSV Files (*.csv)",
            options=options
        )
        
        if file_path:
            self.update_status(f"Exporting results to {os.path.basename(file_path)}...", "info")
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # Simulate export
            for i in range(1, 101):
                QTimer.singleShot(i * 15, lambda i=i: self.progress_bar.setValue(i))
                
            QTimer.singleShot(2000, lambda: self.finish_export(file_path))

    def finish_export(self, file_path):
        """Finish the export process"""
        self.progress_bar.setVisible(False)
        self.update_status(f"Successfully exported to {os.path.basename(file_path)}", "success")
        QMessageBox.information(
            self,
            "Export Complete",
            f"Results successfully exported to:\n{file_path}",
            QMessageBox.Ok
        )

    def open_settings(self):
        """Handle opening settings"""
        self.update_status("Opening settings...", "info")
        QMessageBox.information(
            self,
            "Settings",
            "This would open the application settings",
            QMessageBox.Ok
        )
        QTimer.singleShot(1000, lambda: self.update_status("Settings opened", "info"))

    def open_home(self):
        """Return to home screen"""
        self.stacked_widget.setCurrentIndex(0)
        self.update_status("Returned to home screen", "info")

    def open_mtn_analysis(self):
        """Open MTN CDR analyzer"""
        self.stacked_widget.setCurrentWidget(self.mtn_analyzer)
        self.update_status("MTN CDR Analyzer loaded", "info")

    def open_mobile_money_analysis(self):
        """Open Mobile Money analyzer"""
        self.stacked_widget.setCurrentWidget(self.mobile_money_analyzer)
        self.update_status("Mobile Money Analyzer loaded", "info")

    def open_telecel_analysis(self):
        """Open Telecel CDR analyzer"""
        self.stacked_widget.setCurrentWidget(self.telecel_analyzer)
        self.update_status("Telecel CDR Analyzer loaded", "info")

    def open_telecel_cash_analysis(self):
        """Open Telecel Cash analyzer"""
        self.stacked_widget.setCurrentWidget(self.telecel_cash_analyzer)
        self.update_status("Telecel Cash Analyzer loaded", "info")

    def open_airteltigo_cdr_analysis(self):
        """Open AirtelTigo CDR analyzer"""
        self.stacked_widget.setCurrentWidget(self.airteltigo_cdr_analyzer)
        self.update_status("AirtelTigo CDR Analyzer loaded", "info")

    def open_airteltigo_cash_analysis(self):
        """Open AirtelTigo Cash analyzer"""
        self.stacked_widget.setCurrentWidget(self.airteltigo_cash_analyzer)
        self.update_status("AirtelTigo Cash Analyzer loaded", "info")

    def confirm_exit(self):
        """Confirm before exiting"""
        reply = QMessageBox.question(
            self,
            "Confirm Exit",
            "Are you sure you want to exit?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self.quit_immediately()

    def quit_immediately(self):
        """Immediate application termination"""
        QApplication.quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Set application style and palette
    app.setStyle("Fusion")
    
    # Set custom palette
    palette = app.palette()
    palette.setColor(palette.Window, QColor(BACKGROUND_COLOR))
    palette.setColor(palette.WindowText, QColor(TEXT_COLOR))
    palette.setColor(palette.Base, QColor("#2D2D3A"))
    palette.setColor(palette.AlternateBase, QColor(SIDEBAR_COLOR))
    palette.setColor(palette.ToolTipBase, QColor(BACKGROUND_COLOR))
    palette.setColor(palette.ToolTipText, QColor(TEXT_COLOR))
    palette.setColor(palette.Text, QColor(TEXT_COLOR))
    palette.setColor(palette.Button, QColor(BUTTON_COLOR))
    palette.setColor(palette.ButtonText, QColor(TEXT_COLOR))
    palette.setColor(palette.BrightText, Qt.red)
    palette.setColor(palette.Highlight, QColor(HIGHLIGHT_COLOR))
    palette.setColor(palette.HighlightedText, Qt.black)
    app.setPalette(palette)
    
    window = UnifiedCDRAnalyzer()
    sys.exit(app.exec_())