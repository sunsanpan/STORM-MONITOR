import sys
import time
import serial
import serial.tools.list_ports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                         QHBoxLayout, QComboBox, QLabel, QPushButton, 
                         QTextEdit, QLineEdit, QGroupBox, QCheckBox,
                         QFileDialog)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QObject
from PyQt5.QtGui import QFont, QPalette, QColor, QTextCursor, QTextCharFormat
import pyqtgraph as pg
from pyqtgraph.exporters import ImageExporter
import numpy as np

class SerialReceiver(QObject):
    """Separate class to handle receiving serial data"""
    data_received = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    
    def __init__(self, serial_connection):
        super().__init__()
        self.serial_connection = serial_connection
        
    def receive_data(self):
        """Read data from serial port in a more Putty-like way"""
        if not self.serial_connection or not self.serial_connection.is_open:
            return
            
        try:
            if self.serial_connection.in_waiting > 0:
                # Read available data
                data = self.serial_connection.read(self.serial_connection.in_waiting).decode('utf-8', errors='replace')
                
                # Process the data immediately
                if data:
                    # Emit the raw data with control characters preserved
                    self.data_received.emit(data)
                    
        except Exception as e:
            self.error_occurred.emit(f"Error reading data: {str(e)}")


class ComboBoxWithEdit(QComboBox):
    """Custom ComboBox that allows text editing"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)  # Don't automatically add user text to the dropdown list


class DirectInputLineEdit(QLineEdit):
    """Custom LineEdit that sends each character as it's typed"""
    char_typed = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        # Set focus policy to make sure it can receive key events
        self.setFocusPolicy(Qt.StrongFocus)
    
    def keyPressEvent(self, event):
        """Override key press event to emit signal for each character"""
        # Process the key press normally first
        super().keyPressEvent(event)
        
        # Get the typed character
        text = event.text()
        
        # If there is a character (not just a modifier key), emit it
        if text and len(text) > 0:
            self.char_typed.emit(text)


class SerialPlotter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Serial Plotter')
        self.setGeometry(200, 200, 800, 400)
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Add toolbar for controls
        toolbar_layout = QHBoxLayout()
        layout.addLayout(toolbar_layout)

        # Add theme button before pause button
        self.theme_button = QPushButton("Dark Theme")
        self.theme_button.setCheckable(True)
        self.theme_button.setChecked(True)  # Start with dark theme
        self.theme_button.clicked.connect(self.toggle_theme)
        toolbar_layout.addWidget(self.theme_button)

        # Add pause/resume button
        self.pause_button = QPushButton("Pause")
        self.pause_button.setCheckable(True)
        self.pause_button.clicked.connect(self.toggle_pause)
        toolbar_layout.addWidget(self.pause_button)

        # Add export button
        self.export_button = QPushButton("Export Plot")
        self.export_button.clicked.connect(self.export_plot)
        toolbar_layout.addWidget(self.export_button)

        # Add legend visibility toggle
        self.legend_visible_check = QCheckBox("Show Legend")
        self.legend_visible_check.setChecked(True)
        self.legend_visible_check.stateChanged.connect(self.toggle_legend_visibility)
        toolbar_layout.addWidget(self.legend_visible_check)

        # Add variable name controls
        var_name_layout = QHBoxLayout()
        self.var_names = QLineEdit()
        self.var_names.setPlaceholderText("Variable names (comma-separated)")
        
        # Add clear names button
        self.clear_names_button = QPushButton("Clear Names")
        self.clear_names_button.clicked.connect(self.clear_variable_names)
        
        # Add apply button for names
        self.apply_names_button = QPushButton("Apply Names")
        self.apply_names_button.clicked.connect(self.apply_variable_names)
        
        var_name_layout.addWidget(self.var_names)
        var_name_layout.addWidget(self.apply_names_button)
        var_name_layout.addWidget(self.clear_names_button)
        toolbar_layout.addLayout(var_name_layout)

        # Store variable names dictionary
        self.var_name_dict = {}  # Store assigned names

        toolbar_layout.addStretch()  # Add stretch to keep buttons left-aligned
        
        # Create plot widget
        self.plot_widget = pg.PlotWidget()
        layout.addWidget(self.plot_widget)
        
        # Add variable visibility panel at bottom with fixed width and left alignment
        self.var_panel = QWidget()
        self.var_panel.setFixedWidth(400)  # Fixed width for checkbox area
        self.var_panel_layout = QHBoxLayout()
        self.var_panel_layout.setSpacing(15)  # Set spacing between checkboxes
        self.var_panel_layout.setContentsMargins(10, 5, 10, 5)  # Add some padding
        self.var_panel_layout.setAlignment(Qt.AlignLeft)  # Align checkboxes to the left
        self.var_panel.setLayout(self.var_panel_layout)
        
        # Create wrapper layout to keep panel left-aligned
        bottom_layout = QHBoxLayout()
        bottom_layout.addWidget(self.var_panel)
        bottom_layout.addStretch()  # Push everything to the left
        layout.addLayout(bottom_layout)
        
        # Store visibility checkboxes
        self.var_checkboxes = []
        
        # Setup plot
        self.plot_widget.setBackground('k')  # Black background
        self.plot_widget.showGrid(x=True, y=True, alpha=0.3)
        
        # Add legend
        self.legend = None
        # Track if legend was initialized
        self.legend_initialized = False
        
        # Initialize data storage
        self.max_points = 500  # Maximum number of points to display
        self.data_lines = []  # List to store plot lines
        self.data_buffers = []  # List to store data for each line
        self.colors = ['g', 'r', 'b', 'y', 'c', 'm', 'w']  # Line colors
        self.paused = False  # Pause state
        
        # Initialize empty legend and name dictionary ONCE
        self.var_name_dict = {}
        self.data_lines = []
        self.data_buffers = []
        self.var_checkboxes = []
        self.legend = None
        
        # Remove duplicate initializations that were causing issues
        self.init_legend()
        
    def init_legend(self):
        """Initialize or reinitialize the legend with current names"""
        if self.legend:
            self.plot_widget.removeItem(self.legend)
        self.legend = self.plot_widget.addLegend()
        
        # Add all current lines with their names
        for i, line in enumerate(self.data_lines):
            name = self.var_name_dict.get(i, f"Var {i+1}")
            self.legend.addItem(line, name)
        
    def rebuild_legend(self):
        """Remove and rebuild the legend with current names for all lines."""
        self.plot_widget.clear()
        if self.legend:
            self.plot_widget.removeItem(self.legend)
            self.legend = None

        # Clear all checkboxes
        while self.var_panel_layout.count():
            item = self.var_panel_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.var_checkboxes.clear()

        # Only keep buffers for variables that actually have data
        # Remove trailing buffers that are empty
        while self.data_buffers and not self.data_buffers[-1]:
            self.data_buffers.pop()

        # Synchronize var_name_dict with the actual number of variables
        # self.var_name_dict = {i: self.var_name_dict.get(i, f"Var {i+1}") for i in range(len(self.data_buffers))}

        self.data_lines = []
        self.legend = self.plot_widget.addLegend()

        for i, line_data in enumerate(self.data_buffers):
            color = self.colors[i % len(self.colors)]
            line = self.plot_widget.plot(pen=color, name=self.var_name_dict.get(i, f"Var {i+1}"))
            line.setData(line_data)
            self.data_lines.append(line)

            checkbox = QCheckBox(self.var_name_dict.get(i, f"Var {i+1}"))
            checkbox.setFixedWidth(80)
            checkbox.setChecked(True)
            checkbox.stateChanged.connect(lambda state, idx=i: self.toggle_var_visibility(idx, state))
            self.var_checkboxes.append(checkbox)
            self.var_panel_layout.addWidget(checkbox)

    def toggle_pause(self, checked):
        """Pause or resume plotting"""
        self.paused = checked
        self.pause_button.setText("Resume" if self.paused else "Pause")

    def toggle_theme(self, checked):
        """Toggle between dark and light theme"""
        if checked:
            self.plot_widget.setBackground('k')  # Black background
            self.theme_button.setText("Dark Theme")
            # Update grid color for dark theme
            self.plot_widget.showGrid(x=True, y=True, alpha=0.3)
        else:
            self.plot_widget.setBackground('w')  # White background
            self.theme_button.setText("Light Theme")
            # Update grid color for light theme
            self.plot_widget.showGrid(x=True, y=True, alpha=0.5)
        
        # Update plot colors based on theme
        self.update_plot_colors()

    def update_plot_colors(self):
        """Update plot colors based on current theme"""
        is_dark = self.theme_button.isChecked()
        # Color schemes for dark and light themes
        self.colors = ['g', 'r', 'b', 'y', 'c', 'm'] if is_dark else ['b', 'r', 'g', 'm', 'c', 'k']
        
        # Update existing line colors
        for i, line in enumerate(self.data_lines):
            color = self.colors[i % len(self.colors)]
            line.setPen(color)

    def apply_variable_names(self):
        """Apply new variable names from input field."""
        try:
            names = [n.strip() for n in self.var_names.text().split(',') if n.strip()]
            # Clear existing names before applying new ones
            self.var_name_dict.clear()
            for i, name in enumerate(names):
                if i < len(self.data_buffers):
                    self.var_name_dict[i] = name
            self.rebuild_legend()
            self.var_names.clear()

            # Update checkbox labels with custom names
            for i, checkbox in enumerate(self.var_checkboxes):
                name = self.var_name_dict.get(i, f"Var {i+1}")
                checkbox.setText(name)

        except Exception as e:
            print(f"Error applying names: {str(e)}")

    def clear_variable_names(self):
        """Clear all custom variable names"""
        # Clear the dictionary and names input
        self.var_name_dict.clear()
        self.var_names.clear()
        
        # Rebuild legend with default names
        self.rebuild_legend()
        
        # Update checkbox labels
        for i, checkbox in enumerate(self.var_checkboxes):
            checkbox.setText(f"Var {i+1}")

    def toggle_legend_visibility(self, state):
        """Toggle the visibility of the legend"""
        if self.legend:
            self.legend.setVisible(bool(state))

    def export_plot(self):
        """Export the current plot as an image"""
        # Store current state
        was_paused = self.paused
        self.paused = True
        
        try:
            file_name, _ = QFileDialog.getSaveFileName(
                self,
                "Export Plot As...",
                "",
                "PNG Image (*.png);;All Files (*.*)"
            )
            
            if file_name:
                exporter = ImageExporter(self.plot_widget.getPlotItem())
                exporter.parameters()['width'] = 1920
                exporter.export(file_name)
                
                # Force rebuild to clean up any state changes from export
                # Also clears data_lines to avoid legend/checkbox duplication
                self.rebuild_legend()
                
        finally:
            # Restore pause state
            self.paused = was_paused

    def process_data(self, data):
        """Process incoming data string and update plot"""
        if self.paused:
            return

        try:
            # Split data using various separators (tabs, commas, or spaces)
            values = []
            for part in data.replace('\t', ' ').split(','):
                values.extend([float(x) for x in part.split() if x.strip()])

            # Truncate buffers if number of variables decreases
            if len(values) < len(self.data_buffers):
                self.data_buffers = self.data_buffers[:len(values)]
                self.data_lines = self.data_lines[:len(values)]
                # Remove extra checkboxes as well
                while len(self.var_checkboxes) > len(values):
                    cb = self.var_checkboxes.pop()
                    self.var_panel_layout.removeWidget(cb)
                    cb.deleteLater()

            # Synchronize var_name_dict with the actual number of variables
            # self.var_name_dict = {i: self.var_name_dict.get(i, f"Var {i+1}") for i in range(len(values))}

            # Initialize plot lines if needed
            while len(self.data_lines) < len(values):
                color = self.colors[len(self.data_lines) % len(self.colors)]
                self.data_buffers.append([])
                line = self.plot_widget.plot(pen=color, name=f"Var {len(self.data_lines)+1}")
                self.data_lines.append(line)
                self.add_var_checkbox(len(self.data_lines)-1)

            # Update data buffers
            for i, value in enumerate(values):
                if i >= len(self.data_buffers):
                    break
                self.data_buffers[i].append(value)
                if len(self.data_buffers[i]) > self.max_points:
                    self.data_buffers[i].pop(0)
                self.data_lines[i].setData(self.data_buffers[i])

            # Auto-scale Y axis
            self.plot_widget.enableAutoRange(axis='y')

        except ValueError:
            pass

    def add_var_checkbox(self, index):
        """Add a checkbox for a new variable"""
        name = self.var_name_dict.get(index, f"Var {index+1}")
        checkbox = QCheckBox(name)
        checkbox.setFixedWidth(80)
        checkbox.setChecked(True)
        checkbox.stateChanged.connect(lambda state, idx=index: self.toggle_var_visibility(idx, state))
        self.var_checkboxes.append(checkbox)
        self.var_panel_layout.addWidget(checkbox)
        
    def toggle_var_visibility(self, index, state):
        """Toggle visibility of a variable's plot line"""
        if index < len(self.data_lines):
            self.data_lines[index].setVisible(bool(state))


class SerialMonitor(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Initialize serial connection
        self.serial_connection = None
        self.is_connected = False
        self.show_commands = True  # Default to showing commands
        self.auto_reconnect = False  # Default to no auto-reconnect
        self.last_port = None
        self.last_baud = None
        self.manual_disconnect = False  # Flag to track if disconnect was manual
        
        # Keep track of cursor position for proper CR handling
        self.current_line_start = 0

        # Initialize state variables before calling init_ui
        self.blink_state = True  # Starting state (visible)
        self.data_active = False  # Initialize data activity flag
        self.last_activity_time = time.time()
        
        # Always show cursor when connected (new flag to fix cursor visibility)
        self.show_cursor = True
        
        # Add file handling attributes
        self.save_file = None
        self.is_saving = False
        
        # Add plotter
        self.plotter = None
        
        # Setup dark theme
        self.set_dark_theme()
        
        # Setup UI
        self.init_ui()
        
        # Setup refresh timer for available ports
        self.port_refresh_timer = QTimer()
        self.port_refresh_timer.timeout.connect(self.refresh_ports)
        self.port_refresh_timer.start(1000)  # Refresh every 1 second
        
        # Serial receiver setup (will be initialized when connected)
        self.serial_receiver = None
        self.read_timer = QTimer()
        
        # Setup reconnect timer
        self.reconnect_timer = QTimer()
        self.reconnect_timer.timeout.connect(self.check_reconnect)
        self.reconnect_timer.start(1000)  # Check every second
        
        # Blinking cursor setup
        self.blink_state = True  # Starting state (visible)
        self.blink_timer = QTimer()
        self.blink_timer.timeout.connect(self.toggle_blink_cursor)
        self.blink_timer.start(500)  # Blink every 500ms
        
        # Data activity timer - check every 500ms with 2 second timeout
        self.data_active = True  # Start as active to avoid initial IDLE
        self.last_activity_time = time.time()
        self.data_activity_timer = QTimer()
        self.data_activity_timer.timeout.connect(self.check_data_activity)
        self.data_activity_timer.start(500)  # Check every 500ms
    
    def set_dark_theme(self):
        """Apply dark theme to the application"""
        app = QApplication.instance()
        
        # Create dark palette
        dark_palette = QPalette()
        
        # Set color roles
        dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.WindowText, Qt.white)
        dark_palette.setColor(QPalette.Base, QColor(35, 35, 35))
        dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ToolTipBase, QColor(25, 25, 25))
        dark_palette.setColor(QPalette.ToolTipText, Qt.white)
        dark_palette.setColor(QPalette.Text, Qt.white)
        dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ButtonText, Qt.white)
        dark_palette.setColor(QPalette.BrightText, Qt.red)
        dark_palette.setColor(QPalette.Link, QColor(42, 130, 218))
        dark_palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
        dark_palette.setColor(QPalette.HighlightedText, Qt.black)
        
        # Apply palette
        app.setPalette(dark_palette)
        
        # Set stylesheet for more detailed control
        app.setStyleSheet("""
            QGroupBox {
                border: 1px solid #777;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center;
                padding: 0 5px;
                color: #DDD;
            }
            QPushButton {
                background-color: #444;
                border: 1px solid #555;
                border-radius: 3px;
                padding: 5px;
                color: white;
            }
            QPushButton:hover {
                background-color: #555;
            }
            QPushButton:pressed {
                background-color: #2a82da;
            }
            QComboBox {
                background-color: #444;
                border: 1px solid #555;
                border-radius: 3px;
                padding: 5px;
                color: white;
            }
            QComboBox:hover {
                border: 1px solid #777;
            }
            QComboBox QAbstractItemView {
                background-color: #444;
                color: white;
                selection-background-color: #2a82da;
            }
            QLineEdit {
                background-color: #333;
                border: 1px solid #555;
                border-radius: 3px;
                padding: 5px;
                color: white;
            }
            QLineEdit:focus {
                border: 1px solid #2a82da;
            }
            QTextEdit {
                background-color: #232323;
                border: 1px solid #444;
                border-radius: 3px;
                color: #DDD;
                font-family: "Courier New";
            }
            QCheckBox {
                color: white;
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 9px;
                border: 1px solid #555;
            }
            QCheckBox::indicator:unchecked {
                background-color: #333;
            }
            QCheckBox::indicator:checked {
                background-color: #2a82da;
            }
        """)
        
    def init_ui(self):
        # Set window properties
        self.setWindowTitle('STORM : Serial Monitor')
        self.setGeometry(100, 100, 800, 600)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Create connection settings group
        connection_group = QGroupBox("Connection Settings")
        connection_layout = QHBoxLayout()
        
        # Port selection with editable dropdown
        port_label = QLabel("Port:")
        self.port_combo = ComboBoxWithEdit()
        self.port_combo.setPlaceholderText("Select or type COM port")
        self.refresh_ports()
        
        # Baud rate selection with editable dropdown
        baud_label = QLabel("Baud Rate:")
        self.baud_combo = ComboBoxWithEdit()
        baud_rates = ['9600', '19200', '38400', '57600', '115200', '230400', '460800', '921600']
        self.baud_combo.addItems(baud_rates)
        self.baud_combo.setCurrentText('115200')  # Set default to 115200
        
        # Connect button
        self.connect_button = QPushButton("Connect")
        self.connect_button.clicked.connect(self.toggle_connection)
        
        # Auto-reconnect checkbox
        self.reconnect_check = QCheckBox("Auto-Reconnect")
        self.reconnect_check.setChecked(False)  # Default to unchecked
        self.reconnect_check.stateChanged.connect(self.toggle_auto_reconnect)
        
        # Show commands checkbox
        self.echo_check = QCheckBox("Echo Commands")
        self.echo_check.setChecked(True)  # Default to checked
        self.echo_check.stateChanged.connect(self.toggle_echo)
        
        # Add clear button
        self.clear_button = QPushButton("Clear")
        self.clear_button.clicked.connect(self.clear_monitor)
        
        # Add cursor visibility checkbox
        self.cursor_check = QCheckBox("Show Cursor")
        self.cursor_check.setChecked(True)  # Default to checked
        self.cursor_check.stateChanged.connect(self.toggle_cursor_visibility)
        
        # Add save button
        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(self.toggle_save)
        
        # Add plotter button
        self.plotter_button = QPushButton("Plotter")
        self.plotter_button.clicked.connect(self.toggle_plotter)
        connection_layout.addWidget(self.plotter_button)
        
        # Add to connection layout
        connection_layout.addWidget(port_label)
        connection_layout.addWidget(self.port_combo)
        connection_layout.addWidget(baud_label)
        connection_layout.addWidget(self.baud_combo)
        connection_layout.addWidget(self.connect_button)
        connection_layout.addWidget(self.reconnect_check)
        connection_layout.addWidget(self.echo_check)
        connection_layout.addWidget(self.cursor_check)
        connection_layout.addWidget(self.save_button)  # Add save button
        connection_layout.addWidget(self.clear_button)
        connection_group.setLayout(connection_layout)
        
        # Add connection group to main layout
        main_layout.addWidget(connection_group)
        
        # Serial monitor display - Use PlainTextEdit for better control over text placement
        self.monitor = QTextEdit()
        self.monitor.setReadOnly(True)
        self.monitor.setFont(QFont("Courier New", 10))
        self.monitor.setLineWrapMode(QTextEdit.NoWrap)  # Prevent line wrapping
        main_layout.addWidget(self.monitor)
        
        # Direct input area - where characters are sent immediately
        direct_input_group = QGroupBox("Direct Input (characters sent immediately, no echo)")
        direct_input_layout = QHBoxLayout()
        
        self.direct_input = DirectInputLineEdit()
        self.direct_input.setPlaceholderText("Type here to send characters silently (no echo)...")
        self.direct_input.char_typed.connect(self.send_character)
        
        # Add clear button for direct input
        self.direct_clear_button = QPushButton("Clear")
        self.direct_clear_button.clicked.connect(self.clear_direct_input)
        
        direct_input_layout.addWidget(self.direct_input)
        direct_input_layout.addWidget(self.direct_clear_button)
        direct_input_group.setLayout(direct_input_layout)
        main_layout.addWidget(direct_input_group)
        
        # Command input area
        command_layout = QHBoxLayout()
        self.command_input = QLineEdit()
        self.command_input.setPlaceholderText("Enter command...")
        self.command_input.returnPressed.connect(self.send_command)
        
        self.send_button = QPushButton("Send")
        self.send_button.clicked.connect(self.send_command)
        
        command_layout.addWidget(self.command_input)
        command_layout.addWidget(self.send_button)
        
        main_layout.addLayout(command_layout)
        
        # Set initial state
        self.set_controls_enabled(False)
        
        # Add active data indicator (blinking cursor)
        self.update_blink_cursor()
    
    def toggle_cursor_visibility(self, state):
        """Toggle cursor visibility"""
        self.show_cursor = bool(state)
        if self.show_cursor and self.is_connected:
            self.update_blink_cursor()
        else:
            self.remove_blink_cursor()
    
    def clear_monitor(self):
        """Clear the monitor display and reset line tracking"""
        self.monitor.clear()
        self.current_line_start = 0
        # Add blinking cursor at the beginning
        self.update_blink_cursor()
    
    def toggle_echo(self, state):
        """Toggle whether commands are echoed in the monitor"""
        self.show_commands = bool(state)
        
    def toggle_auto_reconnect(self, state):
        """Toggle auto-reconnect feature"""
        self.auto_reconnect = bool(state)
        
    def refresh_ports(self):
        """Update the available serial ports"""
        current_port = self.port_combo.currentText()
        
        # Get available ports
        ports = [port.device for port in serial.tools.list_ports.comports()]
        
        # Store current text if it's a custom entry
        custom_text = self.port_combo.currentText()
        
        # Clear and update the combo box
        self.port_combo.clear()
        if ports:
            self.port_combo.addItems(ports)
            # Try to set COM9 as default if available
            com9_index = self.port_combo.findText("COM9")
            if (com9_index >= 0):
                self.port_combo.setCurrentIndex(com9_index)
            # Otherwise restore previous selection if still available
            elif current_port in ports:
                self.port_combo.setCurrentText(current_port)
        
        # If there was custom text, restore it
        if custom_text and custom_text not in ports:
            self.port_combo.setCurrentText(custom_text)
    
    def check_reconnect(self):
        """Check if we need to reconnect based on port availability"""
        if not self.auto_reconnect or not self.last_port or not self.last_baud or self.manual_disconnect:
            return
            
        if not self.is_connected:
            # Check if the port is available
            ports = [port.device for port in serial.tools.list_ports.comports()]
            if self.last_port in ports:
                # Try to reconnect
                self.monitor.append(f"<span style='color:#FFA500;'>Auto-reconnect: Trying to connect to {self.last_port}...</span>")
                self.connect_to_port(self.last_port, self.last_baud)
    
    def connect_to_port(self, port, baud_rate):
        """Connect to a specific port and baud rate"""
        try:
            # Create and open serial connection with large buffers
            self.serial_connection = serial.Serial(
                port=port,
                baudrate=baud_rate,
                timeout=0,
                xonxoff=False,
                rtscts=False,
                dsrdtr=False,
                write_timeout=None
            )
            
            # Initialize the receiver
            self.serial_receiver = SerialReceiver(self.serial_connection)
            self.serial_receiver.data_received.connect(self.handle_received_data)
            self.serial_receiver.error_occurred.connect(self.handle_error)
            
            # Update UI
            self.connect_button.setText("Disconnect")
            self.connect_button.setStyleSheet("background-color: #722; color: white;")
            self.is_connected = True
            self.set_controls_enabled(True)
            
            # Save last successful connection parameters
            self.last_port = port
            self.last_baud = baud_rate
            
            # Reset manual disconnect flag on successful connection
            self.manual_disconnect = False
            
            # Start reading data
            self.read_timer.timeout.connect(self.serial_receiver.receive_data)
            self.read_timer.start(1)  # Check for new data very frequently (1ms)
            
            self.remove_blink_cursor()  # Remove cursor before connect message
            self.monitor.append(f"<span style='color:#4CAF50;'>Connected to {port} at {baud_rate} baud</span>")
            
            # Update current line start position
            cursor = self.monitor.textCursor()
            cursor.movePosition(QTextCursor.End)
            cursor.insertText('\n')
            self.current_line_start = cursor.position()
            
            # Add blinking cursor
            self.update_blink_cursor()
            
            return True
            
        except Exception as e:
            self.monitor.append(f"<span style='color:#F44336;'>Error connecting: {str(e)}</span>")
            return False
        
    def toggle_connection(self):
        """Connect or disconnect from the serial port"""
        if not self.is_connected:
            # Get port from the editable combo box
            port = self.port_combo.currentText().strip()
            
            # Get baud rate from the editable combo box
            try:
                baud_rate = int(self.baud_combo.currentText().strip())
            except ValueError:
                self.remove_blink_cursor()  # Remove cursor before error message
                self.monitor.append("<span style='color:#F44336;'>Invalid baud rate! Using 115200.</span>")
                baud_rate = 115200
            
            # Connect to the port
            self.connect_to_port(port, baud_rate)
            
        else:
            # Set the manual disconnect flag
            self.manual_disconnect = True
            
            # Disconnect
            self.read_timer.stop()
            if self.serial_connection and self.serial_connection.is_open:
                self.serial_connection.close()
            
            # Clean up
            if self.serial_receiver:
                self.serial_receiver.data_received.disconnect()
                self.serial_receiver.error_occurred.disconnect()
                self.serial_receiver = None
            
            # Update UI
            self.connect_button.setText("Connect")
            self.connect_button.setStyleSheet("")
            self.is_connected = False
            self.set_controls_enabled(False)
            
            self.remove_blink_cursor()  # Remove cursor before disconnect message
            self.monitor.append("<span style='color:#F44336;'>Disconnected</span>")
            
            # Update current line start position
            cursor = self.monitor.textCursor()
            cursor.movePosition(QTextCursor.End)
            self.current_line_start = cursor.position()
            
            # Remove blinking cursor when disconnected
            self.remove_blink_cursor()
    
    def clear_direct_input(self):
        """Clear the direct input field"""
        self.direct_input.clear()
        # Set focus back to the direct input field for convenience
        self.direct_input.setFocus()
        
    def set_controls_enabled(self, enabled):
        """Enable or disable controls based on connection state"""
        self.port_combo.setEnabled(not enabled)
        self.baud_combo.setEnabled(not enabled)
        self.command_input.setEnabled(enabled)
        self.send_button.setEnabled(enabled)
        self.direct_input.setEnabled(enabled)
        self.direct_clear_button.setEnabled(enabled)
    
    def handle_received_data(self, data):
        """Handle data received from the serial port with proper control character handling and text preservation"""
        # Update plotter if active
        if self.plotter is not None:
            # Send data to plotter if it starts with a number or minus sign
            stripped = data.strip()
            if stripped and (stripped[0].isdigit() or stripped[0] == '-'):
                self.plotter.process_data(stripped)
        
        if not data:
            return
            
        # Save data to file if saving is active
        if self.is_saving and self.save_file:
            try:
                self.save_file.write(data.encode())
                self.save_file.flush()  # Ensure data is written immediately
            except Exception as e:
                self.remove_blink_cursor()  # Remove cursor before error message
                self.monitor.append(f"<span style='color:#F44336;'>Error writing to file: {str(e)}</span>")
                self.stop_saving()
        
        # Mark data activity
        self.data_active = True
        self.last_activity_time = time.time()
        
        # Always remove blinking cursor before adding new content
        self.remove_blink_cursor()
        
        # Ensure the cursor is at the end of the document
        cursor = self.monitor.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.monitor.setTextCursor(cursor)
        
        # Get the current document
        doc = self.monitor.document()
        
        # Process each character in the received data
        i = 0
        while i < len(data):
            char = data[i]
            
            if (char == '\r'):  # Carriage Return
                # Move cursor to beginning of current line
                cursor = self.monitor.textCursor()
                block = doc.findBlock(cursor.position())
                cursor.setPosition(block.position())
                self.monitor.setTextCursor(cursor)
            elif (char == '\n'):  # Line Feed
                # Insert a new line
                cursor = self.monitor.textCursor()
                cursor.movePosition(QTextCursor.End)
                cursor.insertText('\n')
                # Update current line start
                self.current_line_start = cursor.position()
            elif (char == '\b'):  # Backspace
                # Move cursor back one and delete character
                cursor = self.monitor.textCursor()
                cursor.movePosition(QTextCursor.Left, QTextCursor.KeepAnchor)
                cursor.removeSelectedText()
                self.monitor.setTextCursor(cursor)
            elif (char == '\t'):  # Tab
                # Insert 8 spaces for tab
                cursor = self.monitor.textCursor()
                cursor.movePosition(QTextCursor.End)
                cursor.insertText('       ')  # 8 spaces to represent a tab
            else:
                # Insert the character as plain text with white color
                cursor = self.monitor.textCursor()
                cursor.movePosition(QTextCursor.End)
                
                # Create a format with white color
                text_format = cursor.charFormat()
                text_format.setForeground(QColor(255, 255, 255))  # White color
                cursor.setCharFormat(text_format)
                
                # Insert the character with the white format
                cursor.insertText(char)
            
            i += 1
        
        # Ensure the cursor/view is at the end
        scrollbar = self.monitor.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        
        # Add blinking cursor at the end
        self.update_blink_cursor()
        
    def handle_error(self, error_message):
        """Handle error messages from the serial receiver"""
        self.monitor.append(f"<span style='color:#F44336;'>{error_message}</span>")
        # Disconnect on error for safety
        self.disconnect_port()
    
    def disconnect_port(self):
        """Disconnect from the port without changing the auto-reconnect setting"""
        self.read_timer.stop()
        if self.serial_connection and self.serial_connection.is_open:
            self.serial_connection.close()
        
        # Clean up
        if self.serial_receiver:
            self.serial_receiver.data_received.disconnect()
            self.serial_receiver.error_occurred.disconnect()
            self.serial_receiver = None
        
        # Update UI
        self.connect_button.setText("Connect")
        self.connect_button.setStyleSheet("")
        self.is_connected = False
        self.set_controls_enabled(False)
        
        # Remove blinking cursor when disconnected
        self.remove_blink_cursor()
    
    def send_character(self, char):
        """Send a single character to the serial port immediately without echoing"""
        if not self.is_connected or not self.serial_connection:
            return
            
        try:
            # Send the character immediately
            self.serial_connection.write(char.encode())
            self.serial_connection.flush()  # Ensure data is sent immediately
            
            # No echo for direct input characters - they are sent silently
                
        except Exception as e:
            self.monitor.append(f"<span style='color:#F44336;'>Error sending character: {str(e)}</span>")
    
    def send_command(self):
        """Send the command to the serial port"""
        if not self.is_connected or not self.serial_connection:
            return
        
        command = self.command_input.text()
        if not command:
            return
            
        try:
            # Check if command already ends with \r\n
            # If not, add \r\n
            if not command.endswith('\r\n'):
                command_bytes = (command + '\r').encode()
            else:
                command_bytes = command.encode()
                
            # Send command
            self.serial_connection.write(command_bytes)
            self.serial_connection.flush()  # Ensure data is sent immediately
            
            # Show in monitor with a different color if echo is enabled
            if self.show_commands:
                # Temporarily remove blinking cursor if present
                self.remove_blink_cursor()
                
                # Display the command (without visible spaces for commands)
                self.monitor.append(f"<span style='color:#FFD700;'>&gt; {command+'\r'}</span>")
                
                # Update current line start position after command is displayed
                cursor = self.monitor.textCursor()
                cursor.movePosition(QTextCursor.End)
                cursor.insertText('\n')
                self.current_line_start = cursor.position()
                
                # Add blinking cursor at the end
                self.update_blink_cursor()
            
            # Clear input field
            self.command_input.clear()
            
        except Exception as e:
            self.monitor.append(f"<span style='color:#F44336;'>Error sending command: {str(e)}</span>")
            self.monitor.insertPlainText("\n")
    
    def toggle_blink_cursor(self):
        """Toggle the blinking cursor state"""
        self.blink_state = not self.blink_state
        self.update_blink_cursor()
    
    def update_blink_cursor(self):
        """Update the blinking cursor in the monitor"""
        # Remove existing cursor first
        self.remove_blink_cursor()
        
        # Only show cursor if user wants it visible and connection is active
        if self.show_cursor and self.is_connected:
            cursor = self.monitor.textCursor()
            cursor.movePosition(QTextCursor.End)
            
            # Set format for cursor/idle text
            text_format = QTextCharFormat()
            
            if self.data_active:
                # Show blinking cursor only during active data
                if self.blink_state:
                    text_format.setForeground(QColor(0, 255, 0))  # Bright green
                    cursor.setCharFormat(text_format)
                    cursor.insertText('█')
            else:
                # Show IDLE in dimmer green when no data activity
                text_format.setForeground(QColor(0, 180, 0))  # Dimmer green
                cursor.setCharFormat(text_format)
                cursor.insertText('IDLE')
            
            # Keep cursor at the end
            self.monitor.setTextCursor(cursor)
            
            # Make sure the view scrolls to the cursor
            self.monitor.ensureCursorVisible()

    def remove_blink_cursor(self):
        """Remove the blinking cursor from the monitor"""
        cursor = self.monitor.textCursor()
        cursor.movePosition(QTextCursor.End)
        
        # Check if last characters are cursor or IDLE
        document = self.monitor.document()
        last_block = document.lastBlock()
        last_text = last_block.text()[-4:] if len(last_block.text()) >= 4 else last_block.text()
        
        if last_text.endswith('█') or last_text.endswith('IDLE'):
            # If last character is cursor or IDLE, remove it
            cursor.movePosition(QTextCursor.Left, QTextCursor.KeepAnchor, 
                             1 if last_text.endswith('█') else 4)
            cursor.removeSelectedText()
            self.monitor.setTextCursor(cursor)
    
    def check_data_activity(self):
        """Check if there has been data activity recently"""
        # If data was active, but no data received in the last 2 seconds, mark as inactive
        if self.data_active and time.time() - self.last_activity_time > 2.0:  # 2 second timeout
            self.data_active = False
            # Update cursor color (don't remove it)
            self.update_blink_cursor()
    
    def toggle_save(self):
        """Toggle saving received data to file"""
        if not self.is_saving:
            # Open file dialog to select save location
            file_name, _ = QFileDialog.getSaveFileName(
                self,
                "Save Data As...",
                "",
                "Text Files (*.txt);;CSV Files (*.csv);;Binary Files (*.bin);;All Files (*.*)"
            )
            
            if file_name:
                try:
                    # Open file for appending
                    self.save_file = open(file_name, 'ab')
                    self.is_saving = True
                    self.save_button.setText("Stop Saving")
                    self.save_button.setStyleSheet("background-color: #722; color: white;")
                    
                    # Remove existing cursor, add newline, then add status message
                    self.remove_blink_cursor()
                    cursor = self.monitor.textCursor()
                    cursor.movePosition(QTextCursor.End)
                    self.monitor.append(f"<span style='color:#4CAF50;'>Started saving data to: {file_name}</span>")
                    cursor.movePosition(QTextCursor.End)
                    cursor.insertText('\n')  # Add newline after message
                    self.current_line_start = cursor.position()
                    self.update_blink_cursor()
                    
                except Exception as e:
                    self.remove_blink_cursor()
                    self.monitor.append(f"<span style='color:#F44336;'>Error opening file: {str(e)}</span>")
                    cursor = self.monitor.textCursor()
                    cursor.movePosition(QTextCursor.End)
                    cursor.insertText('\n')
                    self.current_line_start = cursor.position()
                    self.update_blink_cursor()
        else:
            # Stop saving and close file
            self.stop_saving()

    def stop_saving(self):
        """Stop saving data and close file"""
        if self.save_file:
            try:
                self.save_file.close()
                self.save_file = None
                self.is_saving = False
                self.save_button.setText("Save")
                self.save_button.setStyleSheet("")
                
                # Remove existing cursor, add newline, then add status message
                self.remove_blink_cursor()
                cursor = self.monitor.textCursor()
                cursor.movePosition(QTextCursor.End)
                self.monitor.append("<span style='color:#4CAF50;'>Stopped saving data</span>")
                cursor.movePosition(QTextCursor.End)
                cursor.insertText('\n')  # Add newline after message
                self.current_line_start = cursor.position()
                self.update_blink_cursor()
                
            except Exception as e:
                self.remove_blink_cursor()
                self.monitor.append(f"<span style='color:#F44336;'>Error closing file: {str(e)}</span>")
                cursor = self.monitor.textCursor()
                cursor.movePosition(QTextCursor.End)
                cursor.insertText('\n')
                self.current_line_start = cursor.position()
                self.update_blink_cursor()

    def toggle_plotter(self):
        """Toggle the plotter window"""
        if self.plotter is None:
            self.plotter = SerialPlotter()
            self.plotter.show()
        else:
            self.plotter.close()
            self.plotter = None

    def closeEvent(self, event):
        """Clean up resources when closing the application"""
        # Stop saving if active
        if self.is_saving:
            self.stop_saving()
            
        # Stop timers
        self.read_timer.stop()
        self.port_refresh_timer.stop()
        self.reconnect_timer.stop()
        self.blink_timer.stop()
        self.data_activity_timer.stop()
        
        # Close connection
        if self.serial_connection and self.serial_connection.is_open:
            self.serial_connection.close()
        
        # Close plotter if open
        if self.plotter is not None:
            self.plotter.close()
        
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SerialMonitor()
    window.show()
    sys.exit(app.exec_())
    