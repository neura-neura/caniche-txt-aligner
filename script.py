#!/usr/bin/env python3
"""
Caniche TXT Aligner - PYQT6 application to edit text files in different languages
It allows to modify the original files directly.
With load on demand to improve performance with large files.

Instructions:
1. Install dependencies: Pip Install Pyqt6
2. Run this script: python script.py

Command to export to EXE:
pyinstaller --onefile --windowed --icon=assets/img/icon.ico --add-data "assets/img/icon.ico;assets/img" script.py
"""

import sys
import os
import time
import ctypes
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
                           QLabel, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
                           QComboBox, QLineEdit, QMessageBox, QHeaderView, QFrame, QSplitter,
                           QStatusBar, QToolBar, QMenu, QDialog, QProgressDialog, QInputDialog,
                           QProgressBar, QTextEdit, QStyledItemDelegate, QAbstractItemView, 
                           QScrollBar)
from PyQt6.QtCore import Qt, QSize, pyqtSignal, QTimer, QThread, QObject, QModelIndex
from PyQt6.QtGui import QAction, QIcon, QColor, QFont, QTextOption, QPixmap


# Define la ruta al icono de manera más robusta
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)


# Obtener ruta absoluta al ícono
ICON_PATH = os.path.abspath(resource_path("assets/img/icon.ico"))


# Define the delegated class for first text edition
class TextEditDelegate(QStyledItemDelegate):
    """Personalized delegate that uses Qtextedit to edit multiline text in cells"""
    def createEditor(self, parent, option, index):
        editor = QTextEdit(parent)
        editor.setAcceptRichText(False)
        editor.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        return editor
    
    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.ItemDataRole.DisplayRole)
        editor.setText(value)
    
    def setModelData(self, editor, model, index):
        model.setData(index, editor.toPlainText(), Qt.ItemDataRole.EditRole)
    
    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)


class LineNumberDelegate(QWidget):
    """Custom widget for line numbers"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(5, 0, 5, 0)
        self.layout.setSpacing(0)
        
        self.label = QLabel("0")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = QFont()
        font.setBold(True)
        self.label.setFont(font)
        
        self.layout.addWidget(self.label)
    
    def setNumber(self, number):
        self.label.setText(str(number))


class FileLoaderWorker(QObject):
    """Working class to load files in a separate thread"""
    progress_updated = pyqtSignal(int, int, float)  # loaded lines, total, percentage
    finished = pyqtSignal(list)  # file content
    error = pyqtSignal(str)
    
    def __init__(self, filepath):
        super().__init__()
        self.filepath = filepath
        self.is_canceled = False
    
    def load_file(self):
        """Load the line file by line with progress update"""
        try:
            content = []
            
            # First count the total number of lines
            total_lines = 0
            with open(self.filepath, 'r', encoding='utf-8') as f:
                for _ in f:
                    total_lines += 1
            
            # Then read the content with progress updates
            with open(self.filepath, 'r', encoding='utf-8') as f:
                for i, line in enumerate(f):
                    if self.is_canceled:
                        break
                    
                    content.append(line.rstrip('\n'))
                    
                    # Update progress every 100 lines or on the last line
                    if (i % 100 == 0) or (i == total_lines - 1):
                        percentage = (i + 1) / total_lines * 100
                        self.progress_updated.emit(i + 1, total_lines, percentage)
            
            self.finished.emit(content)
        
        except Exception as e:
            self.error.emit(str(e))
    
    def cancel(self):
        """Cancel the file loading"""
        self.is_canceled = True


class ProgressDialog(QDialog):
    """Dialogue with Progress Bar to show the load advance"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Progress")
        self.resize(400, 150)
        self.setModal(True)
        
        # Establecer el ícono de la ventana
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))
        
        # Main layout
        layout = QVBoxLayout(self)
        
        # Tag for descriptive text
        self.label = QLabel("Loading file ...")
        layout.addWidget(self.label)
        
        # Tag to show processed/total lines
        self.lines_label = QLabel("Processing lines: 0 / 0")
        layout.addWidget(self.lines_label)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # Tag for percentage
        self.percentage_label = QLabel("0.00%")
        self.percentage_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.percentage_label)
        
        # Cancel button
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        layout.addWidget(self.cancel_button)
    
    def setLabelText(self, text):
        """Establish the text of the main label"""
        self.label.setText(text)
    
    def updateProgress(self, current, total, percentage):
        """Update the progress bar and labels"""
        self.lines_label.setText(f"Processing lines: {current:,} / {total:,}")
        self.progress_bar.setValue(int(percentage))
        self.percentage_label.setText(f"{percentage:.2f}%")
        QApplication.processEvents()  # Ensure that the UI is updated
    
    def closeEvent(self, event):
        """Ensure that dialogue is closed correctly"""
        self.deleteLater()
        super().closeEvent(event)


class VirtualTableWidget(QTableWidget):
    """Virtual table widget with rows on demand"""
    rowModified = pyqtSignal(int)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        # Complete data in memory
        self.file1_content = []
        self.file2_content = []
        self.total_rows = 0
        
        # Visualization configuration
        self.visible_rows_buffer = 50  # Number of additional rows to load above and below
        self.modified_rows = set()
        
        # Set the table
        self.setColumnCount(3)  # Line number, file 1, file 2
        self.setHorizontalHeaderLabels(["#", "File 1", "File 2"])
        
        # Table settings
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.setColumnWidth(0, 50)
        
        # Settings to display multiline text
        self.setWordWrap(True)
        self.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft)
        
        # Automatic rows height adjustment
        self.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        
        # Configure the personalized editing delegate
        self.text_edit_delegate = TextEditDelegate()
        self.setItemDelegateForColumn(1, self.text_edit_delegate)
        self.setItemDelegateForColumn(2, self.text_edit_delegate)
        
        # Configure editing mode
        self.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | 
                             QAbstractItemView.EditTrigger.SelectedClicked | 
                             QAbstractItemView.EditTrigger.EditKeyPressed)
        
        # Allow direct editing in cells
        self.cellChanged.connect(self.handleCellChange)
        
        # Connect Scroll signal
        self.verticalScrollBar().valueChanged.connect(self.handleScroll)
        
        # Timer to optimize the load during scroll
        self.scroll_timer = QTimer(self)
        self.scroll_timer.setSingleShot(True)
        self.scroll_timer.timeout.connect(self.loadVisibleRows)
    
    def handleCellChange(self, row, column):
        # Just register changes in content columns (1 and 2)
        if column > 0:
            self.modified_rows.add(row)
            self.rowModified.emit(row)
            
            # Update memory content
            if column == 1 and row < len(self.file1_content):
                self.file1_content[row] = self.item(row, column).text()
            elif column == 2 and row < len(self.file2_content):
                self.file2_content[row] = self.item(row, column).text()
            
            # Automatically adjust the height of the row
            self.resizeRowToContents(row)
    
    def handleScroll(self, value):
        # Use a timer to avoid loading ranks during the fast scroll
        self.scroll_timer.start(100)
    
    def visibleRowRange(self):
        """Obtain the rank of current rows"""
        rect = self.viewport().rect()
        first_visible_row = self.rowAt(rect.top())
        last_visible_row = self.rowAt(rect.bottom())
        
        # If there are no visible rows, return None
        if first_visible_row == -1:
            first_visible_row = 0
        if last_visible_row == -1:
            if self.rowCount() > 0:
                last_visible_row = self.rowCount() - 1
            else:
                last_visible_row = 0
        
        return first_visible_row, last_visible_row
    
    def loadVisibleRows(self):
        """Load the visible ranks and an additional buffer"""
        if self.total_rows == 0:
            return
        
        # Obtain visible rows range
        first_visible, last_visible = self.visibleRowRange()
        
        # Calculate the ranks to load with buffer
        start_row = max(0, first_visible - self.visible_rows_buffer)
        end_row = min(self.total_rows - 1, last_visible + self.visible_rows_buffer)
        
        # Verify if there are rows that need to be created
        current_rows = self.rowCount()
        if end_row >= current_rows:
            self.setRowCount(end_row + 1)
            
            # Create the necessary rows
            for row in range(current_rows, end_row + 1):
                # Just create the ranks, the content will be filled later
                number_widget = LineNumberDelegate()
                number_widget.setNumber(row + 1)
                self.setCellWidget(row, 0, number_widget)
        
        # Load content for visible ranks and buffer
        self.loadContentForRows(start_row, end_row)
    
    def loadContentForRows(self, start_row, end_row):
        """Load the content for a specific range of rows"""
        for row in range(start_row, end_row + 1):
            # Verify if the cells are already created and have content
            for col in range(1, 3):
                item = self.item(row, col)
                
                # If the item does not exist, create it with the corresponding content
                if item is None:
                    content = ""
                    if col == 1 and row < len(self.file1_content):
                        content = self.file1_content[row]
                    elif col == 2 and row < len(self.file2_content):
                        content = self.file2_content[row]
                    
                    item = self.createCustomTableItem(content)
                    self.setItem(row, col, item)
                # If the item exists but it is empty and should have content, update it
                elif item.text() == "":
                    if col == 1 and row < len(self.file1_content) and self.file1_content[row] != "":
                        item.setText(self.file1_content[row])
                    elif col == 2 and row < len(self.file2_content) and self.file2_content[row] != "":
                        item.setText(self.file2_content[row])
    
    def createCustomTableItem(self, text=""):
        """Create an item of table with multiline text configuration"""
        item = QTableWidgetItem(text)
        item.setTextAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        return item
    
    def setContentFromFile(self, column, content):
        """Establish the content of a memory file"""
        if column == 1:
            self.file1_content = content.copy()
        elif column == 2:
            self.file2_content = content.copy()
        
        # Update the total rows
        self.total_rows = max(len(self.file1_content), len(self.file2_content))
        
        # Load visible initial rows
        self.loadVisibleRows()
        
        # Force an immediate and complete update for visible rows
        QTimer.singleShot(100, self.forceVisibleUpdate)
    
    def equalizeFiles(self):
        """Match the content of both files so that they have the same number of lines"""
        max_rows = self.total_rows
        
        # Extend the shortest file with empty lines
        if len(self.file1_content) < max_rows:
            self.file1_content.extend([""] * (max_rows - len(self.file1_content)))
        
        if len(self.file2_content) < max_rows:
            self.file2_content.extend([""] * (max_rows - len(self.file2_content)))
        
        # Force update after matching
        QTimer.singleShot(100, self.forceVisibleUpdate)
    
    def moveRow(self, row_index, direction):
        """Move a row up or down"""
        target_row = row_index - 1 if direction == 'up' else row_index + 1
        
        # Verify limits
        if target_row < 0 or target_row >= self.total_rows:
            return False
        
        # Exchange content in memory
        # File 1
        if row_index < len(self.file1_content) and target_row < len(self.file1_content):
            self.file1_content[row_index], self.file1_content[target_row] = \
                self.file1_content[target_row], self.file1_content[row_index]
        
        # File 2
        if row_index < len(self.file2_content) and target_row < len(self.file2_content):
            self.file2_content[row_index], self.file2_content[target_row] = \
                self.file2_content[target_row], self.file2_content[row_index]
        
        # Exchange content of visible cells
        for col in range(1, self.columnCount()):
            current_item = self.item(row_index, col)
            target_item = self.item(target_row, col)
            
            if current_item and target_item:
                current_text = current_item.text()
                target_text = target_item.text()
                
                current_item.setText(target_text)
                target_item.setText(current_text)
        
        # Mark both rows as modified
        self.modified_rows.add(row_index)
        self.modified_rows.add(target_row)
        self.rowModified.emit(row_index)
        self.rowModified.emit(target_row)
        
        # Adjust rows height
        self.resizeRowToContents(row_index)
        self.resizeRowToContents(target_row)
        
        # Select the row that moved
        self.selectRow(target_row)
        
        return True
    
    def deleteRow(self, row_index):
        """Delete a row"""
        if row_index < 0 or row_index >= self.total_rows:
            return False
        
        # Remove the row of the memory
        if row_index < len(self.file1_content):
            self.file1_content.pop(row_index)
        
        if row_index < len(self.file2_content):
            self.file2_content.pop(row_index)
        
        # Update the total rows
        self.total_rows -= 1
        
        # Delete the row from the table
        self.removeRow(row_index)
        
        # Update the modified rows
        new_modified_rows = set()
        for index in self.modified_rows:
            if index < row_index:
                new_modified_rows.add(index)
            elif index > row_index:
                new_modified_rows.add(index - 1)
        
        self.modified_rows = new_modified_rows
        
        # Update the line numbers
        self.updateNumberColumn()
        
        return True
    
    def addEmptyRow(self):
        """Add an empty row to the end"""
        # Add an empty line to memory data
        self.file1_content.append("")
        self.file2_content.append("")
        
        # Update the total rows
        self.total_rows += 1
        
        # Verify if the new row is within the visible range
        first_visible, last_visible = self.visibleRowRange()
        if self.total_rows - 1 <= last_visible + self.visible_rows_buffer:
            # Add the row to the table
            row_count = self.rowCount()
            self.insertRow(row_count)
            
            # Add line number
            number_widget = LineNumberDelegate()
            number_widget.setNumber(row_count + 1)
            self.setCellWidget(row_count, 0, number_widget)
            
            # Add empty editable cells
            self.setItem(row_count, 1, self.createCustomTableItem(""))
            self.setItem(row_count, 2, self.createCustomTableItem(""))
            
            # Adjust height automatically
            self.resizeRowToContents(row_count)
            
            # Select the new row
            self.selectRow(row_count)
        
        return self.total_rows - 1  # Return the index of the new row
    
    def getContent(self, column):
        """Obtain the content of a specific column as a list of lines"""
        if column == 1:
            return self.file1_content.copy()
        elif column == 2:
            return self.file2_content.copy()
        return []
    
    def forceVisibleUpdate(self):
        """Force the update of all visible cells"""
        first_visible, last_visible = self.visibleRowRange()
        
        # Walk all visible ranks
        for row in range(first_visible, last_visible + 1):
            if row < self.rowCount():
                # Update the cells in columns 1 and 2
                for col in range(1, 3):
                    item = self.item(row, col)
                    if item:
                        # Get the right content
                        content = ""
                        if col == 1 and row < len(self.file1_content):
                            content = self.file1_content[row]
                        elif col == 2 and row < len(self.file2_content):
                            content = self.file2_content[row]
                            
                        # Update the content
                        if item.text() != content:
                            item.setText(content)
        
        # Necessary to force repainting
        self.viewport().update()
        QApplication.processEvents()
        
    def highlightSearchResults(self, search_term):
        """Highlight the rows that contain the search term"""
        if not search_term:
            # If the search term is empty, clean all those highlighted
            for row in range(self.rowCount()):
                for col in range(1, self.columnCount()):
                    item = self.item(row, col)
                    if item:
                        item.setBackground(QColor(255, 255, 255))
            return 0
        
        search_term = search_term.lower()
        matches = 0
        
        # Search in memory data
        matching_rows = []
        
        # Search in the content of file 1
        for i, line in enumerate(self.file1_content):
            if search_term in line.lower():
                matching_rows.append(i)
                matches += 1
        
        # Search in the content of file 2
        for i, line in enumerate(self.file2_content):
            if i not in matching_rows and search_term in line.lower():
                matching_rows.append(i)
                matches += 1
        
        # Highlight the visible rows that coincide
        for row in range(self.rowCount()):
            if row in matching_rows:
                for col in range(1, self.columnCount()):
                    item = self.item(row, col)
                    if item:
                        item.setBackground(QColor(255, 255, 160))  # Light yellow
            else:
                for col in range(1, self.columnCount()):
                    item = self.item(row, col)
                    if item:
                        item.setBackground(QColor(255, 255, 255))  # White
        
        # If there are coincidences that are not visible, make sure to load at least the first
        if matching_rows and not any(row < self.rowCount() for row in matching_rows):
            # Scroll to the first coincidence
            first_match = min(matching_rows)
            self.setCurrentCell(first_match, 1)
            
            # Load rows around the first coincidence
            self.loadContentForRows(max(0, first_match - self.visible_rows_buffer), 
                                   min(self.total_rows - 1, first_match + self.visible_rows_buffer))
            
            # Highlight the first coincidence
            for col in range(1, self.columnCount()):
                item = self.item(first_match, col)
                if item:
                    item.setBackground(QColor(255, 255, 160))  # Amarillo claro
        
        return matches
    
    def countWords(self, column):
        """Count words in the specified column"""   
        content = []
        if column == 1:
            content = self.file1_content
        elif column == 2:
            content = self.file2_content
        
        words = ' '.join(content).split()
        return len(words)
    
    def updateNumberColumn(self):
        """Update the line numbers"""
        for row in range(self.rowCount()):
            number_widget = self.cellWidget(row, 0)
            if not number_widget:
                number_widget = LineNumberDelegate()
                self.setCellWidget(row, 0, number_widget)
            number_widget.setNumber(row + 1)


class MainWindow(QMainWindow):
    """Main application window"""
    def __init__(self):
        super().__init__()
        
        self.file1_path = ""
        self.file2_path = ""
        self.file1_language = ""
        self.file2_language = ""
        
        # Establecer el ícono de la aplicación para la ventana principal
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))
        
        # Referencias to threads and working objects
        self.file_loader_thread = None
        self.file_loader = None
        
        self.setupUi()
        self.connectSignals()
    
    def setupUi(self):
        """Configure the user interface"""
        self.setWindowTitle("Caniche TXT Aligner (BETA v0.1)")
        self.setMinimumSize(1000, 700)
        
        # Central widget
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        
        # File settings
        file_layout = QHBoxLayout()
        
        # File 1
        file1_frame = QFrame()
        file1_frame.setFrameShape(QFrame.Shape.StyledPanel)
        file1_layout = QVBoxLayout(file1_frame)
        
        file1_header = QLabel("File 1")
        file1_header.setStyleSheet("font-weight: bold; font-size: 14px;")
        file1_layout.addWidget(file1_header)
        
        file1_select_layout = QHBoxLayout()
        self.file1_path_label = QLabel("There is no selected file")
        self.file1_path_label.setWordWrap(True)
        file1_select_button = QPushButton("Select")
        file1_select_button.clicked.connect(lambda: self.selectFile(1))
        
        file1_select_layout.addWidget(self.file1_path_label)
        file1_select_layout.addWidget(file1_select_button)
        file1_layout.addLayout(file1_select_layout)
        
        # Language selector 1
        language1_layout = QHBoxLayout()
        language1_layout.addWidget(QLabel("Language:"))
        self.language1_combo = QComboBox()
        self.setupLanguageCombo(self.language1_combo)
        language1_layout.addWidget(self.language1_combo)
        file1_layout.addLayout(language1_layout)
        
        # File 2
        file2_frame = QFrame()
        file2_frame.setFrameShape(QFrame.Shape.StyledPanel)
        file2_layout = QVBoxLayout(file2_frame)
        
        file2_header = QLabel("File 2")
        file2_header.setStyleSheet("font-weight: bold; font-size: 14px;")
        file2_layout.addWidget(file2_header)
        
        file2_select_layout = QHBoxLayout()
        self.file2_path_label = QLabel("There is no selected file")
        self.file2_path_label.setWordWrap(True)
        file2_select_button = QPushButton("Select")
        file2_select_button.clicked.connect(lambda: self.selectFile(2))
        
        file2_select_layout.addWidget(self.file2_path_label)
        file2_select_layout.addWidget(file2_select_button)
        file2_layout.addLayout(file2_select_layout)
        
        # Language selector 2
        language2_layout = QHBoxLayout()
        language2_layout.addWidget(QLabel("Language:"))
        self.language2_combo = QComboBox()
        self.setupLanguageCombo(self.language2_combo)
        language2_layout.addWidget(self.language2_combo)
        file2_layout.addLayout(language2_layout)
        
        # Add frames to the main layout
        file_layout.addWidget(file1_frame)
        file_layout.addWidget(file2_frame)
        main_layout.addLayout(file_layout)
        
        # Search
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Search:"))
        self.search_input = QLineEdit()
        self.search_input.returnPressed.connect(self.searchText)
        search_layout.addWidget(self.search_input)
        
        search_button = QPushButton("Search")
        search_button.clicked.connect(self.searchText)
        search_layout.addWidget(search_button)
        
        clear_search_button = QPushButton("Clean")
        clear_search_button.clicked.connect(self.clearSearch)
        search_layout.addWidget(clear_search_button)
        
        main_layout.addLayout(search_layout)
        
        # Control buttons
        control_layout = QHBoxLayout()
        
        add_row_button = QPushButton("Add line")
        add_row_button.clicked.connect(self.addRow)
        control_layout.addWidget(add_row_button)
        
        save_button = QPushButton("Save changes")
        save_button.setStyleSheet("background-color: #2a9d8f; color: white;")
        save_button.clicked.connect(self.saveChanges)
        control_layout.addWidget(save_button)
        
        download_button = QPushButton("Save as ...")
        download_button.clicked.connect(self.saveAs)
        control_layout.addWidget(download_button)
        
        reload_button = QPushButton("Reload files")
        reload_button.clicked.connect(self.reloadFiles)
        control_layout.addWidget(reload_button)
        
        main_layout.addLayout(control_layout)
        
        # Edition table with virtual load
        self.table = VirtualTableWidget()
        self.table.rowModified.connect(self.updateStats)
        
        # Assign context menu to the table
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.showTableContextMenu)
        
        main_layout.addWidget(self.table)
        
        # Statistics
        stats_layout = QHBoxLayout()
        
        # Total lines
        total_lines_frame = QFrame()
        total_lines_frame.setFrameShape(QFrame.Shape.StyledPanel)
        total_lines_layout = QVBoxLayout(total_lines_frame)
        total_lines_layout.addWidget(QLabel("Total lines"))
        self.total_lines_label = QLabel("0")
        self.total_lines_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        self.total_lines_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        total_lines_layout.addWidget(self.total_lines_label)
        
        # Words File 1
        words1_frame = QFrame()
        words1_frame.setFrameShape(QFrame.Shape.StyledPanel)
        words1_layout = QVBoxLayout(words1_frame)
        words1_layout.addWidget(QLabel("Words File 1"))
        self.words1_label = QLabel("0")
        self.words1_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        self.words1_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        words1_layout.addWidget(self.words1_label)
        
        # Words File 2
        words2_frame = QFrame()
        words2_frame.setFrameShape(QFrame.Shape.StyledPanel)
        words2_layout = QVBoxLayout(words2_frame)
        words2_layout.addWidget(QLabel("Words File 2"))
        self.words2_label = QLabel("0")
        self.words2_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        self.words2_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        words2_layout.addWidget(self.words2_label)
        
        # Modified lines
        modified_lines_frame = QFrame()
        modified_lines_frame.setFrameShape(QFrame.Shape.StyledPanel)
        modified_lines_layout = QVBoxLayout(modified_lines_frame)
        modified_lines_layout.addWidget(QLabel("Modified lines"))
        self.modified_lines_label = QLabel("0")
        self.modified_lines_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        self.modified_lines_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        modified_lines_layout.addWidget(self.modified_lines_label)
        
        stats_layout.addWidget(total_lines_frame)
        stats_layout.addWidget(words1_frame)
        stats_layout.addWidget(words2_frame)
        stats_layout.addWidget(modified_lines_frame)
        
        main_layout.addLayout(stats_layout)
        
        # Status bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        self.setCentralWidget(central_widget)
    
    def setupLanguageCombo(self, combo):
        """Configure language options in combobox"""
        languages = [
            ("", "Select Language"),
            ("es", "Spanish"),
            ("en", "English"),
            ("fr", "French"),
            ("de", "German"),
            ("it", "Italian"),
            ("pt", "Portuguese"),
            ("ru", "Russian"),
            ("zh", "Chinese"),
            ("ja", "Japanese"),
            ("ko", "Korean"),
            ("ar", "Arabic"),
            ("other", "Other"),
        ]
        
        for code, name in languages:
            combo.addItem(name, code)
    
    def connectSignals(self):
        """Connect signals and slots"""
        self.language1_combo.currentIndexChanged.connect(self.updateTableHeaders)
        self.language2_combo.currentIndexChanged.connect(self.updateTableHeaders)
    
    def selectFile(self, file_num):
        """Select a file to edit"""
        options = QFileDialog.Option.ReadOnly
        filepath, _ = QFileDialog.getOpenFileName(
            self, f"Select file {file_num}", "", 
            "Text files (*.txt);;All files (*)", 
            options=options
        )
        
        if filepath:
            if file_num == 1:
                self.file1_path = filepath
                self.file1_path_label.setText(filepath)
                self.loadFile(1, filepath)
            else:
                self.file2_path = filepath
                self.file2_path_label.setText(filepath)
                self.loadFile(2, filepath)
    
    def loadFile(self, file_num, filepath):
        """Load the content of a file using a separate thread with progress bar"""
        try:
            # Create the file load thread
            self.file_loader_thread = QThread()
            self.file_loader = FileLoaderWorker(filepath)
            self.file_loader.moveToThread(self.file_loader_thread)
            
            # Configure the progress dialog
            progress_dialog = ProgressDialog(self)
            progress_dialog.setWindowTitle(f"Loading file {file_num}")
            progress_dialog.setLabelText(f"Loading {os.path.basename(filepath)}")
            
            # Connect signals
            self.file_loader.progress_updated.connect(progress_dialog.updateProgress)
            self.file_loader.finished.connect(lambda content: self.fileLoadFinished(file_num, filepath, content))
            self.file_loader.finished.connect(self.file_loader_thread.quit)
            self.file_loader.finished.connect(lambda: progress_dialog.accept())
            self.file_loader.error.connect(lambda error: self.fileLoadError(error, progress_dialog))
            self.file_loader_thread.started.connect(self.file_loader.load_file)
            
            # Start the thread and show the dialogue
            self.file_loader_thread.start()
            
            # Execute the dialogue
            progress_dialog.exec()
            
            # Clean resources
            self.file_loader_thread.quit()
            self.file_loader_thread.wait()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"The file load could not be started: {str(e)}")
    
    def fileLoadFinished(self, file_num, filepath, content):
        """Process the loaded file content"""
        try:
            # Stablish the content in the table
            self.table.setContentFromFile(file_num, content)
            
            # If both files are loaded, make sure they have the same number of rows
            if self.file1_path and self.file2_path:
                self.table.equalizeFiles()
                
            # Force an additional update after a brief delay
            QTimer.singleShot(200, self.table.forceVisibleUpdate)
            
            self.updateTableHeaders()
            self.updateStats()
            self.showStatusMessage(f"File {file_num} loaded: {os.path.basename(filepath)}")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error when processing the loaded file: {str(e)}")
    
    def fileLoadError(self, error, progress_dialog):
        """Handle errors during file load"""
        progress_dialog.reject()
        progress_dialog.close()
        progress_dialog.deleteLater()
        QMessageBox.critical(self, "Error", f"The file could not be loaded: {str(error)}")
    
    def updateTableHeaders(self):
        """Update the table headers according to the selected languages"""
        file1_name = os.path.basename(self.file1_path) if self.file1_path else "File 1"
        file2_name = os.path.basename(self.file2_path) if self.file2_path else "File 2"
        
        # Obtain the texts of the selected languages
        lang1_idx = self.language1_combo.currentIndex()
        lang2_idx = self.language2_combo.currentIndex()
        
        lang1_text = self.language1_combo.itemText(lang1_idx) if lang1_idx > 0 else ""
        lang2_text = self.language2_combo.itemText(lang2_idx) if lang2_idx > 0 else ""
        
        # Save language codes
        self.file1_language = self.language1_combo.itemData(lang1_idx)
        self.file2_language = self.language2_combo.itemData(lang2_idx)
        
        # Update the headers
        if lang1_text:
            file1_name += f" ({lang1_text})"
        
        if lang2_text:
            file2_name += f" ({lang2_text})"
        
        self.table.setHorizontalHeaderLabels(["#", file1_name, file2_name])
    
    def showTableContextMenu(self, position):
        """Show the context menu in the table"""
        menu = QMenu()
        
        # Establecer el ícono al menú de contexto
        if os.path.exists(ICON_PATH):
            menu.setWindowIcon(QIcon(ICON_PATH))
        
        row = self.table.rowAt(position.y())
        if row >= 0:
            move_up_action = QAction("Move up", self)
            move_up_action.triggered.connect(lambda: self.moveRow(row, 'up'))
            menu.addAction(move_up_action)
            
            move_down_action = QAction("Move down", self)
            move_down_action.triggered.connect(lambda: self.moveRow(row, 'down'))
            menu.addAction(move_down_action)
            
            menu.addSeparator()
            
            delete_action = QAction("Remove row", self)
            delete_action.triggered.connect(lambda: self.deleteRow(row))
            menu.addAction(delete_action)
        
        add_action = QAction("Add row", self)
        add_action.triggered.connect(self.addRow)
        menu.addAction(add_action)
        
        menu.exec(self.table.mapToGlobal(position))
    
    def moveRow(self, row_index, direction):
        """Move a row up or down"""
        if self.table.moveRow(row_index, direction):
            direction_text = "up" if direction == 'up' else "down"
            self.showStatusMessage(f"Row moved to {direction_text}")
            self.updateStats()
    
    def deleteRow(self, row_index):
        """Delete a row"""
        confirm = QMessageBox.question(
            self, "Confirm", "Are you sure you want to delete this row?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if confirm == QMessageBox.StandardButton.Yes:
            if self.table.deleteRow(row_index):
                self.showStatusMessage("Row deleted")
                self.updateStats()
    
    def addRow(self):
        """Add a new row at the end of the table"""
        new_row = self.table.addEmptyRow()
        self.showStatusMessage("New line added")
        self.updateStats()
        
        # Make sure the new row is visible
        if new_row < self.table.rowCount():
            self.table.scrollToItem(self.table.item(new_row, 1))
    
    def saveChanges(self):
        """Save changes in original files with progress bar"""
        if not (self.file1_path or self.file2_path):
            QMessageBox.warning(self, "Warning", "There are no files to save.")
            return
        
        try:
            files_to_save = []
            if self.file1_path:
                files_to_save.append((self.file1_path, 1))
            if self.file2_path:
                files_to_save.append((self.file2_path, 2))
            
            # Create a progress bar
            total_files = len(files_to_save)
            progress = QProgressDialog("Saving files ...", "Cancel", 0, total_files, self)
            progress.setWindowTitle("Saving changes")
            progress.setWindowModality(Qt.WindowModality.WindowModal)
            
            # Establecer el ícono para el diálogo de progreso
            if os.path.exists(ICON_PATH):
                progress.setWindowIcon(QIcon(ICON_PATH))
            
            saved_files = []
            
            for i, (file_path, column) in enumerate(files_to_save):
                progress.setValue(i)
                progress.setLabelText(f"Saving {os.path.basename(file_path)}...")
                
                if progress.wasCanceled():
                    break
                
                content = self.table.getContent(column)
                total_lines = len(content)
                
                # If the file is large, show detailed progress
                if total_lines > 10000:
                    save_progress = QProgressDialog(
                        f"Saving {os.path.basename(file_path)}...", 
                        "Cancel", 0, total_lines, self)
                    save_progress.setWindowTitle(f"Saving {os.path.basename(file_path)}")
                    save_progress.setWindowModality(Qt.WindowModality.WindowModal)
                    
                    # Establecer ícono para el diálogo de progreso
                    if os.path.exists(ICON_PATH):
                        save_progress.setWindowIcon(QIcon(ICON_PATH))
                    
                    # Write lines in batch for better performance and update progress
                    with open(file_path, 'w', encoding='utf-8') as f:
                        batch_size = 1000
                        for j in range(0, total_lines, batch_size):
                            if save_progress.wasCanceled():
                                # If it is canceled, close the file and exit
                                break
                            
                            end_idx = min(j + batch_size, total_lines)
                            f.write('\n'.join(content[j:end_idx]))
                            if j + batch_size < total_lines:
                                f.write('\n')  # Add line jump between batches
                            
                            save_progress.setValue(end_idx)
                            QApplication.processEvents()
                    
                    save_progress.setValue(total_lines)
                    save_progress.close()
                    save_progress.deleteLater()
                else:
                    # For small files, save directly
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write('\n'.join(content))
                
                saved_files.append(os.path.basename(file_path))
            
            progress.setValue(total_files)
            progress.close()
            progress.deleteLater()
            
            # Clean the set of modified rows
            self.table.modified_rows.clear()
            self.updateStats()
            
            # Success message
            files_str = " y ".join(saved_files)
            QMessageBox.information(self, "Success", f"Changes saved in {files_str}")
            self.showStatusMessage(f"Changes saved in {files_str}")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error saving the changes: {str(e)}")
    
    def saveAs(self):
        """Save as a new file"""
        if self.table.total_rows == 0:
            QMessageBox.warning(self, "Warning", "There is no content to save.")
            return
        
        # Determine which file save
        file_options = []
        if self.file1_path:
            file1_name = os.path.basename(self.file1_path)
            file_options.append((1, f"File 1 ({file1_name})"))
        else:
            file_options.append((1, "File 1"))
        
        if self.file2_path:
            file2_name = os.path.basename(self.file2_path)
            file_options.append((2, f"File 2 ({file2_name})"))
        else:
            file_options.append((2, "File 2"))
        
        # If there is only one file with content, save it directly
        selected_column = None
        if len(file_options) == 1 or (self.file1_path and not self.file2_path):
            selected_column = 1
        elif not self.file1_path and self.file2_path:
            selected_column = 2
        else:
            # Ask which file save
            options = [opt[1] for opt in file_options]
            option, ok = QInputDialog.getItem(
                self, "Save as", "Select the file to save:",
                options, 0, False
            )
            
            if ok and option:
                for col, name in file_options:
                    if name == option:
                        selected_column = col
                        break
        
        if selected_column:
            options = QFileDialog.Option.ReadOnly
            filepath, _ = QFileDialog.getSaveFileName(
                self, "Save as", "", 
                "Text files (*.txt);;All files (*)",
                options=options
            )
            
            if filepath:
                try:
                    content = self.table.getContent(selected_column)
                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.write('\n'.join(content))
                    
                    QMessageBox.information(
                        self, "Success", 
                        f"Saved file as: {os.path.basename(filepath)}"
                    )
                    self.showStatusMessage(f"Saved file as: {os.path.basename(filepath)}")
                
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Error saving the file: {str(e)}")
    
    def reloadFiles(self):
        """Reload the original files"""
        if not (self.file1_path or self.file2_path):
            QMessageBox.warning(self, "Warning", "There are no files to reload.")
            return
        
        confirm = QMessageBox.question(
            self, "Confirm", "Are you sure? All unsaved changes will be lost.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if confirm == QMessageBox.StandardButton.Yes:
            # Clean the table
            self.table.setRowCount(0)
            self.table.file1_content.clear()
            self.table.file2_content.clear()
            self.table.total_rows = 0
            
            # Reload files
            if self.file1_path:
                self.loadFile(1, self.file1_path)
            
            if self.file2_path:
                self.loadFile(2, self.file2_path)
            
            # Clean modified rows
            self.table.modified_rows.clear()
            self.updateStats()
            self.showStatusMessage("Reloaded files")
    
    def searchText(self):
        """Search text in the table"""
        search_term = self.search_input.text().strip()
        
        if not search_term:
            QMessageBox.warning(self, "Warning", "Enter a search term.")
            return
        
        matches = self.table.highlightSearchResults(search_term)
        
        if matches > 0:
            self.showStatusMessage(f"Found {matches} coincidences for '{search_term}'")
        else:
            self.showStatusMessage(f"No coincidences were found for '{search_term}'")
    
    def clearSearch(self):
        """Clean the search"""
        self.search_input.clear()
        self.table.highlightSearchResults("")
        self.showStatusMessage("Cleaned search")
    
    def updateStats(self):
        """Update statistics"""
        total_lines = self.table.total_rows
        words_file1 = self.table.countWords(1)
        words_file2 = self.table.countWords(2)
        modified_lines = len(self.table.modified_rows)
        
        self.total_lines_label.setText(str(total_lines))
        self.words1_label.setText(str(words_file1))
        self.words2_label.setText(str(words_file2))
        self.modified_lines_label.setText(str(modified_lines))
    
    def showStatusMessage(self, message, timeout=5000):
        """Show a message in the status bar"""
        self.statusBar.showMessage(message, timeout)


# Specific function for Windows that explicitly establishes the application identifier
def set_windows_app_id():
    """Set the app id for Windows 7+ taskbar"""
    if sys.platform == 'win32':
        try:
            # When compiled with Pyinstaller, this helps Windows to associate
            # correctly the icon with the application in the taskbar
            app_id = "caniche.txtaligner.app.0.1"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        except Exception:
            pass


if __name__ == '__main__':
    # Stablish the AppUserModelID for Windows (this helps the icon appear in the taskbar)
    set_windows_app_id()
    
    # Start the application
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Consistent style across platforms
    
    # Verify that the icon exists
    print(f"Checking for icon at: {ICON_PATH}")
    if not os.path.exists(ICON_PATH):
        print("WARNING: Icon file not found!")
        print(f"Working directory: {os.getcwd()}")
        # Create directory if there is not one
        icon_dir = os.path.dirname(ICON_PATH)
        if not os.path.exists(icon_dir):
            os.makedirs(icon_dir, exist_ok=True)
            print(f"Created directory: {icon_dir}")
    
    # Establish the application at application level (this affects all windows)
    if os.path.exists(ICON_PATH):
        app_icon = QIcon(ICON_PATH)
        app.setWindowIcon(app_icon)
        print("Application icon set successfully")
    else:
        print("Could not set application icon - file not found")
    
    window = MainWindow()
    
    # Explicitly establishes the icon for the main window
    if os.path.exists(ICON_PATH):
        # Try to load the icon in multiple ways for greater compatibility
        icon = QIcon()
        icon.addFile(ICON_PATH, QSize(16, 16))
        icon.addFile(ICON_PATH, QSize(24, 24))
        icon.addFile(ICON_PATH, QSize(32, 32))
        icon.addFile(ICON_PATH, QSize(48, 48))
        icon.addFile(ICON_PATH, QSize(256, 256))
        
        window.setWindowIcon(icon)
        print("Window icon set successfully")
    
    window.show()
    
    sys.exit(app.exec())