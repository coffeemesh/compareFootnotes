import os
import subprocess
import sys

from docx2python import docx2python
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


class CompareFootnotesApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Compare Footnotes")
        self.resize(800, 600)  # Set the initial window size

        main_layout = QVBoxLayout()

        # Add first row
        main_layout.addLayout(self.create_first_row())

        # Add second row
        main_layout.addLayout(self.create_second_row())

        # Add table for viewing tabular data
        self.create_table(main_layout)

        # Add table for unique extra footnotes
        self.create_unique_footnotes_table(main_layout)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def create_first_row(self):
        """Create the first row with label, text input, and button for selecting a document."""
        first_row_layout = QHBoxLayout()

        self.base_doc_label = QLabel("Base Document")
        first_row_layout.addWidget(self.base_doc_label)

        self.base_doc_text = QLineEdit()
        self.base_doc_text.setReadOnly(True)
        first_row_layout.addWidget(self.base_doc_text)

        self.base_doc_button = QPushButton("Select Document")
        self.base_doc_button.clicked.connect(self.select_base_document)
        first_row_layout.addWidget(self.base_doc_button)

        return first_row_layout

    def create_second_row(self):
        """Create the second row with label, text input, and button for selecting a directory."""
        second_row_layout = QHBoxLayout()

        self.dir_label = QLabel("DocX Files Directory")
        second_row_layout.addWidget(self.dir_label)

        self.dir_text = QLineEdit()
        self.dir_text.setReadOnly(True)
        second_row_layout.addWidget(self.dir_text)

        self.dir_button = QPushButton("Select Directory")
        self.dir_button.setEnabled(False)
        self.dir_button.setToolTip("Please select a base document first.")
        self.dir_button.clicked.connect(self.select_directory)
        second_row_layout.addWidget(self.dir_button)

        return second_row_layout

    def create_table(self, layout):
        """Create a table for viewing tabular data and add it to the layout."""
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(
            [
                "Filename",
                "Amount of Footnotes",
                "Difference to base document",
                "Extra footnotes",
            ]
        )
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive
        )
        self.table.setSortingEnabled(True)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.table.cellDoubleClicked.connect(self.open_file)
        layout.addWidget(self.table)

    def create_unique_footnotes_table(self, layout):
        """Create a table for unique extra footnotes and add it to the layout."""
        self.unique_footnotes_table = QTableWidget()
        self.unique_footnotes_table.setColumnCount(2)
        self.unique_footnotes_table.setHorizontalHeaderLabels(
            ["Select", "Unique Extra Footnotes"]
        )
        self.unique_footnotes_table.horizontalHeader().setDefaultAlignment(
            Qt.AlignmentFlag.AlignLeft
        )
        self.unique_footnotes_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive
        )
        self.unique_footnotes_table.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOn
        )
        self.unique_footnotes_table.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOn
        )
        layout.addWidget(self.unique_footnotes_table)

    def select_base_document(self):
        """Open a file dialog to select a document and display its path in the text input."""
        try:
            file_name, _ = QFileDialog.getOpenFileName(
                self, "Select Document", "", "Word Documents (*.docx)"
            )
            if file_name:
                self.base_doc_text.setText(file_name)
                self.dir_button.setEnabled(True)
                self.dir_button.setToolTip("")
        except Exception as e:
            self.show_error_message(f"Error selecting base document: {e}")

    def select_directory(self):
        """Open a file dialog to select a directory and display its path in the text input."""
        try:
            directory = QFileDialog.getExistingDirectory(self, "Select Directory")
            if directory:
                self.dir_text.setText(directory)
                self.populate_main_table(directory)
                self.populate_unique_footnotes_table()
                self.adjust_table_and_window_size()
        except Exception as e:
            self.show_error_message(f"Error selecting directory: {e}")

    def populate_main_table(self, directory):
        """Populate the main table with .docx files from the selected directory."""
        try:
            self.table.setRowCount(0)  # Clear existing rows
            self.all_extra_footnotes = (
                set()
            )  # Initialize set for unique extra footnotes
            base_doc_filename = os.path.basename(self.base_doc_text.text())
            base_footnotes = self.get_footnotes(self.base_doc_text.text())
            for file in os.listdir(directory):
                if file.endswith(".docx") and file != base_doc_filename:
                    row_position = self.table.rowCount()
                    self.table.insertRow(row_position)
                    filename_without_extension = os.path.splitext(file)[0]
                    self.table.setItem(
                        row_position,
                        0,
                        QTableWidgetItem(str(filename_without_extension)),
                    )

                    footnotes = self.get_footnotes(os.path.join(directory, file))
                    footnotes_count = len(footnotes)
                    footnotes_count_item = QTableWidgetItem(str(footnotes_count))
                    footnotes_count_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    footnotes_count_item.setFlags(
                        Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
                    )
                    self.table.setItem(row_position, 1, footnotes_count_item)

                    extra_footnotes = [
                        tuple(footnote) if isinstance(footnote, list) else footnote
                        for footnote in footnotes
                        if footnote not in base_footnotes
                    ]
                    self.all_extra_footnotes.update(
                        extra_footnotes
                    )  # Add each footnote string directly to the set
                    extra_footnotes_str = ", ".join(
                        f'"{footnote}"' for footnote in extra_footnotes
                    )

                    # Create and configure the item for the count of extra footnotes
                    extra_footnotes_count_item = QTableWidgetItem(
                        str(len(extra_footnotes))
                    )
                    extra_footnotes_count_item.setTextAlignment(
                        Qt.AlignmentFlag.AlignCenter
                    )
                    extra_footnotes_count_item.setFlags(
                        Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
                    )
                    self.table.setItem(row_position, 2, extra_footnotes_count_item)

                    # Create and configure the item for the extra footnotes
                    extra_footnotes_item = QTableWidgetItem(extra_footnotes_str)
                    extra_footnotes_item.setFlags(
                        Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
                    )
                    self.table.setItem(row_position, 3, extra_footnotes_item)
        except Exception as e:
            self.show_error_message(f"Error populating Main table: {e}")

    def populate_unique_footnotes_table(self):
        """Populate the unique extra footnotes table with unique extra footnotes."""
        try:
            # Clear existing rows
            self.unique_footnotes_table.setRowCount(0)

            # Populate the table with unique extra footnotes
            for footnote in self.all_extra_footnotes:
                row_position = self.unique_footnotes_table.rowCount()
                self.unique_footnotes_table.insertRow(row_position)

                # Create and configure the checkbox item
                checkbox_item = QTableWidgetItem()
                checkbox_item.setFlags(
                    Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled
                )
                checkbox_item.setCheckState(Qt.CheckState.Unchecked)
                self.unique_footnotes_table.setItem(row_position, 0, checkbox_item)

                # Ensure footnote is a string
                footnote_str = str(footnote)

                # Create and configure the footnote item
                footnote_item = QTableWidgetItem(footnote_str)
                footnote_item.setFlags(
                    Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
                )
                self.unique_footnotes_table.setItem(row_position, 1, footnote_item)
        except Exception as e:
            print(f"Error populating unique footnotes table: {e}")
            self.show_error_message(f"Error populating unique footnotes table: {e}")

    def adjust_table_and_window_size(self):
        """Adjust the table and window size after populating the table."""
        try:
            self.table.resizeColumnsToContents()
            self.unique_footnotes_table.resizeColumnsToContents()
            self.resize(self.sizeHint())
        except Exception as e:
            self.show_error_message(f"Error adjusting table and window size: {e}")

    def get_footnotes(self, file_path):
        """Extract footnotes from a .docx file using docx2python."""
        try:
            doc_result = docx2python(file_path)
            footnotes = doc_result.footnotes
            extracted_footnotes = []
            for footnote in footnotes:
                for paragraph in footnote:
                    for line in paragraph:
                        extracted_footnotes.append(line)
            return extracted_footnotes
        except Exception as e:
            self.show_error_message(f"Error extracting footnotes: {e}")
            return []

    def open_file(self, row, column):
        """Open the corresponding file when a row is double-clicked."""
        try:
            directory = self.dir_text.text()
            filename = f"{self.table.item(row, 0).text()}.docx"
            file_path = os.path.join(directory, filename)

            if os.path.exists(file_path):
                self._open_file_by_platform(file_path)
            else:
                self.show_error_message(f"File not found: {file_path}")
        except Exception as e:
            self.show_error_message(f"Error opening file: {e}")

    def _open_file_by_platform(self, file_path):
        """Open a file using the appropriate command for the current platform."""
        try:
            if sys.platform.startswith("win32"):
                os.startfile(file_path)
            elif sys.platform.startswith("darwin"):
                subprocess.call(["open", file_path])
            elif sys.platform.startswith("linux") or sys.platform.startswith("linux2"):
                subprocess.call(["xdg-open", file_path])
            else:
                raise OSError(f"Unsupported platform: {sys.platform}")
        except Exception as e:
            self.show_error_message(
                f"Error executing platform-specific file open command: {e}"
            )

    def show_error_message(self, message):
        """Display an error message in a message box."""
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Icon.Critical)
        msg_box.setText(message)
        msg_box.setWindowTitle("Error")
        msg_box.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CompareFootnotesApp()
    window.show()
    sys.exit(app.exec())
