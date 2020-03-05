from _csv import writer
from os.path import basename

import docx2txt
from PyPDF2 import PdfFileReader
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QApplication, QPushButton, QLineEdit, QHBoxLayout, QVBoxLayout, QDialog, QGridLayout, QStyleFactory, QListWidget, QFileDialog, QTableWidget, QTableWidgetItem, QHeaderView


class FileScanner(QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        main_layout = QGridLayout()
        main_layout.addLayout(self.create_keyword_area(), 0, 0)
        main_layout.setColumnMinimumWidth(0, 350)
        self.file_headers = ['files']
        main_layout.addLayout(self.create_files_area(), 0, 1)
        main_layout.setColumnMinimumWidth(1, 650)
        main_layout.setColumnStretch(1, 1)
        main_layout.setRowMinimumHeight(0, 500)
        self.setLayout(main_layout)
        self.setWindowTitle('File Scanner')
        QApplication.setStyle(QStyleFactory.create('windowsvista'))
        self.file_names = []
        # self.showMaximized()
        self.show()

    def create_files_area(self):
        vertical = QVBoxLayout()
        top_buttons = QHBoxLayout()
        self.add_files = QPushButton('Add Files')
        self.add_files.clicked.connect(self.add_files_clicked)
        top_buttons.addWidget(self.add_files)
        self.export_results = QPushButton('Export Results')
        self.export_results.setDisabled(True)
        self.export_results.clicked.connect(self.export_results_clicked)
        top_buttons.addWidget(self.export_results)
        vertical.addLayout(top_buttons)
        self.files = QTableWidget()
        self.files.setEditTriggers(QTableWidget.NoEditTriggers)
        self.files.setAlternatingRowColors(True)
        self.update_file_headers()
        vertical.addWidget(self.files)
        buttons = QHBoxLayout()
        self.save_keywords = QPushButton('Save Keywords')
        self.save_keywords.clicked.connect(self.save_keywords_clicked)
        self.save_keywords.setDisabled(True)
        buttons.addWidget(self.save_keywords)
        load_keywords = QPushButton('Load Keywords')
        load_keywords.clicked.connect(self.load_keywords_clicked)
        buttons.addWidget(load_keywords)
        self.scan_files = QPushButton('Scan Files')
        self.save_keywords.setDisabled(True)
        self.scan_files.clicked.connect(self.scan_files_clicked)
        buttons.addWidget(self.scan_files)
        vertical.addLayout(buttons)
        return vertical

    def export_results_clicked(self):
        file = QFileDialog.getSaveFileName(self, 'Save Keywords', '.', 'Excel (*.csv)')[0]
        if file:
            with open(file, 'wt', newline='') as file:
                w = writer(file)
                w.writerow(self.file_headers)
                for row in range(self.files.rowCount()):
                    w.writerow([self.files.item(row, 0).text()] + [self.files.item(row, i + 1).text() for i in range(self.keyword_list.count())])

    def save_keywords_clicked(self):
        file = QFileDialog.getSaveFileName(self, 'Save Keywords', '.', 'Text (*.txt)')[0]
        if file:
            with open(file, 'wt') as files:
                files.writelines(self.keyword_list.item(i).text() for i in range(self.keyword_list.count()))

    def load_keywords_clicked(self):
        file = QFileDialog.getOpenFileName(self, 'Load Keywords', '.', 'Text (*.txt)')[0]
        if file:
            with open(file, 'rt') as files:
                self.keyword_list.addItems(_ for _ in files.read().split('\n') if _)
            self.update_file_headers()

    def scan_files_clicked(self):
        for row in range(self.files.rowCount()):
            file = self.file_names[row]
            if file[-4:] == '.pdf':
                with open(file, 'rb') as process_file:
                    pdf = PdfFileReader(process_file)
                    process_file_text = '\n'.join(pdf.getPage(i).extractText().upper() for i in range(pdf.getNumPages()))
            elif file[-4:] == 'docx':
                process_file_text = docx2txt.process(file).upper()
            else:
                with open(file, 'rt') as process_file:
                    process_file_text = process_file.read()
            for i in range(self.keyword_list.count()):
                has = self.keyword_list.item(i).text().upper() in process_file_text
                widget = QTableWidgetItem(str(has).upper())
                widget.setBackground(QColor('green' if has else 'red'))
                self.files.setItem(row, i + 1, widget)
        self.export_results.setDisabled(False)

    def add_files_clicked(self):
        files = [file for file in QFileDialog.getOpenFileNames(self, 'Files to Scan', '.', 'Word Documents (*.docx);;PDF Documents (*.pdf);;Text (*.txt);;All Files (*)')[0] if file[-3:] in {'pdf', 'ocx', 'txt', '.py', 'csv', 'tsv'}]
        if files:
            count = self.files.rowCount()
            self.files.setRowCount(count + len(files))
            for i, file in enumerate(files):
                self.file_names.append(file)
                self.files.setItem(i + count, 0, QTableWidgetItem(basename(file)))
            if self.keyword_list.count():
                self.scan_files.setDisabled(False)

    def create_keyword_area(self):
        horizontal = QHBoxLayout()
        self.new_keyword_text = QLineEdit()
        self.new_keyword_text.setPlaceholderText('Key word / phrase')
        self.new_keyword_text.textEdited.connect(lambda: self.add_keyword.setDisabled(not self.new_keyword_text.text()))
        horizontal.addWidget(self.new_keyword_text)
        self.add_keyword = QPushButton('Add')
        self.add_keyword.setDisabled(True)
        self.add_keyword.clicked.connect(self.add_keyword_clicked)
        horizontal.addWidget(self.add_keyword)
        vertical = QVBoxLayout()
        vertical.addLayout(horizontal)
        self.keyword_list = QListWidget()
        self.keyword_list.itemSelectionChanged.connect(lambda: self.remove_keyword.setDisabled(not self.keyword_list.selectedIndexes()))
        vertical.addWidget(self.keyword_list)
        self.remove_keyword = QPushButton('Remove')
        self.remove_keyword.setDisabled(True)
        self.remove_keyword.clicked.connect(self.remove_keyword_clicked)
        vertical.addWidget(self.remove_keyword)
        return vertical

    def add_keyword_clicked(self):
        self.keyword_list.addItem(self.new_keyword_text.text())
        self.add_keyword.setDisabled(True)
        self.new_keyword_text.setText('')
        self.update_file_headers()
        self.save_keywords.setDisabled(False)

    def update_file_headers(self):
        self.file_headers = ['File'] + [self.keyword_list.item(i).text() for i in range(self.keyword_list.count())]
        old_columns = self.files.columnCount()
        len_headers = len(self.file_headers)
        self.files.setColumnCount(len_headers)
        self.files.setHorizontalHeaderLabels(self.file_headers)
        if old_columns < len_headers:
            hori = self.files.horizontalHeader()
            for i in range(old_columns, len_headers):
                hori.setSectionResizeMode(i, QHeaderView.ResizeToContents)

    def remove_keyword_clicked(self):
        for item in self.keyword_list.selectedIndexes()[::-1]:
            self.keyword_list.takeItem(item.row())
        self.remove_keyword.setDisabled(True)
        self.update_file_headers()
        if not self.keyword_list.count():
            self.save_keywords.setDisabled(True)


if __name__ == '__main__':
    app = QApplication([])
    main = FileScanner()
    app.exec_()
