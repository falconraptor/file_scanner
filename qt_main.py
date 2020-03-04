from os.path import basename

from PyQt5.QtWidgets import QApplication, QPushButton, QLineEdit, QHBoxLayout, QVBoxLayout, QDialog, QGridLayout, QStyleFactory, QListWidget, QFileDialog, QTableWidget, QTableWidgetItem, QHeaderView


class FileScanner(QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        main_layout = QGridLayout()
        main_layout.addLayout(self.create_keyword_area(), 0, 0)
        self.file_headers = ['files']
        main_layout.addLayout(self.create_files_area(), 0, 1)
        self.setLayout(main_layout)
        self.setWindowTitle('File Scanner')
        QApplication.setStyle(QStyleFactory.create('windowsvista'))

    def create_files_area(self):
        vertical = QVBoxLayout()
        self.add_files = QPushButton('Add Files')
        self.add_files.clicked.connect(self.add_files_clicked)
        vertical.addWidget(self.add_files)
        self.files = QTableWidget()
        self.files.setEditTriggers(QTableWidget.NoEditTriggers)
        self.update_file_headers()
        vertical.addWidget(self.files)
        return vertical

    def add_files_clicked(self):
        files = [file for file in QFileDialog.getOpenFileNames(self, 'Files to Scan', '', 'Word Documents (*.docx);;PDF Documents (*.pdf);;Text (*.txt);;All Files (*)')[0] if file[-3:] in {'pdf', 'ocx', 'txt', '.py', 'csv', 'tsv'}]
        if files:
            count = self.files.rowCount()
            self.files.setRowCount(count + len(files))
            for i, file in enumerate(files):
                self.files.setItem(i + count, 0, QTableWidgetItem(basename(file)))

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


if __name__ == '__main__':
    app = QApplication([])
    main = FileScanner()
    main.show()
    app.exec_()
