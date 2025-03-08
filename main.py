import sys
import os
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTableWidget, \
    QTableWidgetItem, QLabel


class FileSizeChecker(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Kiểm tra dung lượng file")
        self.setGeometry(100, 100, 600, 400)

        layout = QVBoxLayout()

        self.label = QLabel("Chọn thư mục để kiểm tra dung lượng các file")
        layout.addWidget(self.label)

        self.button = QPushButton("Chọn thư mục")
        self.button.clicked.connect(self.select_folder)
        layout.addWidget(self.button)

        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Tên file", "Dung lượng (KB)"])
        layout.addWidget(self.table)

        self.setLayout(layout)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục")
        if folder:
            self.label.setText(f"Thư mục đã chọn: {folder}")
            self.list_files(folder)

    def list_files(self, folder):
        files = os.listdir(folder)
        self.table.setRowCount(len(files))

        for i, file in enumerate(files):
            file_path = os.path.join(folder, file)
            if os.path.isfile(file_path):
                size_kb = os.path.getsize(file_path) / 1024  # Chuyển byte sang KB
                self.table.setItem(i, 0, QTableWidgetItem(file))
                self.table.setItem(i, 1, QTableWidgetItem(f"{size_kb:.2f}"))
            else:
                self.table.setItem(i, 0, QTableWidgetItem(file))
                self.table.setItem(i, 1, QTableWidgetItem("Thư mục"))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileSizeChecker()
    window.show()
    sys.exit(app.exec())
