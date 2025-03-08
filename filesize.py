import sys
import os

from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog, QTableWidget, \
    QTableWidgetItem, QLabel, QCheckBox, QProgressBar, QMessageBox, QHeaderView

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from PyQt6.QtCore import QThread, pyqtSignal, Qt
import os

class FileSizeWorker(QThread):
    progress_updated = pyqtSignal(int)   # Tín hiệu cập nhật tiến trình
    result_ready = pyqtSignal(list, object)  # Dùng object để tránh ép kiểu


    def __init__(self, folder_path, include_subfolders):
        super().__init__()
        self.folder_path = folder_path
        self.include_subfolders = include_subfolders

    def get_folder_size(self, folder_path):
        """Tính tổng dung lượng của tất cả file bên trong (bao gồm thư mục con)."""
        total_size = 0
        try:
            for root, _, files in os.walk(folder_path):  # Duyệt qua tất cả thư mục con
                for file in files:
                    file_path = os.path.join(root, file)
                    try:
                        total_size += os.path.getsize(file_path)
                    except Exception as e:
                        print(f"Lỗi khi đọc file {file_path}: {e}")
        except Exception as err:
            print("get_folder_size - " , err)
        return total_size

    def run(self):
        file_data = []
        total_size = 0
        file_list = []

        # Lấy danh sách file cần tính toán
        if self.include_subfolders:
            for root, _, files in os.walk(self.folder_path):
                for file in files:
                    file_list.append(os.path.join(root, file))
        else:
            # file_list = [os.path.join(self.folder_path, f) for f in os.listdir(self.folder_path) if os.path.isfile(os.path.join(self.folder_path, f))]
            file_list = [
                os.path.join(self.folder_path, f)
                for f in os.listdir(self.folder_path)
                if os.path.isfile(os.path.join(self.folder_path, f)) or os.path.isdir(os.path.join(self.folder_path, f))
            ]

        total_files = len(file_list)

        for index, item in enumerate(file_list):
            item_path = os.path.join(self.folder_path, item)

            try:
                if os.path.isfile(item_path):  # Nếu là file
                    size_bytes = os.path.getsize(item_path)
                    size_kb = size_bytes / 1024
                    file_data.append([os.path.basename(item), size_kb, "KB", item_path])
                    total_size += size_kb

                elif os.path.isdir(item_path):  # Nếu là thư mục
                    folder_size_bytes = self.get_folder_size(item_path)  # Tính tổng kích thước thư mục
                    folder_size_kb = folder_size_bytes / 1024
                    file_data.append([f"📂 {os.path.basename(item)}" , folder_size_kb, "KB", item_path])
                    total_size += folder_size_kb

            except Exception as e:
                print(f"Lỗi khi đọc {item_path}: {e}")

            # Cập nhật tiến trình
            progress = int((index + 1) / total_files * 100)
            self.progress_updated.emit(progress)

        # Gửi kết quả về giao diện
        # self.result_ready.emit(file_data, round(total_size, 2))  # Làm tròn tổng dung lượng
        print(f"DEBUG - total_size trước khi emit: {total_size} KB, kiểu dữ liệu: {type(total_size)}")
        self.result_ready.emit(file_data, total_size)


class FileSizeChecker(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("FolderSizeMeasure - TrinhThuan - E20250308")
        self.setGeometry(100, 100, 600, 500)
        self.setMinimumSize(200, 150)  # Cho phép kéo nhỏ hơn kích thước ban đầu

        self.setStyleSheet("background-color: #f0f0f0;")
        # Kích hoạt kéo thả
        self.setAcceptDrops(True)

        layout = QVBoxLayout()

        self.label = QLabel("Chọn thư mục để kiểm tra dung lượng các file")
        self.label.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(self.label)

        button_layout = QHBoxLayout()

        self.button = QPushButton("Chọn thư mục")
        self.button.setStyleSheet(
            "background-color: #4CAF50; color: white; font-size: 14px; padding: 8px; border-radius: 5px;")
        self.button.clicked.connect(self.select_folder)
        button_layout.addWidget(self.button)

        self.export_button = QPushButton("Xuất Excel")
        self.export_button.setStyleSheet(
            "background-color: #2196F3; color: white; font-size: 14px; padding: 8px; border-radius: 5px;")
        self.export_button.clicked.connect(self.export_to_excel)
        button_layout.addWidget(self.export_button)

        layout.addLayout(button_layout)

        self.detail_checkbox = QCheckBox("Xem chi tiết tất cả file và thư mục con")
        layout.addWidget(self.detail_checkbox)



        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Tên file", "Dung lượng", "Đơn vị"])
        self.table.setStyleSheet("background-color: white; border: 1px solid #ccc;")

        # Cột đầu tiên mở rộng theo bảng
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)

        # Các cột khác chỉ điều chỉnh theo nội dung
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)

        # self.table.horizontalHeader().setStretchLastSection(True)  # Giãn cột cuối cùng để không bị thừa khoảng trống
        # self.table.setColumnWidth(0, 400)  # Tăng chiều rộng cột tên file
        # self.table.setColumnWidth(1, 100)
        # self.table.setColumnWidth(2, 100)
        layout.addWidget(self.table)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(10)  # Đặt chiều cao thanh tiến trình
        # self.progress_bar.setFixedWidth(600)  # Đặt chiều rộng thanh tiến trình
        self.progress_bar.setStyleSheet(
            "QProgressBar { border: 0px solid grey; border-radius: 5px; text-align: center; height: 10px; width: 500px; } "
            "QProgressBar::chunk { background-color: #00FFCC; width: 10px; }")
        layout.addWidget(self.progress_bar)

        self.total_label = QLabel("Tổng dung lượng: 0 KB")
        self.total_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #d32f2f;")
        layout.addWidget(self.total_label)

        self.setLayout(layout)

    def dragEnterEvent(self, event):
        """ Khi người dùng kéo một thư mục vào giao diện """
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        """ Khi người dùng thả thư mục vào giao diện """
        urls = event.mimeData().urls()
        if urls:
            folder_path = urls[0].toLocalFile()
            if os.path.isdir(folder_path):  # Kiểm tra xem có phải thư mục không
                self.label.setText(f"Thư mục đã chọn: {folder_path}")
                self.get_folder_size(folder_path)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục")
        if folder:
            self.label.setText(f"Thư mục đã chọn: {folder}")
            self.get_folder_size(folder)

    def get_folder_size(self, folder):

        self.table.setRowCount(0)  # Xóa dữ liệu cũ
        self.progress_bar.setValue(0)  # Reset tiến trình
        self.total_label.setText("Đang tính toán...")  # Cập nhật trạng thái

        self.worker = FileSizeWorker(folder, self.detail_checkbox.isChecked())
        self.worker.progress_updated.connect(self.progress_bar.setValue)
        self.worker.result_ready.connect(self.display_results)
        self.worker.start()  # Chạy luồng tính toán

    def display_results(self, file_data, total_size):
        self.table.setRowCount(0)  # Xóa dữ liệu cũ trong bảng
        try:
            # Sắp xếp theo dung lượng giảm dần (chuyển về float để so sánh)
            file_data.sort(key=lambda x: float(x[1]), reverse=True)
            # Chuyển đổi dung lượng sang đơn vị phù hợp
            def convert_size(size_kb):
                if size_kb >= 1024 * 1024:
                    return f"{size_kb / (1024 * 1024):.2f}", "GB"
                elif size_kb >= 1024:
                    return f"{size_kb / 1024:.2f}", "MB"
                else:
                    return f"{size_kb:.2f}","KB"

            show_full_path = self.detail_checkbox.isChecked()  # Kiểm tra trạng thái checkbox

            # Chuyển đổi đơn vị dung lượng cho từng file
            formatted_data = []
            for file_info in file_data:
                file_name = file_info[0] if not show_full_path else file_info[3]
                size, unit = convert_size(file_info[1])
                formatted_data.append([file_name, size, unit])



            # Cập nhật bảng hiển thị
            self.table.setRowCount(len(formatted_data))
            for row, (file_name, size, unit) in enumerate(formatted_data):
                item_name = QTableWidgetItem(file_name)
                item_size = QTableWidgetItem(size)
                item_unit = QTableWidgetItem(unit)

                item_size.setTextAlignment(0x0004 | 0x0080)  # Căn giữa
                item_unit.setTextAlignment(0x0004 | 0x0080)  # Căn giữa

                self.table.setItem(row, 0, item_name)
                self.table.setItem(row, 1, item_size)
                self.table.setItem(row, 2, item_unit)
            # self.table.setRowCount(len(file_data))
            # for row, data in enumerate(file_data):
            #     for col, value in enumerate(data):
            #         item = QTableWidgetItem(value)
            #         item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            #         self.table.setItem(row, col, item)
            print(f"DEBUG - Kiểu dữ liệu total_size: {type(total_size)}, Giá trị: {total_size}")

            if not isinstance(total_size, (int, float)):
                try:
                    total_size = float(total_size)  # Chuyển thành float nếu bị đổi kiểu
                except ValueError:
                    print("Lỗi: Không thể chuyển đổi total_size sang float")
                    total_size = 0
            total_size_formated , unit = convert_size(total_size)

            print(f"DEBUG - total_size sau khi xử lý: {total_size_formated} {unit}")
            self.total_label.setText(f"Tổng dung lượng: {total_size_formated} {unit}")
        except Exception as err:
            print("display_results - exception:", err)

    # def add_file_data(self, file, size_bytes):
    #     if size_bytes >= 1024 * 1024:
    #         size_display = size_bytes / (1024 * 1024)
    #         unit = "MB"
    #     elif size_bytes >= 1024:
    #         size_display = size_bytes / 1024
    #         unit = "KB"
    #     else:
    #         size_display = size_bytes
    #         unit = "Bytes"
    #
    #     self.file_data.append((file, size_display, unit, size_bytes))
    #
    # def list_files(self, folder):
    #     self.folder = folder
    #     self.file_data = []
    #     total_size = 0
    #
    #     files_list = []
    #     if self.detail_checkbox.isChecked():
    #         for dirpath, _, filenames in os.walk(folder):
    #             for file in filenames:
    #                 files_list.append(os.path.join(dirpath, file))
    #     else:
    #         files_list = [os.path.join(folder, file) for file in os.listdir(folder)]
    #
    #     self.progress_bar.setMaximum(len(files_list))
    #
    #     for index, file_path in enumerate(files_list):
    #         if os.path.isfile(file_path):
    #             size_bytes = os.path.getsize(file_path)
    #             total_size += size_bytes
    #         elif os.path.isdir(file_path):
    #             size_bytes = self.get_folder_size(file_path)
    #             total_size += size_bytes
    #         else:
    #             size_bytes = 0
    #
    #         self.add_file_data(file_path, size_bytes)
    #         self.progress_bar.setValue(index + 1)
    #
    #     self.file_data.sort(key=lambda x: x[3], reverse=True)
    #
    #     self.table.setRowCount(len(self.file_data))
    #     for i, (file, size_display, unit, _) in enumerate(self.file_data):
    #         self.table.setItem(i, 0, QTableWidgetItem(file))
    #
    #         item_size = QTableWidgetItem(f"{size_display:.2f}" if isinstance(size_display, float) else size_display)
    #         item_size.setTextAlignment(0x0002 | 0x0080)
    #         self.table.setItem(i, 1, item_size)
    #
    #         item_unit = QTableWidgetItem(unit)
    #         item_unit.setTextAlignment(0x0004 | 0x0080)
    #         self.table.setItem(i, 2, item_unit)
    #
    #     total_size_display, total_unit = (total_size / (1024 * 1024), "MB") if total_size >= 1024 * 1024 else (
    #     total_size / 1024, "KB") if total_size >= 1024 else (total_size, "Bytes")
    #     self.total_label.setText(f"Tổng dung lượng: {total_size_display:.2f} {total_unit}")
    #     self.progress_bar.setValue(len(files_list))

    def export_to_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Lưu file", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                wb = Workbook()
                ws = wb.active
                data = [["Tên file", "Dung lượng", "Đơn vị"]]

                # Lấy dữ liệu từ bảng giao diện
                for row in range(self.table.rowCount()):
                    file_name = self.table.item(row, 0).text()
                    size = self.table.item(row, 1).text()
                    unit = self.table.item(row, 2).text()
                    data.append([file_name, size, unit])

                # Lấy tổng dung lượng hiển thị trên nhãn
                total_size_text = self.total_label.text().replace("Tổng dung lượng: ", "").split(" ")
                total_size, total_unit = total_size_text[0], total_size_text[1]
                data.append(["Tổng cộng", total_size, total_unit])

                # Ghi toàn bộ dữ liệu một lần
                for row in data:
                    ws.append(row)

                # Áp dụng định dạng sau khi ghi dữ liệu
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                     top=Side(style='thin'), bottom=Side(style='thin'))
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

                for col in ws.columns:
                    for cell in col:
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center')
                        if cell.row == 1:
                            cell.font = Font(bold=True)
                            cell.fill = yellow_fill

                wb.save(file_path)
                QMessageBox.information(self, "Xuất Excel", f"Xuất file Excel hoàn tất: {file_path}")
            except PermissionError:
                QMessageBox.warning(self, "Lỗi", "File Excel đang mở, vui lòng đóng file trước khi xuất.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileSizeChecker()
    window.show()
    sys.exit(app.exec())
