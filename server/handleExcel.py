# handleExcel.py
import openpyxl
from openpyxl.utils import get_column_letter
from enum import Enum
import os
import logging
from typing import Optional

# Cấu hình logging
logging.basicConfig(level=logging.INFO)

# Định nghĩa enum EChange (giữ nguyên để tham khảo)
class EChange(Enum):
    K_CNTT_KTD = "Khoa Công nghệ thông tin - Kỹ thuật điện"
    K_DL_KS = "Khoa Du lịch - Khách sạn"
    K_CK = "Khoa Cơ khí"
    K_KT_L = "Khoa Kinh tế- Luật"
    K_CSSD_NDT = "Khoa Chăm sóc sắc đẹp - Nuôi dưỡng trẻ"
    K_YD = "Khoa Y dược"
    K_NN = "Khoa Ngoại Ngữ"
    KHCB = "Khoa học cơ bản"

def copy_dates_and_add_columns(file_path: str) -> bool:
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        logging.debug(f"Processing file: {file_path} for copying dates and adding columns")
        logging.debug(f"Initial max_column: {sheet.max_column}")

        # Log nội dung hàng 1 trước khi xử lý
        header_row = [sheet.cell(row=1, column=col).value for col in range(1, sheet.max_column + 1)]
        logging.debug(f"Header row before processing: {header_row}")

        # Xử lý các cell merged từ cột 7 trở đi
        merged_cells = list(sheet.merged_cells.ranges)
        for merged_range in merged_cells[:]:
            if merged_range.min_row == 1 and merged_range.min_col >= 7:
                # Lưu giá trị của ô đầu tiên trước khi bỏ gộp
                first_cell = sheet.cell(row=1, column=merged_range.min_col)
                cell_value = first_cell.value
                logging.debug(f"Unmerging range {merged_range}, original value: {cell_value}")
                # Bỏ gộp các ô
                sheet.unmerge_cells(str(merged_range))
                # Gán lại giá trị cho ô đầu tiên
                sheet.cell(row=1, column=merged_range.min_col).value = cell_value
                # Đặt các ô còn lại trong vùng gộp thành None
                for row in range(merged_range.min_row, merged_range.max_row + 1):
                    for col in range(merged_range.min_col + 1, merged_range.max_col + 1):
                        sheet.cell(row=row, column=col).value = None

        # Sao chép ngày từ cột 7 trở đi nếu có
        has_date = False
        for col in range(7, sheet.max_column + 1, 5):
            cell = sheet.cell(row=1, column=col)
            logging.debug(f"Checking column {get_column_letter(col)} (col {col}): {cell.value}")
            if isinstance(cell.value, str) and '/' in cell.value:
                has_date = True
                date_value = cell.value
                logging.debug(f"Found date {date_value} at column {get_column_letter(col)}")
                # Sao chép ngày sang các cột tiếp theo
                for i in range(1, 5):
                    if col + i <= sheet.max_column:
                        sheet.cell(row=1, column=col + i, value=date_value)

        # Log nội dung hàng 1 sau khi xử lý
        header_row = [sheet.cell(row=1, column=col).value for col in range(1, sheet.max_column + 1)]
        logging.debug(f"Header row after processing: {header_row}")

        # Thêm hàng 2 cho "Buổi sáng" và "Buổi chiều" từ cột 7 trở đi
        sheet.insert_rows(2)
        for col in range(7, sheet.max_column + 1, 5):
            if isinstance(sheet.cell(row=1, column=col).value, str) and '/' in sheet.cell(row=1, column=col).value:
                for i in range(1, 3):
                    if col + i <= sheet.max_column:
                        sheet.cell(row=2, column=col + i, value="Buổi sáng")
                for i in range(3, 5):
                    if col + i <= sheet.max_column:
                        sheet.cell(row=2, column=col + i, value="Buổi chiều")

        # Thêm header "C1", "C2", ... cho hàng 3 từ cột 7 trở đi
        counter = 1
        last_date_col = 7
        for col in range(7, sheet.max_column + 1, 5):
            if isinstance(sheet.cell(row=1, column=col).value, str) and '/' in sheet.cell(row=1, column=col).value:
                last_date_col = col + 4

        for col in range(7, last_date_col + 1):
            sheet.cell(row=3, column=col, value=f"C{counter}")
            counter += 1

        # Lưu file và kiểm tra
        workbook.save(file_path)
        logging.debug(f"File saved successfully at {file_path}")
        # Xác nhận file đã lưu bằng cách đọc lại
        workbook_check = openpyxl.load_workbook(file_path)
        sheet_check = workbook_check.active
        header_row_check = [sheet_check.cell(row=1, column=col).value for col in range(1, sheet_check.max_column + 1)]
        logging.debug(f"Header row after reload: {header_row_check}")

        logging.info(f"Successfully copied dates and added columns to {file_path}")
        return True
    except Exception as e:
        logging.error(f"Error in copy_dates_and_add_columns for {file_path}: {e}")
        return False

def summarize_k_attendance(file_path: str, faculty_name: str, output_file_path: Optional[str] = None) -> bool:
    if output_file_path is None:
        output_file_path = file_path
    
    try:
        workbook = openpyxl.load_workbook(file_path)
        source_sheet = workbook.active

        logging.debug(f"Summarizing attendance for file: {file_path} with faculty: {faculty_name}")

        # Xóa sheet "Thống kê nghỉ học" nếu đã tồn tại để luôn tạo mới
        summary_sheet_name = "Thống kê nghỉ học"
        if summary_sheet_name in workbook.sheetnames:
            workbook.remove(workbook[summary_sheet_name])
        summary_sheet = workbook.create_sheet(summary_sheet_name)
        
        # Định nghĩa headers cho sheet "Thống kê nghỉ học"
        headers = ["Ngày", "Họ và tên HSSV", "Khoa", "Lớp", "Giáo viên giảng dạy", "Nề nếp", "Buổi", "Phòng"]
        for col, header in enumerate(headers, 1):
            summary_sheet[f"{get_column_letter(col)}1"] = header
        
        # Lấy tên file làm "Lớp"
        class_name = os.path.basename(file_path).replace('.xlsx', '')
        
        current_row = 2  # Bắt đầu từ hàng 2 (sau header)
        for row in range(4, source_sheet.max_row + 1):
            ho_dem = source_sheet[f"C{row}"].value or ""
            ten = source_sheet[f"D{row}"].value or ""
            full_name = f"{ho_dem} {ten}".strip()
            
            logging.debug(f"Processing row {row} for {full_name}")
            
            for col in range(7, source_sheet.max_column + 1, 5):
                date_value = source_sheet.cell(row=1, column=col).value
                if isinstance(date_value, str) and '/' in date_value:
                    room_value = ""  # Để trống, vì không có cột "Phòng" trong file nguồn
                    for i in range(1, 5):
                        cell_value = source_sheet.cell(row=row, column=col + i).value
                        if cell_value == "K":
                            session = source_sheet.cell(row=2, column=col + i).value or ""
                            summary_sheet[f"A{current_row}"] = date_value
                            summary_sheet[f"B{current_row}"] = full_name
                            summary_sheet[f"C{current_row}"] = faculty_name
                            summary_sheet[f"D{current_row}"] = class_name
                            summary_sheet[f"E{current_row}"] = ""
                            summary_sheet[f"F{current_row}"] = "Nghỉ học"
                            summary_sheet[f"G{current_row}"] = session
                            summary_sheet[f"H{current_row}"] = room_value
                            current_row += 1
                            logging.debug(f"Added record for {full_name} on {date_value} with session {session}")
        
        # Định dạng chiều rộng cột
        for col in range(1, 9):
            column_letter = get_column_letter(col)
            summary_sheet.column_dimensions[column_letter].width = 20
        
        workbook.save(output_file_path)
        logging.info(f"Successfully created 'Thống kê nghỉ học' in {output_file_path}")
        return True
    except Exception as e:
        logging.error(f"Error in summarize_k_attendance for {file_path}: {e}")
        return False