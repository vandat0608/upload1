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

        # Xử lý các cell merged
        merged_cells = list(sheet.merged_cells.ranges)
        for merged_range in merged_cells[:]:
            if merged_range.min_row == 1 and merged_range.min_col >= 7:
                sheet.unmerge_cells(str(merged_range))
                first_cell = sheet.cell(row=1, column=merged_range.min_col)
                first_cell.value = first_cell.value

        # Sao chép ngày từ hàng 1, cột 7 trở đi (cứ 4 cột)
        for col in range(7, sheet.max_column + 1, 4):
            cell = sheet.cell(row=1, column=col)
            if isinstance(cell.value, str) and '/' in cell.value:
                date_value = cell.value
                logging.debug(f"Found date {date_value} at column {get_column_letter(col)}")
                for i in range(1, 4):
                    if col + i <= sheet.max_column:
                        sheet.cell(row=1, column=col + i, value=date_value)

        # Thêm hàng 2 cho "Buổi sáng" và "Buổi chiều"
        sheet.insert_rows(2)
        for col in range(7, sheet.max_column + 1, 4):
            if isinstance(sheet.cell(row=1, column=col).value, str) and '/' in sheet.cell(row=1, column=col).value:
                for i in range(2):
                    if col + i <= sheet.max_column:
                        sheet.cell(row=2, column=col + i, value="Buổi sáng")
                for i in range(2, 4):
                    if col + i <= sheet.max_column:
                        sheet.cell(row=2, column=col + i, value="Buổi chiều")

        # Thêm header "C1", "C2", ... cho hàng 3
        counter = 1
        last_date_col = 7
        for col in range(7, sheet.max_column + 1, 4):
            if isinstance(sheet.cell(row=1, column=col).value, str) and '/' in sheet.cell(row=1, column=col).value:
                last_date_col = col + 3

        for col in range(7, last_date_col + 1):
            sheet.cell(row=3, column=col, value=f"C{counter}")
            counter += 1

        workbook.save(file_path)
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
        headers = ["Ngày", "Họ và tên HSSV", "Khoa", "Lớp", "Giáo viên giảng dạy", "Nề nếp", "Buổi"]
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
            
            for col in range(7, source_sheet.max_column + 1, 4):
                date_value = source_sheet.cell(row=1, column=col).value
                if isinstance(date_value, str) and '/' in date_value:
                    for i in range(4):
                        cell_value = source_sheet.cell(row=row, column=col + i).value
                        if cell_value == "K":
                            session = source_sheet.cell(row=2, column=col + i).value or ""
                            summary_sheet[f"A{current_row}"] = date_value
                            summary_sheet[f"B{current_row}"] = full_name
                            summary_sheet[f"C{current_row}"] = faculty_name  # Luôn dùng giá trị từ QComboBox
                            summary_sheet[f"D{current_row}"] = class_name
                            summary_sheet[f"E{current_row}"] = ""
                            summary_sheet[f"F{current_row}"] = "Nghỉ học"
                            summary_sheet[f"G{current_row}"] = session
                            current_row += 1
                            logging.debug(f"Added record for {full_name} on {date_value} with session {session}")
        
        # Định dạng chiều rộng cột
        for col in range(1, 8):
            column_letter = get_column_letter(col)
            summary_sheet.column_dimensions[column_letter].width = 20
        
        workbook.save(output_file_path)
        logging.info(f"Successfully created 'Thống kê nghỉ học' in {output_file_path}")
        return True
    except Exception as e:
        logging.error(f"Error in summarize_k_attendance for {file_path}: {e}")
        return False