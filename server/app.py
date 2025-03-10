from flask import Flask, request, jsonify
from flask_cors import CORS
import os
from handleExcel import copy_dates_and_add_columns, summarize_k_attendance
from uploadGgSheet import extract_spreadsheet_id, upload_to_google_sheets, connect_to_google_sheets
from network_checker import check_network, check_internet_speed
import tempfile
import openpyxl
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv  # Thêm import để đọc biến môi trường

import logging
import shutil

# Load biến môi trường từ file .env
load_dotenv()

# Định nghĩa CREDENTIALS_FILE sử dụng biến môi trường hoặc đường dẫn tĩnh
CREDENTIALS_FILE = os.getenv('GOOGLE_CREDENTIALS_PATH', 'credentials.json')

# Cấu hình logging
logging.basicConfig(level=logging.INFO)

app = Flask(__name__)
CORS(app, resources={r"/check-network": {"origins": "*"}, r"/process": {"origins": "*"}})  # Cho phép mọi nguồn

@app.route('/', methods=['GET'])  # Thêm route cho '/'
def home():
    return jsonify({"message": "Welcome to Excel to Google Sheets Uploader API. Use /check-network or /process for functionality."})

@app.route('/check-network', methods=['POST', 'GET'])  # Thêm GET để kiểm tra
def check_network_status():
    is_connected, network_message = check_network()
    if not is_connected:
        return jsonify({"error": True, "message": network_message})
    
    is_speed_ok, speed_message = check_internet_speed()
    if not is_speed_ok:
        return jsonify({"error": True, "message": speed_message})
    
    message = f"{network_message} - {speed_message}" if "Mạng chậm" in network_message or "Mạng chậm" in speed_message else "Mạng ổn định, tiếp tục xử lý."
    logging.info(message)
    return jsonify({"error": False, "message": message})

@app.route('/process', methods=['POST'])
def process_files():
    files = request.files.getlist('files')
    google_sheet_url = request.form.get('googleSheetUrl')
    sheet_name = request.form.get('sheetName')
    faculty_name = request.form.get('faculty')

    if not files or not google_sheet_url or not sheet_name or not faculty_name:
        return jsonify({"status": "Lỗi: Vui lòng nhập đầy đủ thông tin (file, URL Google Sheet, Tên Sheet, và Khoa)!", "error": True})

    spreadsheet_id = extract_spreadsheet_id(google_sheet_url)
    if not spreadsheet_id:
        return jsonify({"status": "Lỗi: Không thể trích xuất spreadsheet_id từ URL!", "error": True})

    service = connect_to_google_sheets(spreadsheet_id)
    if not service:
        return jsonify({"status": "Lỗi: Không thể kết nối Google Sheets!", "error": True})

    successful_files = 0
    total_rows_added = 0
    temp_dir = tempfile.mkdtemp()
    logging.debug(f"Created temporary directory: {temp_dir}")

    for file in files:
        if not file.filename.lower().endswith('.xlsx'):
            logging.warning(f"Bỏ qua file '{file.filename}' do không phải file Excel (.xlsx)")
            continue

        # Lưu file trực tiếp vào temp_dir
        file_path = os.path.join(temp_dir, file.filename)
        logging.debug(f"Saving file to: {file_path}")
        try:
            file.save(file_path)
            logging.debug(f"File saved to temp directory: {file_path}")
        except Exception as e:
            logging.error(f"Error saving file {file.filename} to {file_path}: {e}")
            continue

        if not os.path.exists(file_path):
            logging.warning(f"Bỏ qua file '{file_path}' do không tồn tại sau khi lưu")
            continue

        # Gọi copy_dates_and_add_columns trước khi validate
        logging.debug(f"Calling copy_dates_and_add_columns for {file_path}")
        if not copy_dates_and_add_columns(file_path):
            logging.error(f"Failed to copy dates and add columns for {file_path}")
            continue

        # Đọc lại file để xác nhận thay đổi
        workbook_check = openpyxl.load_workbook(file_path)
        sheet_check = workbook_check.active
        header_row_check = [sheet_check.cell(row=1, column=col).value for col in range(1, sheet_check.max_column + 1)]
        logging.debug(f"Header row after copy_dates_and_add_columns: {header_row_check}")

        # Kiểm tra dữ liệu sau khi đã thêm cột "Phòng"
        is_valid, error_message = validate_excel_data(file_path)
        if not is_valid:
            logging.error(f"Lỗi với file '{file.filename}': {error_message}")
            continue

        rows_added = process_single_file(file_path, spreadsheet_id, sheet_name, faculty_name, service)
        if rows_added is not False:
            successful_files += 1
            total_rows_added += rows_added
        else:
            logging.error(f"Bỏ qua file '{file.filename}' do lỗi xử lý")

    shutil.rmtree(temp_dir)

    status = f"Hoàn tất toàn bộ quá trình: Tổng số file được thêm thành công: {successful_files}, Tổng số hàng được thêm: {total_rows_added}"
    logging.info(status)
    return jsonify({"status": status, "successfulFiles": successful_files, "totalRowsAdded": total_rows_added, "error": False})

def validate_excel_data(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        
        if sheet.max_row < 4:
            return False, f"Thiếu dữ liệu từ hàng 4 trở đi (tối thiểu cần 4 hàng)."
        for row in range(4, sheet.max_row + 1):
            if not sheet[f'C{row}'].value or not sheet[f'D{row}'].value:
                return False, f"Thiếu họ tên ở hàng {row}, cột C hoặc D."
        
        if sheet.max_column < 7:
            return False, f"Thiếu dữ liệu từ cột 7 trở đi (cột G)."
        
        # Log nội dung hàng 1 để debug
        header_row = [sheet.cell(row=1, column=col).value for col in range(1, sheet.max_column + 1)]
        logging.debug(f"Header row of {file_path}: {header_row}")

        # Kiểm tra dữ liệu ngày từ cột 7 trở đi
        has_date = False
        for col in range(7, sheet.max_column + 1, 5):
            date_value = sheet.cell(row=1, column=col).value
            if isinstance(date_value, str) and '/' in date_value:
                has_date = True
                break
        if not has_date:
            return False, "Không tìm thấy dữ liệu ngày hợp lệ từ cột 7 trở đi (phải chứa '/')."
        
        return True, ""
    except Exception as e:
        return False, f"Lỗi khi kiểm tra dữ liệu - {str(e)}."

def process_single_file(file_path, spreadsheet_id, sheet_name, faculty_name, service):
    try:
        logging.debug(f"Processing file: {file_path}")
        if not copy_dates_and_add_columns(file_path):
            logging.error(f"Failed to copy dates and add columns for {file_path}")
            return False
        
        if not summarize_k_attendance(file_path, faculty_name):
            logging.error(f"Failed to summarize attendance for {file_path}")
            return False

        workbook = openpyxl.load_workbook(file_path)
        summary_sheet = workbook["Thống kê nghỉ học"]  # Sử dụng cú pháp mới
        rows_added = summary_sheet.max_row - 1  # Trừ hàng tiêu đề (bỏ header)
        
        # Gọi upload_to_google_sheets đã sửa trong uploadGgSheet.py
        if not upload_to_google_sheets(file_path, spreadsheet_id, sheet_name, service):
            logging.error(f"Failed to upload {file_path} to Google Sheets")
            return False
        
        logging.debug(f"Successfully processed file {file_path}, added {rows_added} rows")
        return rows_added
    except Exception as e:
        logging.error(f"Error processing file {file_path}: {e}")
        return False

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)  # Sử dụng debug=False cho sản xuất