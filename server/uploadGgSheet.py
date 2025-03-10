# server/uploadGgSheet.py
import os
import logging
import json
import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from typing import Optional, List, Tuple
from dotenv import load_dotenv

# Load biến môi trường từ file .env
load_dotenv()

# Cấu hình logging
logging.basicConfig(level=logging.INFO)

# Định nghĩa CREDENTIALS_FILE sử dụng biến môi trường hoặc đường dẫn tĩnh
CREDENTIALS_FILE = os.getenv('GOOGLE_CREDENTIALS_PATH', 'credentials.json')

# Log để kiểm tra biến
logging.debug(f"Using credentials file: {CREDENTIALS_FILE}")

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def extract_spreadsheet_id(url: str) -> Optional[str]:
    """
    Trích xuất ID của Google Spreadsheet từ URL.
    """
    if not url or not isinstance(url, str):
        return None
    try:
        if 'spreadsheets/d/' in url:
            start = url.find('spreadsheets/d/') + len('spreadsheets/d/')
            end = url.find('/edit', start)
            if end == -1:
                end = url.find('/', start)  # Trường hợp không có /edit
            if start != -1 and end != -1:
                spreadsheet_id = url[start:end].replace(' ', '')
                return spreadsheet_id
        return None
    except Exception as e:
        logging.error(f"Error extracting spreadsheet ID: {e}")
        return None

def connect_to_google_sheets(spreadsheet_id: str):
    try:
        # Lấy credentials từ biến môi trường trước
        creds_json = os.getenv('GOOGLE_CREDENTIALS')
        if creds_json:
            logging.debug("Loading credentials from environment variable")
            credentials_data = json.loads(creds_json)
            creds = service_account.Credentials.from_service_account_info(
                credentials_data, scopes=SCOPES
            )
        elif os.path.exists(CREDENTIALS_FILE):
            logging.debug(f"Loading credentials from file: {CREDENTIALS_FILE}")
            with open(CREDENTIALS_FILE, 'r') as f:
                credentials_data = json.load(f)
            creds = service_account.Credentials.from_service_account_file(
                CREDENTIALS_FILE, scopes=SCOPES
            )
        else:
            raise FileNotFoundError(f"No credentials found in env or file: {CREDENTIALS_FILE}")

        service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
        return service
    except Exception as e:
        logging.error(f"Error connecting to Google Sheets: {e}")
        return None

def read_excel_data(file_path, sheet_name='Thống kê nghỉ học') -> List[Tuple]:
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):  # Bỏ tiêu đề (hàng 1)
            data.append(row)
        return data
    except Exception as e:
        logging.error(f"Error reading Excel data: {e}")
        return []

def get_first_empty_row(service, spreadsheet_id: str, sheet_name: str) -> int:
    """
    Tìm hàng ngay sau hàng cuối cùng có dữ liệu trong Google Sheets, kiểm tra từ cột B đến H.
    Đảm bảo dữ liệu mới được thêm nối tiếp sau dữ liệu cũ, không ghi đè.
    """
    try:
        range_name = f'{sheet_name}!B:H'  # Kiểm tra toàn bộ từ cột B đến H (thêm cột Phòng)
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            logging.debug("No existing data in Google Sheets, starting from row 2 (after header)")
            return 2  # Nếu không có dữ liệu, bắt đầu từ hàng 2 (giả sử hàng 1 là header)
        
        # Tìm hàng cuối cùng có ít nhất một giá trị không rỗng trong cột B đến H
        last_row_with_data = 1  # Mặc định là hàng 1 (header)
        for row_idx, row in enumerate(values, 1):
            # Kiểm tra nếu có bất kỳ giá trị không rỗng nào trong cột B đến H
            has_data = any(
                cell is not None and
                not (isinstance(cell, str) and cell.strip() == '') and
                not (isinstance(cell, (int, float)) and cell == 0) and
                not (isinstance(cell, bool) and not cell)
                for cell in row[:7]  # Kiểm tra 7 cột (B đến H)
            )
            if has_data:
                last_row_with_data = row_idx  # Cập nhật hàng cuối cùng có dữ liệu
        
        # Trả về hàng tiếp theo sau hàng cuối cùng có dữ liệu
        next_row = last_row_with_data + 1
        logging.debug(f"Found last row with data at {last_row_with_data}, appending at row {next_row}")
        return next_row
    except HttpError as e:
        logging.error(f"Error checking data on Google Sheets: {e}")
        return 2  # Mặc định bắt đầu từ hàng 2 nếu có lỗi (giả sử hàng 1 là header)

def push_data_to_google_sheets(service, spreadsheet_id: str, sheet_name: str, excel_data: List[Tuple]):
    """
    Đẩy dữ liệu từ Excel lên Google Sheets, bắt đầu từ hàng sau cùng có dữ liệu,
    gộp cột G và H vào cột H, đảm bảo dữ liệu nối tiếp nhau.
    """
    try:
        start_row = get_first_empty_row(service, spreadsheet_id, sheet_name)
        
        # Chuẩn bị dữ liệu, gộp cột G (Nề nếp) và H (Buổi) vào cột H với định dạng rõ ràng
        formatted_data = []
        for row in excel_data:
            # Đảm bảo xử lý trường hợp dữ liệu thiếu (None hoặc rỗng)
            col_g = str(row[6]) if row[6] is not None else 'Nghỉ học'  # Mặc định "Nghỉ học" nếu rỗng
            col_h = str(row[7]) if row[7] is not None else ''  # Buổi, mặc định rỗng nếu thiếu
            combined_h = f"{col_g.strip()} {col_h.strip()}".strip() if col_h.strip() else col_g  # Gộp nếu có Buổi
            
            formatted_row = [
                str(row[0]) if row[0] is not None else '',  # Ngày
                str(row[1]) if row[1] is not None else '',  # Phòng
                str(row[2]) if row[2] is not None else '',  # Họ và tên HSSV
                str(row[3]) if row[3] is not None else '',  # Khoa
                str(row[4]) if row[4] is not None else '',  # Lớp
                str(row[5]) if row[5] is not None else '',  # Giáo viên giảng dạy
                combined_h  # Gộp Nề nếp và Buổi
            ]
            formatted_data.append(formatted_row)

        if not formatted_data:
            logging.warning("No data to push to Google Sheets")
            return

        body = {
            'values': formatted_data
        }
        
        range_name = f'{sheet_name}!B{start_row}'
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()
        
        logging.info(f"Đã đẩy dữ liệu lên Google Sheets, bắt đầu từ hàng {start_row}!")
    except HttpError as e:
        logging.error(f"Error pushing data to Google Sheets: {e}")

def upload_to_google_sheets(file_path, spreadsheet_id, sheet_name, service):
    try:
        workbook = openpyxl.load_workbook(file_path)
        if "Thống kê nghỉ học" not in workbook.sheetnames:
            logging.error(f"Sheet 'Thống kê nghỉ học' not found in {file_path}")
            return False
        
        sheet = workbook["Thống kê nghỉ học"]
        data_to_upload = []

        # Chuẩn bị dữ liệu từ sheet "Thống kê nghỉ học"
        for row in sheet.iter_rows(min_row=2, values_only=True):  # Lấy dữ liệu từ hàng 2 trở đi
            date_value = row[0] or ""  # Cột A: Ngày
            full_name = row[1] or ""   # Cột B: Họ và tên HSSV
            faculty = row[2] or ""     # Cột C: Khoa
            class_name = row[3] or ""  # Cột D: Lớp
            teacher = row[4] or ""     # Cột E: Giáo viên giảng dạy
            attendance = row[5] or ""  # Cột F: Nề nếp
            session = row[6] or ""     # Cột G: Buổi
            room = row[7] or ""        # Cột H: Phòng

            # Kết hợp Nề nếp và Buổi cho cột H, bỏ dấu "+", chỉ dùng khoảng trắng
            combined_attendance = f"{attendance} {session}" if attendance and session else (attendance or session or "")

            # Tạo mảng dữ liệu theo thứ tự ánh xạ
            mapped_row = [
                "",          # Cột A: Để trống
                date_value,  # Cột B: Ngày
                room,        # Cột C: Phòng
                full_name,   # Cột D: Họ và tên HSSV
                faculty,     # Cột E: Khoa
                class_name,  # Cột F: Lớp
                teacher,     # Cột G: Giáo viên giảng dạy
                combined_attendance  # Cột H: Nề nếp Buổi
            ]
            data_to_upload.append(mapped_row)
        
        # Thêm hàng tiêu đề nếu cần (chỉ thêm nếu sheet chưa có dữ liệu)
        headers = ["", "Ngày", "Phòng", "Họ và tên HSSV", "Khoa", "Lớp", "Giáo viên giảng dạy", "Nề nếp Buổi"]

        # Lấy dữ liệu hiện tại từ Google Sheets để kiểm tra hàng rỗng
        range_name = f"{sheet_name}!A:H"  # Lấy dữ liệu từ cột A đến H
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        existing_values = result.get('values', [])

        # Tìm hàng rỗng tiếp theo (hàng mà tất cả các cột từ A đến H đều trống)
        next_row = 1
        for row in existing_values:
            # Nếu hàng không rỗng (có ít nhất 1 cột chứa dữ liệu), tăng next_row
            if any(cell for cell in row):
                next_row += 1
            else:
                break

        # Nếu sheet chưa có dữ liệu, thêm hàng tiêu đề
        if next_row == 1 and not existing_values:
            data_to_upload.insert(0, headers)
        else:
            # Nếu sheet đã có dữ liệu, kiểm tra hàng tiêu đề
            if next_row == 1:
                next_row = 2  # Bắt đầu từ hàng 2 nếu hàng 1 là tiêu đề

        # Xác định phạm vi để cập nhật
        sheet_range = f"{sheet_name}!A{next_row}"
        body = {"values": data_to_upload}
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=sheet_range,
            valueInputOption="RAW",
            body=body
        ).execute()
        
        logging.info(f"Successfully uploaded {file_path} to Google Sheets starting at row {next_row}")
        return True
    except Exception as e:
        logging.error(f"Error uploading {file_path} to Google Sheets: {e}")
        return False