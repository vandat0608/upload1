# network_checker.py
import socket
import time
import requests

def check_network(timeout=5):
    """
    Kiểm tra kết nối mạng bằng cách ping đến Google DNS (8.8.8.8) và đo tốc độ phản hồi.
    Trả về tuple (is_connected, message).
    """
    try:
        # Kiểm tra kết nối bằng socket
        start_time = time.time()
        socket.setdefaulttimeout(timeout)
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex(('8.8.8.8', 53))  # Google DNS, port 53
        sock.close()
        
        latency = (time.time() - start_time) * 1000  # Đổi sang milliseconds
        
        if result == 0:
            if latency > 500:  # Nếu độ trễ > 500ms, coi là mạng chậm
                return True, f"Mạng chậm: Độ trễ {latency:.2f}ms vượt quá ngưỡng 500ms."
            return True, "Kết nối mạng ổn định."
        else:
            return False, "Không có kết nối mạng: Không thể ping đến 8.8.8.8."
    except socket.timeout:
        return False, "Lỗi mạng: Hết thời gian chờ khi kiểm tra kết nối (timeout)."
    except Exception as e:
        return False, f"Lỗi mạng: {str(e)}."

def check_internet_speed():
    """
    Kiểm tra tốc độ tải xuống từ một URL đơn giản.
    Trả về tuple (success, message).
    """
    try:
        url = "http://www.google.com"  # URL đơn giản để kiểm tra
        start_time = time.time()
        response = requests.get(url, timeout=5)
        response.raise_for_status()  # Kiểm tra lỗi HTTP
        elapsed = time.time() - start_time
        
        if elapsed > 2:  # Nếu tải lâu hơn 2 giây
            return True, f"Mạng chậm: Thời gian tải {elapsed:.2f}s vượt quá ngưỡng 2s."
        return True, "Tốc độ mạng ổn định."
    except requests.exceptions.RequestException as e:
        return False, f"Lỗi mạng: Không thể kết nối internet - {str(e)}."