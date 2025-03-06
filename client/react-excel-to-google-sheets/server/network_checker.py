import os
import socket
import speedtest

def check_network():
    try:
        # Check if we can resolve a domain name
        socket.gethostbyname("www.google.com")
        return True, "Network is available."
    except socket.error:
        return False, "Network is not available."

def check_internet_speed():
    try:
        st = speedtest.Speedtest()
        st.get_best_server()
        download_speed = st.download() / 1_000_000  # Convert to Mbps
        upload_speed = st.upload() / 1_000_000  # Convert to Mbps
        return True, f"Download speed: {download_speed:.2f} Mbps, Upload speed: {upload_speed:.2f} Mbps."
    except Exception as e:
        return False, f"Error checking internet speed: {str(e)}."