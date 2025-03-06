// client/src/App.tsx
import React, { useState } from "react";
import axios, { AxiosError, AxiosResponse } from "axios";
import "./App.css"; // Import CSS hiện tại

// Định nghĩa kiểu dữ liệu cho response từ backend
interface ProcessResponse {
  status: string;
  successfulFiles?: number;
  totalRowsAdded?: number;
  error: boolean;
}

function App() {
  const [files, setFiles] = useState<File[]>([]);
  const [googleSheetUrl, setGoogleSheetUrl] = useState<string>("");
  const [sheetName, setSheetName] = useState<string>("");
  const [faculty, setFaculty] = useState<string>("");
  const [status, setStatus] = useState<{
    message: string;
    isError: boolean;
    isSuccess: boolean;
  }>({
    message: "",
    isError: false,
    isSuccess: false,
  });

  const validateInput = (fieldName: string, value: string) => {
    if (!value) {
      setStatus({
        message: `Lỗi: Vui lòng nhập ${fieldName}!`,
        isError: true,
        isSuccess: false,
      });
      return false;
    }
    return true;
  };

  const runProcess = async () => {
    if (
      !files ||
      !validateInput("Google Sheet URL", googleSheetUrl) ||
      !validateInput("Sheet Name", sheetName) ||
      !faculty
    ) {
      return;
    }

    setStatus({
      message: "Kiểm tra mạng...",
      isError: false,
      isSuccess: false,
    });
    try {
      const networkResponse = await axios.post(
        "https://upload1-xjte.onrender.com/check-network"
        // "http://127.0.0.1:5000/check-network"
      );
      if (networkResponse.data.error) {
        setStatus({
          message: `Lỗi: ${networkResponse.data.message} - Tạm dừng thực thi.`,
          isError: true,
          isSuccess: false,
        });
        return;
      }
      setStatus({
        message: networkResponse.data.message,
        isError: false,
        isSuccess: false,
      });
    } catch (error) {
      const axiosError = error as AxiosError;
      const errorMessage =
        axiosError.response?.data?.message ||
        axiosError.message ||
        "Lỗi không xác định";
      setStatus({
        message: `Lỗi: Không thể kiểm tra mạng - ${errorMessage}`,
        isError: true,
        isSuccess: false,
      });
      return;
    }

    setStatus({
      message: "Bắt đầu xử lý...",
      isError: false,
      isSuccess: false,
    });
    const formData = new FormData();
    files.forEach((file) => formData.append("files", file));
    formData.append("googleSheetUrl", googleSheetUrl);
    formData.append("sheetName", sheetName);
    formData.append("faculty", faculty);

    try {
      const response = await axios.post<ProcessResponse>(
        "https://upload1-xjte.onrender.com/process",
        // "http://127.0.0.1:5000/process",
        formData,
        {
          headers: { "Content-Type": "multipart/form-data" },
        }
      );

      setStatus({
        message: response.data.status,
        isError: false,
        isSuccess: response.data.successfulFiles !== undefined,
      });
      if (response.data.successfulFiles && response.data.totalRowsAdded) {
        setStatus({
          message: `Hoàn tất toàn bộ quá trình: Tổng số file được thêm thành công: ${response.data.successfulFiles}, Tổng số hàng được thêm: ${response.data.totalRowsAdded}`,
          isError: false,
          isSuccess: true,
        });
      }
    } catch (error) {
      const axiosError = error as AxiosError;
      const errorMessage =
        axiosError.response?.data?.message ||
        axiosError.message ||
        "Lỗi không xác định";
      setStatus({
        message: `Lỗi: ${errorMessage}`,
        isError: true,
        isSuccess: false,
      });
    }
  };

  return (
    <div className="app-container">
      <div className="content-frame">
        <h1>Ứng dụng chuyển Excel lên Google Sheets</h1>

        <div className="input-group">
          <label>Chọn file Excel (.xlsx):</label>
          <input
            type="file"
            multiple
            accept=".xlsx"
            onChange={(e) => setFiles(Array.from(e.target.files || []))}
          />
        </div>

        <div className="input-group">
          <label>URL Google Sheet:</label>
          <input
            type="text"
            value={googleSheetUrl}
            onChange={(e) => setGoogleSheetUrl(e.target.value)}
            placeholder="Nhập URL Google Sheet"
          />
        </div>

        <div className="input-group">
          <label>Tên Sheet:</label>
          <input
            type="text"
            value={sheetName}
            onChange={(e) => setSheetName(e.target.value)}
            placeholder="Nhập tên sheet trong Google Sheets"
          />
        </div>

        <div className="input-group">
          <label>Chọn Khoa:</label>
          <select value={faculty} onChange={(e) => setFaculty(e.target.value)}>
            <option value="">-- Chọn Khoa --</option>
            <option value="Khoa Công nghệ thông tin - Kỹ thuật điện">
              Khoa Công nghệ thông tin - Kỹ thuật điện
            </option>
            <option value="Khoa Du lịch - Khách sạn">
              Khoa Du lịch - Khách sạn
            </option>
            <option value="Khoa Cơ khí">Khoa Cơ khí</option>
            <option value="Khoa Kinh tế- Luật">Khoa Kinh tế- Luật</option>
            <option value="Khoa Chăm sóc sắc đẹp - Nuôi dưỡng trẻ">
              Khoa Chăm sóc sắc đẹp - Nuôi dưỡng trẻ
            </option>
            <option value="Khoa Y dược">Khoa Y dược</option>
            <option value="Khoa Ngoại Ngữ">Khoa Ngoại Ngữ</option>
            <option value="Khoa học cơ bản">Khoa học cơ bản</option>
          </select>
        </div>

        <button onClick={runProcess}>Thực Thi</button>

        {status.message && (
          <div
            className={`status ${
              status.isSuccess ? "success" : status.isError ? "error" : ""
            }`}
          >
            {status.message}
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
