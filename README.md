# 🎓 HỆ THỐNG ĐIỂM DANH SINH VIÊN BẰNG KHUÔN MẶT (AI FACE ID)

> **Phiên bản:** 2.0 (Pro)
> **Tác giả:** [Tên Của Bạn]
> **Lớp/MSSV:** [Lớp Của Bạn] - [Mã Số Của Bạn]

---

## 📖 GIỚI THIỆU

Đây là ứng dụng điểm danh tự động sử dụng công nghệ Deep Learning (Face Recognition) để nhận diện khuôn mặt sinh viên từ ảnh chụp tập thể. Hệ thống hỗ trợ quản lý nhiều lớp học, điểm danh qua mã QR trên điện thoại và tự động xuất báo cáo Excel.

## 🚀 TÍNH NĂNG NỔI BẬT

- **Đa nền tảng:** Chạy Server trên máy tính, giao diện điều khiển trên cả PC và Điện thoại.
- **Quét QR Code:** Kết nối điện thoại với máy tính cực nhanh để chụp ảnh điểm danh.
- **Quản lý đa lớp:** Tự động tạo thư mục lưu trữ riêng cho từng lớp học.
- **Xử lý Excel thông minh:** Giữ nguyên định dạng file gốc của nhà trường, chỉ thêm cột điểm danh mới.
- **Zero-Setup:** Có file chạy tự động, không cần cấu hình thủ công.

---

## 🛠️ YÊU CẦU HỆ THỐNG

- Hệ điều hành: Windows 10/11.
- Python: Phiên bản 3.8 trở lên.
- Kết nối mạng: Máy tính và Điện thoại phải dùng chung một mạng Wifi (để quét QR).

---

## 📦 CẤU TRÚC THƯ MỤC

```text
Project_DiemDanh/
├── DU_LIEU_LOP_HOC/       # (Tự động tạo) Chứa dữ liệu các lớp
├── app.py                 # Source code chính (Python)
├── requirements.txt       # Danh sách thư viện cần thiết
├── START_APP.bat          # File khởi động nhanh (Launcher)
└── README.md              # Tài liệu hướng dẫn này
```
