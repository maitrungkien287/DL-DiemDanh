import streamlit as st
import face_recognition
import pandas as pd
import os
import pickle
import numpy as np
from datetime import datetime
import shutil
import socket
import qrcode
from PIL import Image

# --- CẤU HÌNH HỆ THỐNG ---
ROOT_FOLDER = "DU_LIEU_LOP_HOC"  # Thư mục gốc chứa tất cả các lớp
DATA_FILE_NAME = "data_khuon_mat.pickle"
EXCEL_NAME = "DanhSachLop.xlsx"

# Tạo thư mục gốc nếu chưa có
if not os.path.exists(ROOT_FOLDER):
    os.makedirs(ROOT_FOLDER)

# --- CÁC HÀM XỬ LÝ (BACKEND) ---

def get_local_ip():
    """Lấy địa chỉ IP nội bộ của máy tính để tạo QR Code"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "localhost"

def generate_qr(url):
    """Tạo ảnh QR Code từ URL"""
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    # Sửa lỗi hiển thị bằng cách chuyển sang hệ màu RGB
    return img.convert('RGB')

def get_list_classes():
    """Lấy danh sách các lớp hiện có"""
    if not os.path.exists(ROOT_FOLDER):
        return []
    return [d for d in os.listdir(ROOT_FOLDER) if os.path.isdir(os.path.join(ROOT_FOLDER, d))]

def load_class_data(class_name):
    """Load dữ liệu training của một lớp cụ thể"""
    file_path = os.path.join(ROOT_FOLDER, class_name, DATA_FILE_NAME)
    if not os.path.exists(file_path):
        return {"encodings": [], "ids": []}
    with open(file_path, "rb") as f:
        return pickle.load(f)

def save_class_data(class_name, data):
    """Lưu dữ liệu training của lớp"""
    file_path = os.path.join(ROOT_FOLDER, class_name, DATA_FILE_NAME)
    with open(file_path, "wb") as f:
        pickle.dump(data, f)

def delete_class(class_name):
    """Xóa toàn bộ dữ liệu của một lớp"""
    path = os.path.join(ROOT_FOLDER, class_name)
    if os.path.exists(path):
        shutil.rmtree(path)
        return True
    return False

def train_new_class(class_name, excel_file, uploaded_images):
    """Tạo lớp mới và train dữ liệu ngay lập tức"""
    class_path = os.path.join(ROOT_FOLDER, class_name)
    
    # 1. Tạo thư mục lớp
    if os.path.exists(class_path):
        return False, f"⚠️ Lớp '{class_name}' đã tồn tại!"
    os.makedirs(class_path)
    
    # 2. Lưu file Excel danh sách
    excel_path = os.path.join(class_path, EXCEL_NAME)
    with open(excel_path, "wb") as f:
        f.write(excel_file.getbuffer())
        
    # 3. Lưu ảnh và Train
    images_path = os.path.join(class_path, "Anh_Sinh_Vien")
    os.makedirs(images_path)
    
    data = {"encodings": [], "ids": []}
    count = 0
    
    # Progress bar
    progress_bar = st.progress(0)
    
    for i, uploaded_file in enumerate(uploaded_images):
        # Lưu file ảnh gốc
        file_path = os.path.join(images_path, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
            
        # Train (Trích xuất đặc trưng)
        try:
            mssv = os.path.splitext(uploaded_file.name)[0].strip() # Tên file là MSSV
            image = face_recognition.load_image_file(file_path)
            encs = face_recognition.face_encodings(image)
            if encs:
                data["encodings"].append(encs[0])
                data["ids"].append(mssv)
                count += 1
        except Exception as e:
            print(f"Lỗi ảnh {uploaded_file.name}: {e}")
            
        progress_bar.progress((i + 1) / len(uploaded_images))
        
    # 4. Lưu dữ liệu đã train
    save_class_data(class_name, data)
    return True, f"✅ Đã tạo lớp '{class_name}' và học xong {count} sinh viên!"

def process_attendance_excel(class_name, found_ids):
    """Xử lý file Excel: Giữ cột cũ, thêm cột điểm danh mới"""
    excel_path = os.path.join(ROOT_FOLDER, class_name, EXCEL_NAME)
    
    if not os.path.exists(excel_path):
        return None, "❌ Không tìm thấy file danh sách lớp!"

    try:
        # Đọc file để tìm header (như logic cũ)
        df_temp = pd.read_excel(excel_path, header=None)
        start_row = 0
        for index, row in df_temp.iterrows():
            row_str = row.astype(str).str.lower().values
            if any("mã sinh viên" in x for x in row_str) or any("mã sv" in x for x in row_str):
                start_row = index
                break
        
        df = pd.read_excel(excel_path, header=start_row)
        
        # Tìm cột MSSV
        mssv_col = next((col for col in df.columns if "mã sinh viên" in str(col).lower() or "mã sv" in str(col).lower()), None)
        if not mssv_col:
            return None, "❌ File Excel thiếu cột 'Mã sinh viên'!"

        # Thêm cột điểm danh mới
        today_col = f"DD_{datetime.now().strftime('%d/%m/%Y')}"
        
        # Chuẩn hóa và so sánh
        df[mssv_col] = df[mssv_col].astype(str).str.strip()
        found_list = [str(x).strip() for x in found_ids]
        
        # Logic: Có -> 'x', Vắng -> 'V'
        df[today_col] = df[mssv_col].apply(lambda x: "x" if x in found_list else "V")
        
        # Lưu đè lại file cũ (để cập nhật liên tục)
        df.to_excel(excel_path, index=False)
        
        return df, None
    except Exception as e:
        return None, str(e)

# --- GIAO DIỆN CHÍNH (FRONTEND) ---

st.set_page_config(page_title="Điểm Danh", page_icon="🎓", layout="wide")

# Sidebar quản lý lớp
with st.sidebar:
    st.title("🎓 QUẢN LÝ LỚP HỌC")
    st.write("---")
    
    mode = st.radio("Chức năng:", ["🏠 Điểm danh (Trang chủ)", "➕ Thêm lớp mới", "🗑️ Xóa lớp"])
    
    st.info("💡 Hướng dẫn: Thêm lớp trước -> Sau đó ra Trang chủ để điểm danh.")

# --- TRANG 1: THÊM LỚP MỚI ---
if mode == "➕ Thêm lớp mới":
    st.header("Tạo Lớp Học Mới & Nạp Dữ Liệu")
    
    new_class_name = st.text_input("Nhập tên lớp (Ví dụ: DH19ATT):")
    uploaded_excel = st.file_uploader("1. Tải file Excel danh sách lớp:", type=['xlsx'])
    uploaded_photos = st.file_uploader("2. Tải ảnh thẻ sinh viên (Đặt tên là MSSV):", accept_multiple_files=True, type=['jpg','png','jpeg'])
    
    if st.button("🚀 Tạo Lớp Ngay"):
        if new_class_name and uploaded_excel and uploaded_photos:
            with st.spinner("Đang khởi tạo lớp và training AI..."):
                success, msg = train_new_class(new_class_name, uploaded_excel, uploaded_photos)
            if success:
                st.success(msg)
                st.balloons()
            else:
                st.error(msg)
        else:
            st.warning("Vui lòng nhập đủ thông tin!")

# --- TRANG 2: XÓA LỚP ---
elif mode == "🗑️ Xóa lớp":
    st.header("Quản lý Xóa Lớp")
    classes = get_list_classes()
    if classes:
        selected_del = st.selectbox("Chọn lớp cần xóa:", classes)
        if st.button("❌ Xác nhận Xóa vĩnh viễn"):
            if delete_class(selected_del):
                st.success(f"Đã xóa lớp {selected_del}!")
                st.rerun()
    else:
        st.info("Chưa có lớp nào.")

# --- TRANG 3: ĐIỂM DANH (TRANG CHỦ) ---
elif mode == "🏠 Điểm danh (Trang chủ)":
    st.title("📸 ĐIỂM DANH THÔNG MINH")
    
    classes = get_list_classes()
    if not classes:
        st.warning("⚠️ Chưa có lớp nào! Vui lòng vào mục 'Thêm lớp mới' để tạo.")
    else:
        # 1. Chọn Lớp
        col1, col2 = st.columns([1, 2])
        with col1:
            selected_class = st.selectbox("📚 Chọn lớp đang dạy:", classes)
            
            # --- TÍNH NĂNG QR CODE ---
            st.write("---")
            st.subheader("📱 Quét QR để tải ảnh")
            ip = get_local_ip()
            url = f"http://{ip}:8501"
            st.image(generate_qr(url), caption="Dùng điện thoại quét mã này", width=200)
            st.code(url, language="text")
            st.info("Thầy/Cô quét mã để mở giao diện này trên điện thoại và tải ảnh lên.")

        with col2:
            st.subheader(f"Điểm danh lớp: {selected_class}")
            
            # 2. Upload ảnh (Dùng trên cả PC hoặc Điện thoại)
            uploaded_scan = st.file_uploader("📷 Tải lên ảnh chụp bao quát lớp:", accept_multiple_files=True, type=['jpg','png'])
            
            if st.button("🔍 Bắt đầu Điểm danh") and uploaded_scan:
                data = load_class_data(selected_class)
                
                if not data["encodings"]:
                    st.error("Lớp này chưa có dữ liệu khuôn mặt!")
                else:
                    found_ids = set()
                    with st.spinner("AI đang phân tích ảnh..."):
                        # Quét từng ảnh
                        for photo in uploaded_scan:
                            image = face_recognition.load_image_file(photo)
                            # Upsample=2 để nhìn rõ mặt xa
                            locs = face_recognition.face_locations(image, number_of_times_to_upsample=2)
                            encs = face_recognition.face_encodings(image, locs)
                            
                            for enc in encs:
                                matches = face_recognition.compare_faces(data["encodings"], enc, tolerance=0.45)
                                if True in matches:
                                    best = np.argmin(face_recognition.face_distance(data["encodings"], enc))
                                    if matches[best]:
                                        found_ids.add(data["ids"][best])
                    
                    # 3. Kết quả & Xuất Excel
                    st.success(f"✅ Đã tìm thấy {len(found_ids)} sinh viên!")
                    st.write(f"Danh sách: {', '.join(list(found_ids))}")
                    
                    # Cập nhật Excel
                    df_result, err = process_attendance_excel(selected_class, found_ids)
                    if df_result is not None:
                        st.dataframe(df_result.tail(10)) # Hiện 10 dòng cuối
                        
                        # Nút tải file về
                        excel_path = os.path.join(ROOT_FOLDER, selected_class, EXCEL_NAME)
                        with open(excel_path, "rb") as f:
                            st.download_button(
                                label="📥 Tải File Excel Kết Quả",
                                data=f,
                                file_name=f"DiemDanh_{selected_class}_{datetime.now().strftime('%d_%m')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error(f"Lỗi Excel: {err}")