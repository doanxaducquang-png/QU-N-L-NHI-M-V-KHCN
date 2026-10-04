import os
import json
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Hệ thống Quản lý Hồ sơ KH&CN", page_icon="📊", layout="wide")

GOOGLE_SHEET_NAME = "QuanLy_DonDangKy_KHCN" 

def get_gspread_client():
    """Khởi tạo kết nối Google Sheets an toàn cho cả Local và Cloud"""
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    
    # 1. Ưu tiên kiểm tra file local google_key.json trước
    if os.path.exists("google_key.json"):
        try:
            credentials = Credentials.from_service_account_file("google_key.json", scopes=scopes)
            return gspread.authorize(credentials)
        except Exception as e:
            st.error(f"Lỗi đọc file google_key.json: {e}")
            return None
            
    # 2. Nếu không có file local mới đọc st.secrets trên Cloud
    try:
        if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
            service_account_info = json.loads(st.secrets["gcp_service_account"]["json_key"])
            credentials = Credentials.from_service_account_info(service_account_info, scopes=scopes)
            return gspread.authorize(credentials)
    except Exception:
        pass
        
    st.warning("⚠️ Chưa tìm thấy file 'google_key.json' trong thư mục dự án!")
    return None

@st.cache_data(ttl=10)
def load_data_from_gsheet():
    try:
        gc = get_gspread_client()
        if not gc:
            return pd.DataFrame()
            
        sh = gc.open(GOOGLE_SHEET_NAME)
        ws = sh.worksheet("QUAN_LY_HO_SO")
        
        data = ws.get_all_values()
        if len(data) <= 1:
            return pd.DataFrame()
            
        headers = data[1] # Hàng 2 là tiêu đề cột
        rows = data[2:]   # Dữ liệu từ hàng 3 trở đi
        
        # 1. Xử lý tên cột trùng hoặc rỗng để tránh lỗi "Duplicate column names"
        clean_headers = []
        counts = {}
        for idx, h in enumerate(headers):
            name = str(h).strip()
            if not name:
                name = f"Unnamed_{idx}" # Đặt tên tạm cho các cột rỗng
            if name in counts:
                counts[name] += 1
                name = f"{name}_{counts[name]}" # Đánh số phân biệt cột trùng: Trạng thái hồ sơ_2
            else:
                counts[name] = 0
            clean_headers.append(name)
        
        df = pd.DataFrame(rows, columns=clean_headers)
        return df
    except Exception as e:
        st.error(f"Lỗi kết nối Google Sheets: {e}")
        return pd.DataFrame()

# Nạp dữ liệu
df_data = load_data_from_gsheet()

# Giao diện chính
st.title("📊 Báo cáo & Quản lý Hồ sơ KH&CN")

if not df_data.empty:
    st.success(f"✅ Đã kết nối Google Sheets thành công! Tổng số bản ghi: {len(df_data)}")
    st.dataframe(df_data)
else:
    st.info("Chưa có dữ liệu hoặc đang chờ kết nối đến Google Sheets...")