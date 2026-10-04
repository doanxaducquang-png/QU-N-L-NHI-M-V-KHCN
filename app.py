import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Thư viện kết nối Google Sheets
import gspread
from google.oauth2.service_account import Credentials

# -----------------------------------------------------------------------------
# CẤU HÌNH TRANG WEB & KẾT NỐI GOOGLE SHEETS
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Hệ thống Quản lý Hồ sơ KH&CN", page_icon="📊", layout="wide")

# Đổi tên này trùng với Tên File Google Sheet bạn vừa tạo trên web
GOOGLE_SHEET_NAME = "QuanLy_DonDangKy_KHCN" 

import json

def get_gspread_client():
    """Khởi tạo kết nối Google Sheets từ Streamlit Secrets hoặc File Local"""
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    
    # 1. Nếu chạy trên Streamlit Cloud
    if "gcp_service_account" in st.secrets:
        # Chuyển chuỗi JSON trong secrets thành dictionary
        service_account_info = json.loads(st.secrets["gcp_service_account"]["json_key"])
        credentials = Credentials.from_service_account_info(
            service_account_info, scopes=scopes
        )
    # 2. Nếu chạy Local máy tính
    elif os.path.exists("google_key.json"):
        credentials = Credentials.from_service_account_file("google_key.json", scopes=scopes)
    else:
        st.error("Chưa tìm thấy khóa kết nối Google Sheets!")
        return None
        
    return gspread.authorize(credentials)

@st.cache_data(ttl=10)
def load_data_from_gsheet():
    """Đọc dữ liệu từ Google Sheet"""
    try:
        gc = get_gspread_client()
        if not gc:
            return pd.DataFrame()
            
        sh = gc.open(GOOGLE_SHEET_NAME)
        ws = sh.worksheet("QUAN_LY_HO_SO")
        
        data = ws.get_all_values()
        if len(data) <= 1:
            return pd.DataFrame()
            
        headers = data[1] # Hàng 2 làm Tiêu đề
        rows = data[2:]   # Dữ liệu từ hàng 3 trở đi
        
        df = pd.DataFrame(rows, columns=headers)
        
        # Ép kiểu dữ liệu số
        num_cols = ["STT", "Tổng kinh phí đề xuất (VNĐ)", "Kinh phí NSNN (VNĐ)", "Kinh phí Ngoài NSNN (VNĐ)", "Thời gian thực hiện (Tháng)"]
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                
        return df
    except Exception as e:
        st.error(f"Lỗi đọc dữ liệu từ Google Sheets: {e}")
        return pd.DataFrame()

def save_record_to_gsheet(new_row_dict):
    """Ghi thêm 1 dòng hồ sơ mới vào Google Sheet"""
    try:
        gc = get_gspread_client()
        if not gc:
            return False, "Không thể kết nối Google Sheets."
            
        sh = gc.open(GOOGLE_SHEET_NAME)
        ws = sh.worksheet("QUAN_LY_HO_SO")
        
        headers = ws.row_values(2)
        row_values = [str(new_row_dict.get(h, "")) for h in headers]
        
        ws.append_row(row_values)
        st.cache_data.clear()
        return True, "Lưu thành công lên Google Sheets!"
    except Exception as e:
        return False, f"Lỗi ghi dữ liệu: {e}"

# Tải dữ liệu lên ứng dụng
df_data = load_data_from_gsheet()

# -----------------------------------------------------------------------------
# GIAO DIỆN STREAMLIT
# -----------------------------------------------------------------------------
st.title("📌 HỆ THỐNG QUẢN LÝ & BÁO CÁO HỒ SƠ KH&CN (DỮ LIỆU ĐÁM MÂY)")
st.write(f"Số lượng hồ sơ hiện có: **{len(df_data)}**")
if not df_data.empty:
    st.dataframe(df_data, use_container_width=True)