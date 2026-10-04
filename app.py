import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
import json
from datetime import datetime
from docx import Document

# Thư viện kết nối Google Sheets
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Hệ thống Quản lý Hồ sơ KH&CN", page_icon="📊", layout="wide")

GOOGLE_SHEET_NAME = "QuanLy_DonDangKy_KHCN" 

def get_gspread_client():
    """Khởi tạo kết nối Google Sheets an toàn cho cả Local và Cloud"""
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    
    # 1. Ưu tiên kiểm tra file local trên máy tính trước
    if os.path.exists("google_key.json"):
        credentials = Credentials.from_service_account_file("google_key.json", scopes=scopes)
    # 2. Nếu không có file local thì mới đọc từ Secrets trên Streamlit Cloud
    elif hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
        service_account_info = json.loads(st.secrets["gcp_service_account"]["json_key"])
        credentials = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    else:
        st.error("Chưa tìm thấy chìa khóa kết nối Google Sheets! (Cần file google_key.json ở Local hoặc cài Secrets trên Cloud)")
        return None
        
    return gspread.authorize(credentials)

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
        
        df = pd.DataFrame(rows, columns=headers)
        
        num_cols = ["STT", "Tổng kinh phí đề xuất (VNĐ)", "Kinh phí NSNN (VNĐ)", "Kinh phí Ngoài NSNN (VNĐ)", "Thời gian thực hiện (Tháng)"]
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                
        return df
    except Exception as e:
        st.error(f"Lỗi kết nối Google Sheets: {e}")
        return pd.DataFrame()

# Tải dữ liệu chính
df_data = load_data_from_gsheet()