import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io
from datetime import datetime
from docx import Document
import gspread
from google.oauth2.service_account import Credentials

# -----------------------------------------------------------------------------
# CẤU HÌNH TRANG WEB
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Hệ thống Quản lý Hồ sơ KH&CN",
    page_icon="📊",
    layout="wide"
)

# Tên file Google Sheet
SPREADSHEET_NAME = "QuanLy_DonDangKy_KHCN"
SHEET_NAME = "QUAN_LY_HO_SO"

# Danh mục các tùy chọn chuẩn hóa
LIST_LINH_VUC = [
    "Khoa học tự nhiên", "Khoa học kỹ thuật và công nghệ", "Khoa học y, dược",
    "Khoa học nông nghiệp", "Khoa học xã hội", "Khoa học nhân văn", "Công nghệ chiến lược"
]
LIST_LOAI_HINH = [
    "Nghiên cứu cơ bản", "Nghiên cứu ứng dụng", "Phát triển công nghệ", "Phát triển giải pháp xã hội"
]
LIST_HT_TRIEN_KHAI_1 = ["Nhiệm vụ", "Cụm nhiệm vụ", "Chuỗi nhiệm vụ"]
LIST_HT_TRIEN_KHAI_2 = [
    "Thực hiện theo hình thức liên kết", 
    "Thực hiện theo hình thức hợp tác công tư", 
    "Không thuộc 02 trường hợp trên"
]
LIST_HT_XET = ["Đặt hàng", "Tài trợ"]
LIST_PHAN_LOAI_CN = [
    "Nhân tài về KH, CN & ĐMST", 
    "Nhà khoa học trẻ / Kỹ sư trẻ tài năng", 
    "Không thuộc hai trường hợp trên"
]
LIST_HOC_HAM = ["Giáo sư", "Phó Giáo sư", "Tiến sĩ", "Thạc sĩ", "Cử nhân"]
LIST_TRANG_THAI = ["Đang xét", "Đang thực hiện", "Đã hoàn thành"]
LIST_NEN_TANG_SO = ["Đã cập nhật", "Chưa cập nhật"]

# -----------------------------------------------------------------------------
# KẾT NỐI GOOGLE SHEETS VIA GSPREAD
# -----------------------------------------------------------------------------
@st.cache_resource
def get_gspread_client():
    """Kết nối tới Google Sheets qua Secrets của Streamlit hoặc File JSON cục bộ"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    # Kiểm tra nếu chạy trên Streamlit Cloud (Sử dụng st.secrets)
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    else:
        # Nếu chạy cục bộ Localhost (Sử dụng file JSON)
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
        
    return gspread.authorize(creds)

def get_worksheet():
    client = get_gspread_client()
    spreadsheet = client.open(SPREADSHEET_NAME)
    try:
        worksheet = spreadsheet.worksheet(SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        # Nếu chưa có trang tính tên QUAN_LY_HO_SO thì lấy trang đầu tiên
        worksheet = spreadsheet.sheet1
    return worksheet

# -----------------------------------------------------------------------------
# 1. HÀM ĐỌC & TRÍCH XUẤT TỰ ĐỘNG TỪ BIỂU MẪU WORD (BM-09)
# -----------------------------------------------------------------------------
def is_checked(text):
    check_symbols = ["☒", "☑", "■", "[x]", "[X]", "(x)", "(X)", "v", "V"]
    for sym in check_symbols:
        if sym in text:
            return True
    return False

def parse_bm09_word(doc_file):
    doc = Document(doc_file)
    
    full_paragraphs = [p.text for p in doc.paragraphs]
    table_texts = []
    for table in doc.tables:
        for row in table.rows:
            row_str = " | ".join([cell.text.strip() for cell in row.cells])
            table_texts.append(row_str)
            
    all_lines = full_paragraphs + table_texts
    full_text = "\n".join(all_lines)

    extracted = {}

    # 1. Tên nhiệm vụ
    match_ten = re.search(r"1\.\ Tên nhiệm vụ/cụm nhiệm vụ/chuỗi nhiệm vụ:\s*(.*?)(?=\n|Thuộc lĩnh vực:|$)", full_text, re.DOTALL)
    extracted["ten_nhiem_vu"] = match_ten.group(1).strip().strip("…").strip() if match_ten else ""

    # 2. Lĩnh vực
    extracted["linh_vuc"] = LIST_LINH_VUC[0]
    for lv in LIST_LINH_VUC:
        for line in all_lines:
            if lv.lower() in line.lower() and is_checked(line):
                extracted["linh_vuc"] = lv
                break

    # 3. Loại hình
    extracted["loai_hinh"] = LIST_LOAI_HINH[0]
    for lh in LIST_LOAI_HINH:
        for line in all_lines:
            if lh.lower() in line.lower() and is_checked(line):
                extracted["loai_hinh"] = lh
                break

    # 4. Hình thức triển khai 1 & 2
    extracted["ht_tk1"] = "Nhiệm vụ"
    extracted["ht_tk2"] = "Không thuộc 02 trường hợp trên"
    
    if "3.2. Cụm nhiệm vụ:" in full_text and is_checked(full_text.split("3.2. Cụm nhiệm vụ:")[1][:50]):
        extracted["ht_tk1"] = "Cụm nhiệm vụ"
    elif "3.3. Chuỗi nhiệm vụ:" in full_text and is_checked(full_text.split("3.3. Chuỗi nhiệm vụ:")[1][:50]):
        extracted["ht_tk1"] = "Chuỗi nhiệm vụ"

    if "Thực hiện theo hình thức liên kết" in full_text and is_checked(full_text.split("Thực hiện theo hình thức liên kết")[1][:30]):
        extracted["ht_tk2"] = "Thực hiện theo hình thức liên kết"
    elif "Thực hiện theo hình thức hợp tác công tư" in full_text and is_checked(full_text.split("Thực hiện theo hình thức hợp tác công tư")[1][:30]):
        extracted["ht_tk2"] = "Thực hiện theo hình thức hợp tác công tư"

    # 5. Hình thức xét
    extracted["ht_xet"] = "Đặt hàng"
    if "Tài trợ" in full_text and is_checked(full_text.split("Tài trợ")[0][-10:] if "Tài trợ" in full_text else ""):
        extracted["ht_xet"] = "Tài trợ"

    # 6. Tổ chức chủ trì
    match_tc = re.search(r"5\.\ Tổ chức chủ trì:\s*\n?Tên:\s*(.*?)\n", full_text)
    extracted["ten_tc"] = match_tc.group(1).strip().strip("…").strip() if match_tc else ""

    match_dc = re.search(r"Địa chỉ:\s*(.*?)\n", full_text)
    extracted["dc_tc"] = match_dc.group(1).strip().strip("…").strip() if match_dc else ""

    match_dt_tc = re.search(r"Điện thoại:\s*(.*?)(?=Email:|Website:|\n|$)", full_text)
    extracted["dt_tc"] = match_dt_tc.group(1).strip().strip("…").strip() if match_dt_tc else ""

    # 7. Người đại diện pháp luật
    match_dd = re.search(r"6\.\ Người đại diện theo pháp luật\s*\n?Họ tên:\s*(.*?)(?=Chức vụ:|\n|$)", full_text)
    extracted["dd_pl"] = match_dd.group(1).strip().strip("…").strip() if match_dd else ""

    match_cv = re.search(r"Chức vụ:\s*(.*?)(?=\n|$)", full_text)
    extracted["chuc_vu_dd"] = match_cv.group(1).strip().strip("…").strip() if match_cv else ""

    # 8. Trạng thái Nền tảng số
    extracted["tl_phap_ly"] = "Chưa cập nhật"
    extracted["tt_tc_de_xuat"] = "Chưa cập nhật"
    if "6.1. Tài liệu chứng minh" in full_text:
        part61 = full_text.split("6.1. Tài liệu chứng minh")[1].split("6.2.")[0]
        if "Đã cập nhật" in part61 and is_checked(part61.split("Đã cập nhật")[1][:20]):
            extracted["tl_phap_ly"] = "Đã cập nhật"

    if "6.2. Thông tin về tổ chức đề xuất" in full_text:
        part62 = full_text.split("6.2. Thông tin về tổ chức đề xuất")[1].split("7.")[0]
        if "Đã cập nhật" in part62 and is_checked(part62.split("Đã cập nhật")[1][:20]):
            extracted["tt_tc_de_xuat"] = "Đã cập nhật"

    # 9. Chủ nhiệm
    match_cn = re.search(r"7\.\ Cá nhân đăng ký chủ nhiệm nhiệm vụ:.*?\nHọ tên:\s*(.*?)(?=Học hàm|\n|$)", full_text, re.DOTALL)
    extracted["ho_ten_cn"] = match_cn.group(1).strip().strip("…").strip() if match_cn else ""

    match_hh = re.search(r"Học hàm, học vị:\s*(.*?)(?=Đơn vị công tác|\n|$)", full_text)
    hh_str = match_hh.group(1).strip().strip("…").strip() if match_hh else ""
    extracted["hoc_ham_hoc_vi"] = "Tiến sĩ"
    for hh in LIST_HOC_HAM:
        if hh.lower() in hh_str.lower():
            extracted["hoc_ham_hoc_vi"] = hh
            break

    match_dv = re.search(r"Đơn vị công tác:\s*(.*?)(?=Email:|\n|$)", full_text)
    extracted["don_vi_ct"] = match_dv.group(1).strip().strip("…").strip() if match_dv else ""

    match_email = re.search(r"Email:\s*(.*?)(?=Điện thoại:|\n|$)", full_text)
    extracted["email_cn"] = match_email.group(1).strip().strip("…").strip() if match_email else ""

    match_dt_cn = re.search(r"Điện thoại:\s*(.*?)(?=\n|Lý lịch|$)", full_text)
    extracted["dt_cn"] = match_dt_cn.group(1).strip().strip("…").strip() if match_dt_cn else ""

    extracted["phan_loai_cn"] = "Không thuộc hai trường hợp trên"
    if "Nhân tài về khoa học" in full_text and is_checked(full_text.split("Nhân tài về khoa học")[1][:30]):
        extracted["phan_loai_cn"] = "Nhân tài về KH, CN & ĐMST"
    elif "Nhà khoa học trẻ tài năng" in full_text and is_checked(full_text.split("Nhà khoa học trẻ tài năng")[1][:30]):
        extracted["phan_loai_cn"] = "Nhà khoa học trẻ / Kỹ sư trẻ tài năng"

    extracted["ly_lich_cn"] = "Chưa cập nhật"
    if "Lý lịch cá nhân đăng ký chủ nhiệm" in full_text:
        part_ll = full_text.split("Lý lịch cá nhân đăng ký chủ nhiệm")[1].split("8.")[0]
        if "Đã cập nhật" in part_ll and is_checked(part_ll.split("Đã cập nhật")[1][:20]):
            extracted["ly_lich_cn"] = "Đã cập nhật"

    # 10. Kinh phí
    match_kp = re.search(r"8\.\ Tổng kinh phí đề xuất\s*([\d\.,]+)", full_text)
    if match_kp:
        raw_kp = match_kp.group(1).replace(".", "").replace(",", "")
        extracted["tong_kinh_phi"] = float(raw_kp) if raw_kp.isdigit() else 0.0
    else:
        extracted["tong_kinh_phi"] = 0.0

    match_nsnn_pct = re.search(r"Ngân sách nhà nước:\s*([\d\.,]+)\s*%", full_text)
    pct_nsnn = float(match_nsnn_pct.group(1).replace(",", ".")) if match_nsnn_pct else 0.0
    extracted["kinh_phi_nsnn"] = (extracted["tong_kinh_phi"] * pct_nsnn) / 100.0
    extracted["kinh_phi_ngoai_nsnn"] = max(0.0, extracted["tong_kinh_phi"] - extracted["kinh_phi_nsnn"])

    # 11. Thời gian
    match_tg = re.search(r"9\.\ Thời gian thực hiện:\s*(\d+)", full_text)
    extracted["thoi_gian_th"] = int(match_tg.group(1)) if match_tg else 12

    return extracted

# -----------------------------------------------------------------------------
# 2. XỬ LÝ DỮ LIỆU GOOGLE SHEETS
# -----------------------------------------------------------------------------
@st.cache_data(ttl=5) # Tự động làm mới cache sau 5 giây
def load_data():
    try:
        ws = get_worksheet()
        all_values = ws.get_all_values()
        
        if len(all_values) < 2:
            return pd.DataFrame()
        
        # Tiêu đề cột nằm ở dòng 2 (Index 1)
        headers = all_values[1]
        data_rows = all_values[2:]
        
        # Loại bỏ các cột trống vô danh (nếu có)
        clean_headers = [h.strip() for h in headers if h.strip() != ""]
        
        df = pd.DataFrame(data_rows)
        # Chỉ lấy số lượng cột tương ứng với headers
        df = df.iloc[:, :len(headers)]
        df.columns = headers
        
        # Bỏ các dòng hoàn toàn trống
        df = df.dropna(how='all')
        df = df[df.astype(str).join(axis=1).str.strip() != ""]
        
        num_cols = ["STT", "Tổng kinh phí đề xuất (VNĐ)", "Kinh phí NSNN (VNĐ)", "Kinh phí Ngoài NSNN (VNĐ)", "Thời gian thực hiện (Tháng)"]
        for col in num_cols:
            if col in df.columns:
                # Xử lý định dạng chuỗi số từ Google Sheet (xóa dấu phẩy, chấm)
                df[col] = df[col].astype(str).str.replace(",", "").str.replace(".", "")
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
        pct_cols = ["Tỷ lệ NSNN (%)", "Tỷ lệ Ngoài NSNN (%)"]
        for col in pct_cols:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace("%", "").str.replace(",", ".")
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
                
        return df
    except Exception as e:
        st.error(f"Lỗi khi đọc dữ liệu từ Google Sheets: {e}")
        return pd.DataFrame()

def save_new_record(new_row_dict):
    try:
        ws = get_worksheet()
        headers = ws.row_values(2) # Lấy danh sách cột từ dòng thứ 2
        
        row_values = []
        for h in headers:
            val = new_row_dict.get(h, "")
            row_values.append(str(val))
            
        ws.append_row(row_values)
        st.cache_data.clear()
        return True, "Thêm hồ sơ mới và cập nhật thành công lên Google Sheets!"
    except Exception as e:
        return False, f"Lỗi khi lưu dữ liệu lên Google Sheets: {str(e)}"

# Khởi tạo Session State
if "form_data" not in st.session_state:
    st.session_state.form_data = {}
if "show_validation_errors" not in st.session_state:
    st.session_state.show_validation_errors = False

# -----------------------------------------------------------------------------
# 3. GIAO DIỆN CHÍNH
# -----------------------------------------------------------------------------
st.title("📌 HỆ THỐNG QUẢN LÝ & BÁO CÁO HỒ SƠ ĐĂNG KÝ KH&CN")
st.markdown("---")

col_head1, col_head2 = st.columns([8, 2])
with col_head2:
    if st.button("🔄 Cập nhật dữ liệu từ Google Sheets"):
        st.cache_data.clear()
        st.rerun()

df_data = load_data()

tab_dashboard, tab_report, tab_search, tab_input = st.tabs([
    "📊 Dashboard Trực Quan", 
    "📑 Báo Cáo Tổng Hợp", 
    "🔍 Tra Cứu & Quản Lý Hồ Sơ", 
    "➕ Nhập Hồ Sơ Mới"
])

# -----------------------------------------------------------------------------
# TAB 1: DASHBOARD TRỰC QUAN
# -----------------------------------------------------------------------------
with tab_dashboard:
    st.header("📊 DASHBOARD TỔNG QUAN")
    if df_data.empty:
        st.warning("Chưa có dữ liệu trong Google Sheets hoặc không thể kết nối.")
    else:
        total_hoso = len(df_data)
        total_budget = df_data["Tổng kinh phí đề xuất (VNĐ)"].sum()
        total_nsnn = df_data["Kinh phí NSNN (VNĐ)"].sum()
        total_ngoai_nsnn = df_data["Kinh phí Ngoài NSNN (VNĐ)"].sum()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Tổng số hồ sơ", f"{total_hoso} hồ sơ")
        col2.metric("Tổng kinh phí đề xuất", f"{total_budget:,.0f} VNĐ")
        col3.metric("Kinh phí NSNN", f"{total_nsnn:,.0f} VNĐ")
        col4.metric("Kinh phí Ngoài NSNN", f"{total_ngoai_nsnn:,.0f} VNĐ")
        
        st.markdown("---")
        
        row1_col1, row1_col2 = st.columns(2)
        with row1_col1:
            st.subheader("1. Cơ cấu Kinh phí theo Lĩnh vực Nghiên cứu")
            df_lv = df_data.groupby("Lĩnh vực")[["Kinh phí NSNN (VNĐ)", "Kinh phí Ngoài NSNN (VNĐ)"]].sum().reset_index()
            df_lv = df_lv[(df_lv["Kinh phí NSNN (VNĐ)"] > 0) | (df_lv["Kinh phí Ngoài NSNN (VNĐ)"] > 0)]
            fig_lv = px.bar(
                df_lv, x="Lĩnh vực", y=["Kinh phí NSNN (VNĐ)", "Kinh phí Ngoài NSNN (VNĐ)"],
                title="Kinh phí NSNN & Ngoài NSNN theo Lĩnh vực", barmode="stack", text_auto='.2s',
                color_discrete_sequence=["#1f77b4", "#ff7f0e"]
            )
            fig_lv.update_layout(xaxis_title="", yaxis_title="Kinh phí (VNĐ)", legend_title="Nguồn kinh phí")
            st.plotly_chart(fig_lv, use_container_width=True)

        with row1_col2:
            st.subheader("2. Số lượng hồ sơ theo Loại hình Nhiệm vụ")
            df_lh = df_data[df_data["Loại hình nhiệm vụ"].notna()]["Loại hình nhiệm vụ"].value_counts().reset_index()
            df_lh.columns = ["Loại hình", "Số lượng"]
            fig_lh = px.pie(df_lh, names="Loại hình", values="Số lượng", title="Tỷ lệ Loại hình Nhiệm vụ", hole=0.4)
            st.plotly_chart(fig_lh, use_container_width=True)

        st.markdown("---")
        row2_col1, row2_col2 = st.columns(2)
        
        with row2_col1:
            st.subheader("3. Trình độ Chủ nhiệm Đề tài")
            df_hh_clean = df_data[df_data["Học hàm, học vị"].notna()].copy()
            df_hh_clean["Học hàm, học vị"] = df_hh_clean["Học hàm, học vị"].astype(str).str.strip()
            counts_hh = df_hh_clean["Học hàm, học vị"].value_counts()
            
            df_hh = pd.DataFrame({"Học hàm/Học vị": LIST_HOC_HAM})
            df_hh["Số lượng"] = df_hh["Học hàm/Học vị"].map(counts_hh).fillna(0).astype(int)
            
            fig_hh = px.bar(
                df_hh, x="Học hàm/Học vị", y="Số lượng", color="Học hàm/Học vị", text="Số lượng",
                category_orders={"Học hàm/Học vị": LIST_HOC_HAM}
            )
            fig_hh.update_layout(showlegend=False, yaxis_title="Số lượng hồ sơ")
            st.plotly_chart(fig_hh, use_container_width=True)

        with row2_col2:
            st.subheader("4. Phân loại Chủ nhiệm Nhiệm vụ")
            df_pl_clean = df_data[df_data["Phân loại Chủ nhiệm"].notna()].copy()
            df_pl_clean["Phân loại Chủ nhiệm"] = df_pl_clean["Phân loại Chủ nhiệm"].astype(str).str.strip()
            
            def normalize_phan_loai(val):
                if "Nhân tài" in val:
                    return "Nhân tài về KH, CN & ĐMST"
                elif "Nhà khoa học trẻ" in val or "Kỹ sư trẻ" in val:
                    return "Nhà khoa học trẻ / Kỹ sư trẻ tài năng"
                return "Không thuộc hai trường hợp trên"

            df_pl_clean["Phân loại Chuẩn"] = df_pl_clean["Phân loại Chủ nhiệm"].apply(normalize_phan_loai)
            df_pl = df_pl_clean["Phân loại Chuẩn"].value_counts().reset_index()
            df_pl.columns = ["Phân loại", "Số lượng"]
            
            fig_pl = px.pie(df_pl, names="Phân loại", values="Số lượng", title="Phân loại Chủ nhiệm")
            st.plotly_chart(fig_pl, use_container_width=True)

        st.markdown("---")
        row3_col1, row3_col2, row3_col3 = st.columns(3)
        with row3_col1:
            st.subheader("5. Hình thức Triển khai 1")
            df_ht1 = df_data[df_data["Hình thức triển khai 1"].notna()]["Hình thức triển khai 1"].value_counts().reset_index()
            df_ht1.columns = ["Hình thức 1", "Số lượng"]
            fig_ht1 = px.bar(df_ht1, x="Hình thức 1", y="Số lượng", color="Hình thức 1")
            fig_ht1.update_layout(showlegend=False)
            st.plotly_chart(fig_ht1, use_container_width=True)
            
        with row3_col2:
            st.subheader("6. Hình thức Xét")
            df_xet = df_data[df_data["Hình thức xét"].notna()]["Hình thức xét"].value_counts().reset_index()
            df_xet.columns = ["Hình thức xét", "Số lượng"]
            fig_xet = px.pie(df_xet, names="Hình thức xét", values="Số lượng", hole=0.3)
            st.plotly_chart(fig_xet, use_container_width=True)

        with row3_col3:
            st.subheader("7. Trạng thái Hồ sơ")
            df_tt = df_data[df_data["Trạng thái hồ sơ"].notna()]["Trạng thái hồ sơ"].value_counts().reset_index()
            df_tt.columns = ["Trạng thái", "Số lượng"]
            fig_tt = px.bar(df_tt, x="Trạng thái", y="Số lượng", color="Trạng thái")
            fig_tt.update_layout(showlegend=False)
            st.plotly_chart(fig_tt, use_container_width=True)

# -----------------------------------------------------------------------------
# TAB 2: BÁO CÁO TỔNG HỢP
# -----------------------------------------------------------------------------
with tab_report:
    st.header("📑 BÁO CÁO TỔNG HỢP HỒ SƠ ĐĂNG KÝ CHỦ TRÌ NHIỆM VỤ KH&CN")
    st.caption("Dữ liệu được thống kê tự động dựa trên dữ liệu Google Sheets.")
    
    if not df_data.empty:
        st.subheader("I. TỔNG QUAN HỒ SƠ & KINH PHÍ")
        df_overview = pd.DataFrame([
            {"Chỉ tiêu": "1. Tổng số hồ sơ tiếp nhận", "Giá trị": f"{len(df_data)} hồ sơ"},
            {"Chỉ tiêu": "2. Tổng kinh phí đề xuất", "Giá trị": f"{df_data['Tổng kinh phí đề xuất (VNĐ)'].sum():,.0f} VNĐ"},
            {"Chỉ tiêu": "3. Tổng kinh phí Ngân sách nhà nước (NSNN)", "Giá trị": f"{df_data['Kinh phí NSNN (VNĐ)'].sum():,.0f} VNĐ"},
            {"Chỉ tiêu": "4. Tổng kinh phí Ngoài NSNN", "Giá trị": f"{df_data['Kinh phí Ngoài NSNN (VNĐ)'].sum():,.0f} VNĐ"},
        ])
        st.table(df_overview)

        st.subheader("II. THỐNG KÊ THEO LĨNH VỰC NGHIÊN CỨU")
        df_lv_stat = []
        for idx, lv in enumerate(LIST_LINH_VUC, 1):
            sub = df_data[df_data["Lĩnh vực"] == lv]
            df_lv_stat.append({
                "STT & Lĩnh vực": f"{idx}. {lv}",
                "Số lượng hồ sơ": len(sub),
                "Tổng kinh phí (VNĐ)": f"{sub['Tổng kinh phí đề xuất (VNĐ)'].sum():,.0f}",
                "Kinh phí NSNN tương ứng (VNĐ)": f"{sub['Kinh phí NSNN (VNĐ)'].sum():,.0f}",
                "Kinh phí Ngoài NSNN tương ứng (VNĐ)": f"{sub['Kinh phí Ngoài NSNN (VNĐ)'].sum():,.0f}"
            })
        st.dataframe(pd.DataFrame(df_lv_stat), use_container_width=True, hide_index=True)

        st.subheader("III. THỐNG KÊ THEO LOẠI HÌNH NHIỆM VỤ")
        df_lh_stat = []
        for idx, lh in enumerate(LIST_LOAI_HINH, 1):
            sub = df_data[df_data["Loại hình nhiệm vụ"] == lh]
            df_lh_stat.append({
                "STT & Loại hình": f"{idx}. {lh}",
                "Số lượng hồ sơ": len(sub),
                "Tổng kinh phí (VNĐ)": f"{sub['Tổng kinh phí đề xuất (VNĐ)'].sum():,.0f}",
                "Kinh phí NSNN tương ứng (VNĐ)": f"{sub['Kinh phí NSNN (VNĐ)'].sum():,.0f}",
                "Kinh phí Ngoài NSNN tương ứng (VNĐ)": f"{sub['Kinh phí Ngoài NSNN (VNĐ)'].sum():,.0f}"
            })
        st.dataframe(pd.DataFrame(df_lh_stat), use_container_width=True, hide_index=True)

        col_r1, col_r2 = st.columns(2)
        with col_r1:
            st.subheader("IV. PHÂN LOẠI CHỦ NHIỆM NHIỆM VỤ")
            df_cn_stat = []
            for idx, pl in enumerate(LIST_PHAN_LOAI_CN, 1):
                count = len(df_data[df_data["Phân loại Chủ nhiệm"] == pl])
                df_cn_stat.append({"Phân loại Chủ nhiệm": f"{idx}. {pl}", "Số lượng hồ sơ": count})
            st.dataframe(pd.DataFrame(df_cn_stat), use_container_width=True, hide_index=True)

        with col_r2:
            st.subheader("V. TÌNH TRẠNG CẬP NHẬT NỀN TẢNG SỐ QUỐC GIA")
            full_updated = len(df_data[
                (df_data["Tài liệu chứng minh tư cách pháp lý (Nền tảng số)"] == "Đã cập nhật") &
                (df_data["Thông tin về tổ chức đề xuất (Nền tảng số)"] == "Đã cập nhật") &
                (df_data["Lý lịch cá nhân (Nền tảng số)"] == "Đã cập nhật")
            ])
            not_full = len(df_data) - full_updated
            st.table(pd.DataFrame([
                {"Tình trạng": "1. Hồ sơ đã cập nhật hoàn thiện (Cả 3 mục)", "Số lượng hồ sơ": full_updated},
                {"Tình trạng": "2. Hồ sơ chưa cập nhật đủ thông tin", "Số lượng hồ sơ": not_full}
            ]))

        col_r3, col_r4 = st.columns(2)
        with col_r3:
            st.subheader("VI. HÌNH THỨC TRIỂN KHAI 1")
            df_ht1_stat = []
            for idx, ht in enumerate(LIST_HT_TRIEN_KHAI_1, 1):
                count = len(df_data[df_data["Hình thức triển khai 1"] == ht])
                df_ht1_stat.append({"Hình thức triển khai 1": f"{idx}. {ht}", "Số lượng hồ sơ": count})
            st.dataframe(pd.DataFrame(df_ht1_stat), use_container_width=True, hide_index=True)

        with col_r4:
            st.subheader("VII. HÌNH THỨC TRIỂN KHAI 2")
            df_ht2_stat = []
            for idx, ht in enumerate(LIST_HT_TRIEN_KHAI_2, 1):
                count = len(df_data[df_data["Hình thức triển khai 2"] == ht])
                df_ht2_stat.append({"Hình thức triển khai 2": f"{idx}. {ht}", "Số lượng hồ sơ": count})
            st.dataframe(pd.DataFrame(df_ht2_stat), use_container_width=True, hide_index=True)

        col_r5, col_r6, col_r7 = st.columns(3)
        with col_r5:
            st.subheader("VIII. HÌNH THỨC XẾT")
            df_xet_stat = []
            for idx, hx in enumerate(LIST_HT_XET, 1):
                count = len(df_data[df_data["Hình thức xét"] == hx])
                df_xet_stat.append({"Hình thức xét": f"{idx}. {hx}", "Số lượng hồ sơ": count})
            st.dataframe(pd.DataFrame(df_xet_stat), use_container_width=True, hide_index=True)

        with col_r6:
            st.subheader("IX. TRÌNH ĐỘ CHỦ NHIỆM ĐỀ TÀI")
            df_hh_stat = []
            for idx, hh in enumerate(LIST_HOC_HAM, 1):
                count = len(df_data[df_data["Học hàm, học vị"] == hh])
                df_hh_stat.append({"Học hàm / Học vị": f"{idx}. {hh}", "Số lượng hồ sơ": count})
            st.dataframe(pd.DataFrame(df_hh_stat), use_container_width=True, hide_index=True)

        with col_r7:
            st.subheader("X. TRẠNG THÁI HỒ SƠ")
            df_tt_stat = []
            for idx, tt in enumerate(LIST_TRANG_THAI, 1):
                count = len(df_data[df_data["Trạng thái hồ sơ"] == tt])
                df_tt_stat.append({"Trạng thái hồ sơ": f"{idx}. {tt}", "Số lượng hồ sơ": count})
            st.dataframe(pd.DataFrame(df_tt_stat), use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# TAB 3: TRA CỨU & XUẤT EXCEL TỪ GOOGLE SHEETS
# -----------------------------------------------------------------------------
with tab_search:
    st.header("🔍 TRA CỨU HỒ SƠ & XUẤT DỮ LIỆU EXCEL")
    
    if not df_data.empty:
        # Xuất dữ liệu Google Sheets thành file Excel tự động
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_data.to_excel(writer, index=False, sheet_name='QUAN_LY_HO_SO')
        buffer.seek(0)
        
        st.download_button(
            label="📥 Tải File Excel Dữ Liệu Hiện Tại (Xuất từ Google Sheets)",
            data=buffer,
            file_name=f"QuanLy_DonDangKy_KHCN_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        st.markdown("---")

        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            search_keyword = st.text_input("🔑 Tìm theo tên nhiệm vụ / Chủ nhiệm / Đơn vị:")
        with col_s2:
            filter_linhvuc = st.multiselect("📂 Lọc theo Lĩnh vực:", options=LIST_LINH_VUC)
        with col_s3:
            filter_trangthai = st.multiselect("📌 Lọc theo Trạng thái:", options=LIST_TRANG_THAI)

        df_filtered = df_data.copy()
        if search_keyword:
            kw = search_keyword.lower()
            df_filtered = df_filtered[
                df_filtered["Tên nhiệm vụ / Cụm / Chuỗi"].astype(str).str.lower().str.contains(kw) |
                df_filtered["Họ tên Chủ nhiệm"].astype(str).str.lower().str.contains(kw) |
                df_filtered["Tên tổ chức chủ trì"].astype(str).str.lower().str.contains(kw)
            ]
        if filter_linhvuc:
            df_filtered = df_filtered[df_filtered["Lĩnh vực"].isin(filter_linhvuc)]
        if filter_trangthai:
            df_filtered = df_filtered[df_filtered["Trạng thái hồ sơ"].isin(filter_trangthai)]

        st.write(f"**Kết quả tìm kiếm:** {len(df_filtered)} / {len(df_data)} hồ sơ")
        st.dataframe(df_filtered, use_container_width=True)
    else:
        st.info("Chưa có hồ sơ nào trong hệ thống.")

# -----------------------------------------------------------------------------
# TAB 4: NHẬP HỒ SƠ MỚI & TRÍCH XUẤT AUTO TỪ WORD (BM-09)
# -----------------------------------------------------------------------------
with tab_input:
    st.header("➕ NHẬP HỒ SƠ MỚI VÀO HỆ THỐNG")
    
    st.subheader("📄 TỰ ĐỘNG ĐIỀN TỪ FILE WORD ĐƠN ĐĂNG KÝ (BM-09)")
    uploaded_docx = st.file_uploader("Tải lên file Word Đơn đăng ký (.docx):", type=["docx"])
    
    if uploaded_docx is not None:
        try:
            parsed_data = parse_bm09_word(uploaded_docx)
            for k, v in parsed_data.items():
                st.session_state.form_data[k] = v
            st.success("🎉 Đã đọc file thành công! Dữ liệu BM-09 đã tự động điền vào Form.")
        except Exception as e:
            st.error(f"Lỗi khi đọc file Word: {e}")

    st.markdown("---")
    next_stt = int(df_data["STT"].max() + 1) if not df_data.empty and "STT" in df_data.columns else 1
    st.info(f"💡 **Thông báo:** Số thứ tự (STT) tiếp theo được hệ thống tự động gán là: **{next_stt}**")

    def check_error(key_name):
        if st.session_state.show_validation_errors:
            val = st.session_state.form_data.get(key_name)
            if val is None or (isinstance(val, str) and str(val).strip() == "") or (isinstance(val, (int, float)) and val <= 0):
                st.caption("⚠️ :red[Trường này là bắt buộc, vui lòng nhập bổ sung!]")

    fd = st.session_state.form_data

    # --- FORM NHẬP LIỆU ---
    st.subheader("1. Thông tin chung Nhiệm vụ")
    c1, c2 = st.columns(2)
    with c1:
        ten_nhiem_vu = st.text_area("Tên nhiệm vụ / Cụm / Chuỗi (*)", value=fd.get("ten_nhiem_vu", ""), placeholder="Nhập tên nhiệm vụ...")
        fd["ten_nhiem_vu"] = ten_nhiem_vu
        check_error("ten_nhiem_vu")

        idx_lv = LIST_LINH_VUC.index(fd["linh_vuc"]) if fd.get("linh_vuc") in LIST_LINH_VUC else 0
        linh_vuc = st.selectbox("Lĩnh vực (*)", options=LIST_LINH_VUC, index=idx_lv)
        fd["linh_vuc"] = linh_vuc

        idx_lh = LIST_LOAI_HINH.index(fd["loai_hinh"]) if fd.get("loai_hinh") in LIST_LOAI_HINH else 0
        loai_hinh = st.selectbox("Loại hình nhiệm vụ (*)", options=LIST_LOAI_HINH, index=idx_lh)
        fd["loai_hinh"] = loai_hinh

    with c2:
        idx_ht1 = LIST_HT_TRIEN_KHAI_1.index(fd["ht_tk1"]) if fd.get("ht_tk1") in LIST_HT_TRIEN_KHAI_1 else 0
        ht_tk1 = st.selectbox("Hình thức triển khai 1 (*)", options=LIST_HT_TRIEN_KHAI_1, index=idx_ht1)
        fd["ht_tk1"] = ht_tk1

        idx_ht2 = LIST_HT_TRIEN_KHAI_2.index(fd["ht_tk2"]) if fd.get("ht_tk2") in LIST_HT_TRIEN_KHAI_2 else 0
        ht_tk2 = st.selectbox("Hình thức triển khai 2 (*)", options=LIST_HT_TRIEN_KHAI_2, index=idx_ht2)
        fd["ht_tk2"] = ht_tk2

        idx_xet = LIST_HT_XET.index(fd["ht_xet"]) if fd.get("ht_xet") in LIST_HT_XET else 0
        ht_xet = st.selectbox("Hình thức xét (*)", options=LIST_HT_XET, index=idx_xet)
        fd["ht_xet"] = ht_xet

        idx_tt = LIST_TRANG_THAI.index(fd["trang_thai"]) if fd.get("trang_thai") in LIST_TRANG_THAI else 0
        trang_thai = st.selectbox("Trạng thái hồ sơ (*)", options=LIST_TRANG_THAI, index=idx_tt)
        fd["trang_thai"] = trang_thai

    st.markdown("---")
    st.subheader("2. Thông tin Tổ chức Chủ trì")
    c3, c4 = st.columns(2)
    with c3:
        ten_tc = st.text_input("Tên tổ chức chủ trì (*)", value=fd.get("ten_tc", ""))
        fd["ten_tc"] = ten_tc
        check_error("ten_tc")

        dc_tc = st.text_input("Địa chỉ tổ chức (*)", value=fd.get("dc_tc", ""))
        fd["dc_tc"] = dc_tc
        check_error("dc_tc")

        dt_tc = st.text_input("Điện thoại tổ chức (Không bắt buộc)", value=fd.get("dt_tc", ""))
        fd["dt_tc"] = dt_tc

    with c4:
        dd_pl = st.text_input("Người đại diện pháp luật (*)", value=fd.get("dd_pl", ""))
        fd["dd_pl"] = dd_pl
        check_error("dd_pl")

        chuc_vu_dd = st.text_input("Chức vụ người đại diện (*)", value=fd.get("chuc_vu_dd", ""))
        fd["chuc_vu_dd"] = chuc_vu_dd
        check_error("chuc_vu_dd")

        idx_tl = LIST_NEN_TANG_SO.index(fd["tl_phap_ly"]) if fd.get("tl_phap_ly") in LIST_NEN_TANG_SO else 0
        tl_phap_ly = st.selectbox("Tài liệu chứng minh tư cách pháp lý (Nền tảng số) (*)", options=LIST_NEN_TANG_SO, index=idx_tl)
        fd["tl_phap_ly"] = tl_phap_ly

        idx_tt_tc = LIST_NEN_TANG_SO.index(fd["tt_tc_de_xuat"]) if fd.get("tt_tc_de_xuat") in LIST_NEN_TANG_SO else 0
        tt_tc_de_xuat = st.selectbox("Thông tin về tổ chức đề xuất (Nền tảng số) (*)", options=LIST_NEN_TANG_SO, index=idx_tt_tc)
        fd["tt_tc_de_xuat"] = tt_tc_de_xuat

    st.markdown("---")
    st.subheader("3. Thông tin Chủ nhiệm Nhiệm vụ")
    c5, c6 = st.columns(2)
    with c5:
        ho_ten_cn = st.text_input("Họ tên Chủ nhiệm (*)", value=fd.get("ho_ten_cn", ""))
        fd["ho_ten_cn"] = ho_ten_cn
        check_error("ho_ten_cn")

        idx_pl_cn = LIST_PHAN_LOAI_CN.index(fd["phan_loai_cn"]) if fd.get("phan_loai_cn") in LIST_PHAN_LOAI_CN else 0
        phan_loai_cn = st.selectbox("Phân loại Chủ nhiệm (*)", options=LIST_PHAN_LOAI_CN, index=idx_pl_cn)
        fd["phan_loai_cn"] = phan_loai_cn

        idx_hh = LIST_HOC_HAM.index(fd["hoc_ham_hoc_vi"]) if fd.get("hoc_ham_hoc_vi") in LIST_HOC_HAM else 0
        hoc_ham_hoc_vi = st.selectbox("Học hàm, học vị (*)", options=LIST_HOC_HAM, index=idx_hh)
        fd["hoc_ham_hoc_vi"] = hoc_ham_hoc_vi

        don_vi_ct = st.text_input("Đơn vị công tác (*)", value=fd.get("don_vi_ct", ""))
        fd["don_vi_ct"] = don_vi_ct
        check_error("don_vi_ct")

    with c6:
        email_cn = st.text_input("Email Chủ nhiệm (Không bắt buộc)", value=fd.get("email_cn", ""))
        fd["email_cn"] = email_cn

        dt_cn = st.text_input("Điện thoại Chủ nhiệm (*)", value=fd.get("dt_cn", ""))
        fd["dt_cn"] = dt_cn
        check_error("dt_cn")

        idx_ll = LIST_NEN_TANG_SO.index(fd["ly_lich_cn"]) if fd.get("ly_lich_cn") in LIST_NEN_TANG_SO else 0
        ly_lich_cn = st.selectbox("Lý lịch cá nhân (Nền tảng số) (*)", options=LIST_NEN_TANG_SO, index=idx_ll)
        fd["ly_lich_cn"] = ly_lich_cn

    st.markdown("---")
    st.subheader("4. Kinh phí & Thời gian Thực hiện")
    
    auto_calc = st.checkbox("🔄 Tự động tính Kinh phí Ngoài NSNN (= Tổng kinh phí - Kinh phí NSNN)", value=True)
    
    c7, c8 = st.columns(2)
    with c7:
        tong_kinh_phi = st.number_input("Tổng kinh phí đề xuất (VNĐ) (*)", min_value=0.0, value=float(fd.get("tong_kinh_phi", 0.0)), step=1000000.0, format="%.0f")
        fd["tong_kinh_phi"] = tong_kinh_phi
        check_error("tong_kinh_phi")

        kinh_phi_nsnn = st.number_input("Kinh phí Ngân sách Nhà nước (NSNN) (VNĐ) (*)", min_value=0.0, value=float(fd.get("kinh_phi_nsnn", 0.0)), step=1000000.0, format="%.0f")
        fd["kinh_phi_nsnn"] = kinh_phi_nsnn

    with c8:
        if auto_calc:
            default_ngoai = max(0.0, tong_kinh_phi - kinh_phi_nsnn)
            kinh_phi_ngoai_nsnn = st.number_input("Kinh phí Ngoài NSNN (VNĐ) (Tự động tính)", min_value=0.0, value=default_ngoai, step=1000000.0, format="%.0f", disabled=True)
        else:
            kinh_phi_ngoai_nsnn = st.number_input("Kinh phí Ngoài NSNN (VNĐ) (*)", min_value=0.0, value=float(fd.get("kinh_phi_ngoai_nsnn", 0.0)), step=1000000.0, format="%.0f")
        fd["kinh_phi_ngoai_nsnn"] = kinh_phi_ngoai_nsnn

        thoi_gian_th = st.number_input("Thời gian thực hiện (Tháng) (*)", min_value=1, value=int(fd.get("thoi_gian_th", 12)))
        fd["thoi_gian_th"] = thoi_gian_th

    ty_le_nsnn = (kinh_phi_nsnn / tong_kinh_phi) if tong_kinh_phi > 0 else 0.0
    ty_le_ngoai_nsnn = (kinh_phi_ngoai_nsnn / tong_kinh_phi) if tong_kinh_phi > 0 else 0.0

    st.info(f"📊 **Hệ thống tự động tính tỷ lệ:** Tỷ lệ NSNN: **{ty_le_nsnn*100:.2f}%** | Tỷ lệ Ngoài NSNN: **{ty_le_ngoai_nsnn*100:.2f}%**")

    st.markdown("---")
    
    if st.button("💾 LƯU HỒ SƠ VÀO GOOGLE SHEETS", type="primary", use_container_width=True):
        st.session_state.show_validation_errors = True
        
        required_fields = [
            ("ten_nhiem_vu", ten_nhiem_vu),
            ("ten_tc", ten_tc),
            ("dc_tc", dc_tc),
            ("dd_pl", dd_pl),
            ("chuc_vu_dd", chuc_vu_dd),
            ("ho_ten_cn", ho_ten_cn),
            ("don_vi_ct", don_vi_ct),
            ("dt_cn", dt_cn),
            ("tong_kinh_phi", tong_kinh_phi)
        ]
        
        has_empty_error = False
        for k, v in required_fields:
            if v is None or (isinstance(v, str) and v.strip() == "") or (isinstance(v, (int, float)) and v <= 0):
                has_empty_error = True
                break
                
        is_budget_valid = abs(tong_kinh_phi - (kinh_phi_nsnn + kinh_phi_ngoai_nsnn)) < 1.0

        if has_empty_error:
            st.error("🚫 Vui lòng điền đầy đủ các trường dữ liệu bắt buộc (Có thông báo màu đỏ)!")
        elif not is_budget_valid:
            st.error(f"🚫 **Lỗi logic kinh phí:** Tổng kinh phí ({tong_kinh_phi:,.0f} VNĐ) không bằng (Kinh phí NSNN + Kinh phí Ngoài NSNN)!")
        else:
            new_data = {
                "STT": next_stt,
                "Tên nhiệm vụ / Cụm / Chuỗi": ten_nhiem_vu,
                "Lĩnh vực": linh_vuc,
                "Loại hình nhiệm vụ": loai_hinh,
                "Hình thức triển khai 1": ht_tk1,
                "Hình thức triển khai 2": ht_tk2,
                "Hình thức xét": ht_xet,
                "Tên tổ chức chủ trì": ten_tc,
                "Địa chỉ tổ chức": dc_tc,
                "Điện thoại tổ chức": dt_tc,
                "Người đại diện pháp luật": dd_pl,
                "Chức vụ người đại diện": chuc_vu_dd,
                "Tài liệu chứng minh tư cách pháp lý (Nền tảng số)": tl_phap_ly,
                "Thông tin về tổ chức đề xuất (Nền tảng số)": tt_tc_de_xuat,
                "Phân loại Chủ nhiệm": phan_loai_cn,
                "Họ tên Chủ nhiệm": ho_ten_cn,
                "Học hàm, học vị": hoc_ham_hoc_vi,
                "Đơn vị công tác": don_vi_ct,
                "Email Chủ nhiệm": email_cn,
                "Điện thoại Chủ nhiệm": dt_cn,
                "Lý lịch cá nhân (Nền tảng số)": ly_lich_cn,
                "Tổng kinh phí đề xuất (VNĐ)": tong_kinh_phi,
                "Tỷ lệ NSNN (%)": f"{ty_le_nsnn*100:.2f}%",
                "Kinh phí NSNN (VNĐ)": kinh_phi_nsnn,
                "Tỷ lệ Ngoài NSNN (%)": f"{ty_le_ngoai_nsnn*100:.2f}%",
                "Kinh phí Ngoài NSNN (VNĐ)": kinh_phi_ngoai_nsnn,
                "Thời gian thực hiện (Tháng)": thoi_gian_th,
                "Trạng thái hồ sơ": trang_thai
            }
            
            success, msg = save_new_record(new_data)
            if success:
                st.success(msg)
                st.balloons()
                st.session_state.form_data = {}
                st.session_state.show_validation_errors = False
                st.rerun()
            else:
                st.error(msg)