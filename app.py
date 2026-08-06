import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
import plotly.express as px
from datetime import datetime

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG VÀ KẾT NỐI DATABASE (SUPABASE)
# ---------------------------------------------------------
st.set_page_config(page_title="Quản Lý Hóa Đơn & Giá Cả", layout="wide", initial_sidebar_state="expanded")

@st.cache_resource
def get_database_engine():
    db_url = st.secrets["postgres"]["url"]
    return create_engine(
        db_url, 
        pool_pre_ping=True, 
        pool_size=10, 
        max_overflow=20,
        pool_recycle=300
    )

engine = get_database_engine()

# Tự động tạo các bảng nếu chưa có
def init_db():
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS ten_chuan (
                id SERIAL PRIMARY KEY,
                ten_chuan TEXT UNIQUE NOT NULL
            );
        """))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS anh_xa (
                id SERIAL PRIMARY KEY,
                ten_npp TEXT UNIQUE NOT NULL,
                ten_chuan TEXT NOT NULL
            );
        """))
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS lich_su (
                id SERIAL PRIMARY KEY,
                nha_phan_phoi TEXT,
                so_hoa_don TEXT,
                ngay_nhap_hang TEXT,
                ten_sp_npp TEXT,
                ten_sp_chuan TEXT,
                quy_cach FLOAT,
                so_luong_thung FLOAT,
                don_gia_thung FLOAT,
                don_gia_le FLOAT,
                tong_tien FLOAT
            );
        """))

init_db()

# --- CÁC HÀM BỔ TRỢ ---
def find_col(df, possible_names):
    for col in df.columns:
        if str(col).strip().lower() in [p.lower() for p in possible_names]:
            return col
    return None

def parse_excel_invoice(file_path):
    df_raw = pd.read_excel(file_path, header=None)
    
    detected_npp = ""
    for r in range(min(10, len(df_raw))):
        row_str = " ".join([str(val) for val in df_raw.iloc[r].dropna().values])
        if "nhà cung cấp" in row_str.lower() or "npp" in row_str.lower() or "nhà phân phối" in row_str.lower():
            vals = [str(v).strip() for v in df_raw.iloc[r].dropna().values if str(v).strip()]
            if len(vals) >= 2:
                detected_npp = vals[-1]
            elif ":" in row_str:
                detected_npp = row_str.split(":")[-1].strip()
            break

    header_idx = None
    for r in range(min(15, len(df_raw))):
        row_vals = [str(v).strip().lower() for v in df_raw.iloc[r].dropna().values]
        if any(k in row_vals for k in ['tên hàng', 'tên sản phẩm', 'tên sp', 'diễn giải']):
            header_idx = r
            break
            
    if header_idx is None:
        header_idx = 0
        
    df_data = pd.read_excel(file_path, header=header_idx)
    col_name = find_col(df_data, ['tên hàng', 'tên sản phẩm', 'tên sp', 'diễn giải'])
    if col_name:
        df_data = df_data[df_data[col_name].notna()]
        df_data = df_data[~df_data[col_name].astype(str).str.lower().str.contains('tổng cộng|cộng|thành tiền')]
        
    return df_data, detected_npp

@st.cache_data(ttl=300, show_spinner=False)
def load_data_from_db():
    with engine.connect() as conn:
        try:
            df_lich_su = pd.read_sql(text("SELECT * FROM lich_su ORDER BY id DESC"), conn)
        except Exception:
            df_lich_su = pd.DataFrame()

        try:
            df_anh_xa = pd.read_sql(text("SELECT * FROM anh_xa"), conn)
        except Exception:
            df_anh_xa = pd.DataFrame()

        try:
            df_chuan = pd.read_sql(text("SELECT * FROM ten_chuan"), conn)
        except Exception:
            df_chuan = pd.DataFrame()
        
    if not df_lich_su.empty:
        col_npp = find_col(df_lich_su, ['nha_phan_phoi', 'npp'])
        col_ngay = find_col(df_lich_su, ['ngay_nhap_hang', 'ngay_nhap'])
        col_sp_npp = find_col(df_lich_su, ['ten_sp_npp', 'ten_npp'])
        col_sp_chuan = find_col(df_lich_su, ['ten_sp_chuan', 'ten_chuan'])
        col_qc = find_col(df_lich_su, ['quy_cach'])
        col_sl = find_col(df_lich_su, ['so_luong_thung', 'so_luong'])
        col_gia_thung = find_col(df_lich_su, ['don_gia_thung', 'don_gia'])
        col_gia_le = find_col(df_lich_su, ['don_gia_le'])
        col_tong = find_col(df_lich_su, ['tong_tien', 'thanh_tien'])

        if col_npp: df_lich_su['nha_phan_phoi'] = df_lich_su[col_npp]
        if col_ngay: df_lich_su['ngay_nhap_hang'] = df_lich_su[col_ngay]
        if col_sp_npp: df_lich_su['ten_sp_npp'] = df_lich_su[col_sp_npp]
        if col_sp_chuan: df_lich_su['ten_sp_chuan'] = df_lich_su[col_sp_chuan]
        if col_qc: df_lich_su['quy_cach'] = df_lich_su[col_qc]
        if col_sl: df_lich_su['so_luong_thung'] = df_lich_su[col_sl]
        if col_gia_thung: df_lich_su['don_gia_thung'] = df_lich_su[col_gia_thung]
        if col_gia_le: df_lich_su['don_gia_le'] = df_lich_su[col_gia_le]
        if col_tong: df_lich_su['tong_tien'] = df_lich_su[col_tong]

        if 'ngay_nhap_hang' in df_lich_su.columns:
            df_lich_su['ngay_dt'] = pd.to_datetime(df_lich_su['ngay_nhap_hang'], format='%d/%m/%Y', errors='coerce')
        else:
            df_lich_su['ngay_dt'] = pd.NaT

    return df_lich_su, df_anh_xa, df_chuan

def clear_app_cache():
    st.cache_data.clear()

df_lich_su, df_anh_xa, df_chuan = load_data_from_db()

st.title("📦 Quản Lý Nhập Hàng & So Sánh Giá")

# ---------------------------------------------------------
# 2. HỆ THỐNG TAB CHÍNH
# ---------------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📥 Nhập Hóa Đơn", 
    "📋 Danh Sách Hóa Đơn", 
    "🏷️ Quản Lý Tên Chuẩn", 
    "🔍 Chi Tiết & So Sánh Giá", 
    "📊 Báo Cáo & Biểu Đồ"
])

# =========================================================
# TAB 1: NHẬP HÓA ĐƠN
# =========================================================
with tab1:
    st.subheader("📥 Nhập Hóa Đơn Mới Từ Excel")
    
    col1, col2 = st.columns(2)
    with col1:
        ngay_nhap_selected = st.date_input("Ngày nhập hóa đơn:", value=datetime.now())
        ngay_nhap_str = ngay_nhap_selected.strftime("%d/%m/%Y")
    with col2:
        nha_phan_phoi_user = st.text_input("Tên Nhà Phân Phối (NPP - Tùy chọn):", placeholder="Để trống nếu muốn đọc từ file...")
        
    uploaded_file = st.file_uploader("Chọn file Excel hóa đơn (.xlsx, .xls)", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            df_upload, detected_npp = parse_excel_invoice(uploaded_file)
            final_npp = nha_phan_phoi_user.strip() if nha_phan_phoi_user.strip() else (detected_npp if detected_npp else "NPP Mới")
            
            st.success(f"Tải file thành công! [NPP: **{final_npp}** | Ngày: **{ngay_nhap_str}**]")
            st.dataframe(df_upload.head(10), use_container_width=True)
            
            if st.button("💾 Lưu Hóa Đơn Vào Hệ Thống", type="primary"):
                with st.spinner("Đang xử lý dữ liệu..."):
                    anh_xa_dict = dict(zip(df_anh_xa['ten_npp'], df_anh_xa['ten_chuan'])) if not df_anh_xa.empty and 'ten_npp' in df_anh_xa.columns else {}
                    
                    c_sp = find_col(df_upload, ['Tên hàng', 'Tên sản phẩm', 'Tên SP', 'Diễn giải'])
                    c_qc = find_col(df_upload, ['Quy cách đóng', 'Quy cách', 'Số lượng/Thùng'])
                    c_sl = find_col(df_upload, ['Số lượng', 'SL'])
                    c_gia = find_col(df_upload, ['Đơn giá', 'Giá thùng', 'Giá'])
                    
                    rows_to_insert = []
                    for idx, row in df_upload.iterrows():
                        ten_sp_npp = str(row.get(c_sp, '')).strip() if c_sp else ""
                        try: quy_cach = float(row.get(c_qc, 1)) if pd.notna(row.get(c_qc)) else 1.0
                        except: quy_cach = 1.0
                        
                        try: so_luong_thung = float(row.get(c_sl, 0)) if pd.notna(row.get(c_sl)) else 0.0
                        except: so_luong_thung = 0.0
                        
                        try: don_gia_thung = float(row.get(c_gia, 0)) if pd.notna(row.get(c_gia)) else 0.0
                        except: don_gia_thung = 0.0
                        
                        if not ten_sp_npp or so_luong_thung <= 0: continue
                            
                        don_gia_le = don_gia_thung / quy_cach if quy_cach > 0 else don_gia_thung
                        tong_tien = so_luong_thung * don_gia_thung
                        ten_sp_chuan = anh_xa_dict.get(ten_sp_npp, ten_sp_npp)

                        rows_to_insert.append({
                            "nha_phan_phoi": final_npp,
                            "so_hoa_don": "",
                            "ngay_nhap_hang": ngay_nhap_str,
                            "ten_sp_npp": ten_sp_npp,
                            "ten_sp_chuan": ten_sp_chuan,
                            "quy_cach": quy_cach,
                            "so_luong_thung": so_luong_thung,
                            "don_gia_thung": don_gia_thung,
                            "don_gia_le": don_gia_le,
                            "tong_tien": tong_tien
                        })
                    
                    if rows_to_insert:
                        df_insert = pd.DataFrame(rows_to_insert)
                        df_insert.to_sql("lich_su", engine, if_exists="append", index=False, method="multi", chunksize=1000)
                        clear_app_cache()
                        st.balloons()
                        st.success(f"Đã lưu thành công {len(rows_to_insert)} sản phẩm!")
                        st.rerun()
                    else:
                        st.warning("Không tìm thấy dòng dữ liệu hợp lệ trong file Excel.")
        except Exception as e:
            st.error(f"Lỗi xử lý file Excel: {e}")

# =========================================================
# TAB 2: DANH SÁCH HÓA ĐƠN & XÓA HÓA ĐƠN
# =========================================================
with tab2:
    st.subheader("📋 Danh Sách Hóa Đơn Đã Nhập")
    
    if not df_lich_su.empty and 'nha_phan_phoi' in df_lich_su.columns:
        df_lich_su_temp = df_lich_su.copy()
        
        c_filter1, c_filter2 = st.columns([2, 2])
        with c_filter1:
            list_npp = ["Tất cả"] + list(df_lich_su_temp['nha_phan_phoi'].dropna().unique())
            sel_npp = st.selectbox("Lọc theo Nhà Phân Phối:", list_npp)
        with c_filter2:
            sort_type = st.radio("Sắp xếp theo ngày:", ["Mới nhất trước", "Cũ nhất trước"], horizontal=True)
            
        if sel_npp != "Tất cả":
            df_show = df_lich_su_temp[df_lich_su_temp['nha_phan_phoi'] == sel_npp]
        else:
            df_show = df_lich_su_temp.copy()
            
        ascending = False if sort_type == "Mới nhất trước" else True
        df_show = df_show.sort_values(by=['ngay_dt', 'id'], ascending=[ascending, ascending])
            
        grouped_hd = df_show.groupby(['nha_phan_phoi', 'ngay_nhap_hang'], sort=False)
        
        for (npp_item, ngay_item), group in grouped_hd:
            tong_tien_hd = group['tong_tien'].sum() if 'tong_tien' in group.columns else 0
            so_luong_mon = len(group)
            
            col_exp, col_btn = st.columns([6, 1])
            with col_exp:
                with st.expander(f"📄 NPP: **{npp_item}** | Ngày: **{ngay_item}** | Tổng: **{tong_tien_hd:,.0f} đ** ({so_luong_mon} món)"):
                    cols_to_show = [c for c in ['ten_sp_npp', 'quy_cach', 'so_luong_thung', 'don_gia_thung', 'don_gia_le', 'tong_tien'] if c in group.columns]
                    st.dataframe(group[cols_to_show], use_container_width=True)
            with col_btn:
                if st.button("🗑️ Xóa", key=f"del_{npp_item}_{ngay_item}"):
                    with engine.begin() as conn:
                        conn.execute(text("DELETE FROM lich_su WHERE nha_phan_phoi = :n AND ngay_nhap_hang = :d"), {"n": npp_item, "d": ngay_item})
                    clear_app_cache()
                    st.success("Đã xóa hóa đơn!")
                    st.rerun()
    else:
        st.info("Chưa có dữ liệu hóa đơn nào trong hệ thống.")

# =========================================================
# TAB 3: QUẢN LÝ TÊN CHUẨN & ÁNH XẠ
# =========================================================
with tab3:
    st.subheader("🏷️ Quản Lý Tên Chuẩn & Bảng Ánh Xạ NPP")
    
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.markdown("##### 1. Thêm Tên Chuẩn Mới")
        ten_chuan_moi = st.text_input("Nhập tên sản phẩm chuẩn mới:")
        if st.button("Thêm Tên Chuẩn"):
            if ten_chuan_moi.strip():
                with engine.begin() as conn:
                    conn.execute(text("INSERT INTO ten_chuan (ten_chuan) VALUES (:t) ON CONFLICT DO NOTHING"), {"t": ten_chuan_moi.strip()})
                clear_app_cache()
                st.success(f"Đã thêm: {ten_chuan_moi}")
                st.rerun()
                
    with col_t2:
        st.markdown("##### 2. Khớp Tên NPP với Tên Chuẩn")
        if not df_lich_su.empty and 'ten_sp_npp' in df_lich_su.columns:
            ds_ten_npp_chua_ax = list(df_lich_su['ten_sp_npp'].unique())
            ten_npp_selected = st.selectbox("Chọn tên sản phẩm NPP:", ds_ten_npp_chua_ax)
            
            ds_chuan = list(df_chuan['ten_chuan'].unique()) if not df_chuan.empty and 'ten_chuan' in df_chuan.columns else []
            ten_chuan_selected = st.selectbox("Chọn Tên Chuẩn tương ứng:", ds_chuan)
            
            if st.button("Lưu Ánh Xạ"):
                if ten_npp_selected and ten_chuan_selected:
                    with engine.begin() as conn:
                        conn.execute(
                            text("INSERT INTO anh_xa (ten_npp, ten_chuan) VALUES (:n, :c) ON CONFLICT (ten_npp) DO UPDATE SET ten_chuan = :c"),
                            {"n": ten_npp_selected, "c": ten_chuan_selected}
                        )
                        conn.execute(
                            text("UPDATE lich_su SET ten_sp_chuan = :c WHERE ten_sp_npp = :n"),
                            {"n": ten_npp_selected, "c": ten_chuan_selected}
                        )
                    clear_app_cache()
                    st.success("Đã lưu ánh xạ thành công!")
                    st.rerun()

    st.markdown("---")
    st.markdown("##### Danh Sách Ánh Xạ Hiện Tại")
    if not df_anh_xa.empty:
        st.dataframe(df_anh_xa, use_container_width=True)

# =========================================================
# TAB 4: SO SÁNH GIÁ VỚI HÓA ĐƠN TRƯỚC ĐÓ
# =========================================================
with tab4:
    st.subheader("🔍 Chi Tiết & So Sánh Giá Hóa Đơn")
    
    if not df_lich_su.empty and 'nha_phan_phoi' in df_lich_su.columns and 'ngay_nhap_hang' in df_lich_su.columns:
        df_lich_su_t4 = df_lich_su.copy()
        df_hd_list = df_lich_su_t4[['nha_phan_phoi', 'ngay_nhap_hang', 'ngay_dt']].drop_duplicates().sort_values(by='ngay_dt', ascending=False)
        
        hd_options = [f"NPP: {row['nha_phan_phoi']} | Ngày: {row['ngay_nhap_hang']}" for _, row in df_hd_list.iterrows()]
        selected_hd = st.selectbox("Chọn hóa đơn cần xem chi tiết so sánh:", hd_options)
        
        if selected_hd:
            parts = selected_hd.split(" | ")
            curr_npp = parts[0].replace("NPP: ", "").strip()
            curr_ngay = parts[1].replace("Ngày: ", "").strip()
            
            curr_hd_df = df_lich_su_t4[(df_lich_su_t4['nha_phan_phoi'] == curr_npp) & (df_lich_su_t4['ngay_nhap_hang'] == curr_ngay)]
            curr_dt = curr_hd_df['ngay_dt'].iloc[0] if not curr_hd_df.empty else None
            
            prev_hd_all = df_lich_su_t4[df_lich_su_t4['ngay_dt'] < curr_dt] if curr_dt is not None else pd.DataFrame()
            
            result_rows = []
            for _, row in curr_hd_df.iterrows():
                ten_npp = row.get('ten_sp_npp', '')
                ten_chuan = row.get('ten_sp_chuan', ten_npp)
                gia_thung_curr = row.get('don_gia_thung', 0)
                
                prev_price_row = prev_hd_all[prev_hd_all['ten_sp_chuan'] == ten_chuan].sort_values(by='ngay_dt', ascending=False) if not prev_hd_all.empty and 'ten_sp_chuan' in prev_hd_all.columns else pd.DataFrame()
                
                if not prev_price_row.empty:
                    gia_thung_prev = prev_price_row.iloc[0].get('don_gia_thung', 0)
                    ngay_prev = prev_price_row.iloc[0].get('ngay_nhap_hang', '')
                    diff_thung = gia_thung_curr - gia_thung_prev
                    
                    if diff_thung > 0: trang_thai = f"🔴 TĂNG {diff_thung:,.0f}đ"
                    elif diff_thung < 0: trang_thai = f"🟢 GIẢM {abs(diff_thung):,.0f}đ"
                    else: trang_thai = "⚪ Không đổi"
                else:
                    gia_thung_prev = 0
                    ngay_prev = "Chưa có"
                    trang_thai = "🆕 Hàng mới"
                    
                result_rows.append({
                    "Tên Hàng Hóa": ten_npp,
                    "Số Lượng": row.get('so_luong_thung', 0),
                    "Đơn Giá Thùng Hiện Tại": gia_thung_curr,
                    "Đơn Giá Lẻ Hiện Tại": row.get('don_gia_le', 0),
                    "Đơn Giá Thùng Trước": gia_thung_prev,
                    "Ngày Nhập Trước": ngay_prev,
                    "Tăng/Giảm": trang_thai,
                    "Thành Tiền": row.get('tong_tien', 0)
                })
                
            st.dataframe(pd.DataFrame(result_rows), use_container_width=True)

# =========================================================
# TAB 5: BÁO CÁO & BIỂU ĐỒ
# =========================================================
with tab5:
    st.subheader("📊 Báo Cáo & Biểu Đồ Nhập Hàng")
    
    if not df_lich_su.empty and 'ngay_dt' in df_lich_su.columns:
        df_bc = df_lich_su.copy()
        time_option = st.selectbox("Chọn khoảng thời gian báo cáo:", ["Tháng này", "Tất cả thời gian"])
        
        now = datetime.now()
        if time_option == "Tháng này":
            df_bc_filtered = df_bc[(df_bc['ngay_dt'].dt.month == now.month) & (df_bc['ngay_dt'].dt.year == now.year)]
        else:
            df_bc_filtered = df_bc.copy()
            
        if not df_bc_filtered.empty and 'tong_tien' in df_bc_filtered.columns:
            tong_so_hd = df_bc_filtered.groupby(['nha_phan_phoi', 'ngay_nhap_hang']).ngroups if 'nha_phan_phoi' in df_bc_filtered.columns else len(df_bc_filtered)
            tong_tien_nhap = df_bc_filtered['tong_tien'].sum()
            tong_mat_hang = len(df_bc_filtered)
            
            c1, c2, c3 = st.columns(3)
            c1.metric("💰 Tổng tiền nhập", f"{tong_tien_nhap:,.0f} đ")
            c2.metric("📄 Tổng số HD", f"{tong_so_hd} hóa đơn")
            c3.metric("📦 Tổng mặt hàng", f"{tong_mat_hang} món")
            
            st.markdown("---")
            st.markdown("##### 📈 Biểu Đồ Chi Phí Nhập Hàng")
            
            df_chart = df_bc_filtered.groupby(['ngay_nhap_hang', 'ngay_dt'], as_index=False)['tong_tien'].sum().sort_values(by='ngay_dt')
            fig = px.bar(
                df_chart, 
                x='ngay_nhap_hang', 
                y='tong_tien', 
                text_auto=',.0f',
                labels={'ngay_nhap_hang': 'Ngày Nhập', 'tong_tien': 'Số Tiền (VNĐ)'},
                title="Tổng Chi Phí Nhập Hàng Theo Ngày"
            )
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Không có dữ liệu trong thời gian chọn.")
    else:
        st.info("Chưa có dữ liệu để lập báo cáo.")
