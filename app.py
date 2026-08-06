import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
import plotly.express as px
from datetime import datetime

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG VÀ KẾT NỐI DATABASE
# ---------------------------------------------------------
st.set_page_config(page_title="App Nhập Hàng", layout="wide", initial_sidebar_state="expanded")

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

# Đọc dữ liệu an toàn từ Database (Giữ nguyên toàn bộ dữ liệu cũ)
@st.cache_data(ttl=600, show_spinner=False)
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
        
    # Chuyển toàn bộ tên cột về chữ thường để tránh lỗi lệch tên cột giữa DB và Pandas
    if not df_lich_su.empty:
        df_lich_su.columns = [str(c).lower() for c in df_lich_su.columns]
        
        # Tìm cột chứa thông tin ngày nhập hàng
        col_date = next((c for c in ['ngay_nhap_hang', 'ngay_nhap', 'ngay'] if c in df_lich_su.columns), None)
        if col_date:
            df_lich_su['ngay_nhap_hang'] = df_lich_su[col_date]
            df_lich_su['ngay_dt'] = pd.to_datetime(df_lich_su['ngay_nhap_hang'], format='%d/%m/%Y', errors='coerce')
        else:
            df_lich_su['ngay_dt'] = pd.NaT

    if not df_anh_xa.empty:
        df_anh_xa.columns = [str(c).lower() for c in df_anh_xa.columns]

    if not df_chuan.empty:
        df_chuan.columns = [str(c).lower() for c in df_chuan.columns]
        
    return df_lich_su, df_anh_xa, df_chuan

def clear_app_cache():
    st.cache_data.clear()

with st.spinner("Đang tải dữ liệu từ máy chủ..."):
    df_lich_su, df_anh_xa, df_chuan = load_data_from_db()

st.title("📦 Phần Mềm Quản Lý Nhập Hàng & Giá Cả")

# ---------------------------------------------------------
# 2. KHỞI TẠO CÁC TAB
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
    
    col_input1, col_input2 = st.columns(2)
    with col_input1:
        nha_phan_phoi = st.text_input("Tên Nhà Phân Phối (NPP):", placeholder="Nhập tên NPP...")
    with col_input2:
        ngay_nhap_selected = st.date_input("Ngày nhập hóa đơn:", value=datetime.now())
        ngay_nhap_str = ngay_nhap_selected.strftime("%d/%m/%Y")
        
    uploaded_file = st.file_uploader("Tải lên file hóa đơn Excel (.xlsx, .xls)", type=["xlsx", "xls"])
    
    if uploaded_file and nha_phan_phoi.strip():
        try:
            df_upload = pd.read_excel(uploaded_file)
            st.success("Tải file thành công! Vui lòng kiểm tra lại danh sách sản phẩm bên dưới:")
            st.dataframe(df_upload.head(10), use_container_width=True)
            
            if st.button("💾 Lưu Hóa Đơn Vào Hệ Thống", type="primary"):
                with st.spinner("Đang xử lý và lưu dữ liệu..."):
                    anh_xa_dict = dict(zip(df_anh_xa['ten_npp'], df_anh_xa['ten_chuan'])) if not df_anh_xa.empty and 'ten_npp' in df_anh_xa.columns else {}
                    
                    rows_to_insert = []
                    for idx, row in df_upload.iterrows():
                        ten_sp_npp = str(row.get('Tên sản phẩm', row.get('Tên hàng', ''))).strip()
                        quy_cach = float(row.get('Quy cách', row.get('Số lượng/Thùng', 1)))
                        so_luong_thung = float(row.get('Số lượng', row.get('SL', 0)))
                        don_gia_thung = float(row.get('Đơn giá', row.get('Giá thùng', 0)))
                        
                        if not ten_sp_npp or so_luong_thung <= 0:
                            continue
                            
                        don_gia_le = don_gia_thung / quy_cach if quy_cach > 0 else don_gia_thung
                        tong_tien = so_luong_thung * don_gia_thung
                        ten_sp_chuan = anh_xa_dict.get(ten_sp_npp, ten_sp_npp)
                        
                        rows_to_insert.append({
                            "nha_phan_phoi": nha_phan_phoi.strip(),
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
                        st.success(f"Đã lưu thành công hóa đơn của {nha_phan_phoi} ngày {ngay_nhap_str}!")
                        st.rerun()
                    else:
                        st.warning("Không tìm thấy dòng dữ liệu hợp lệ trong file Excel.")
        except Exception as e:
            st.error(f"Lỗi xử lý file Excel: {e}")
    elif uploaded_file and not nha_phan_phoi.strip():
        st.info("Vui lòng nhập Tên Nhà Phân Phối trước khi lưu.")

# =========================================================
# TAB 2: DANH SÁCH HÓA ĐƠN
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
            
        if sort_type == "Mới nhất trước":
            df_show = df_show.sort_values(by=['ngay_dt', 'id'], ascending=[False, False]) if 'id' in df_show.columns else df_show.sort_values(by=['ngay_dt'], ascending=[False])
        else:
            df_show = df_show.sort_values(by=['ngay_dt', 'id'], ascending=[True, True]) if 'id' in df_show.columns else df_show.sort_values(by=['ngay_dt'], ascending=[True])
            
        grouped_hd = df_show.groupby(['nha_phan_phoi', 'ngay_nhap_hang'], sort=False)
        
        for (npp_item, ngay_item), group in grouped_hd:
            tong_tien_hd = group['tong_tien'].sum() if 'tong_tien' in group.columns else 0
            so_luong_mon = len(group)
            
            with st.expander(f"📄 NPP: **{npp_item}** | Ngày nhập: **{ngay_item}** | Tổng tiền: **{tong_tien_hd:,.0f} đ** ({so_luong_mon} mặt hàng)"):
                cols_to_show = [c for c in ['ten_sp_npp', 'quy_cach', 'so_luong_thung', 'don_gia_thung', 'don_gia_le', 'tong_tien'] if c in group.columns]
                df_view_hd = group[cols_to_show].copy()
                st.dataframe(df_view_hd, use_container_width=True)
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
                st.success(f"Đã thêm tên chuẩn: {ten_chuan_moi}")
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
# TAB 4: CHI TIẾT & SO SÁNH GIÁ
# =========================================================
with tab4:
    st.subheader("🔍 Chi Tiết & So Sánh Giá Hóa Đơn Gần Nhất")
    
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
            
            if curr_dt is not None:
                prev_hd_all = df_lich_su_t4[df_lich_su_t4['ngay_dt'] < curr_dt]
            else:
                prev_hd_all = pd.DataFrame()
            
            result_rows = []
            for _, row in curr_hd_df.iterrows():
                ten_npp = row.get('ten_sp_npp', '')
                ten_chuan = row.get('ten_sp_chuan', ten_npp)
                gia_thung_curr = row.get('don_gia_thung', 0)
                gia_le_curr = row.get('don_gia_le', 0)
                
                prev_price_row = prev_hd_all[prev_hd_all['ten_sp_chuan'] == ten_chuan].sort_values(by='ngay_dt', ascending=False) if not prev_hd_all.empty and 'ten_sp_chuan' in prev_hd_all.columns else pd.DataFrame()
                
                if not prev_price_row.empty:
                    gia_thung_prev = prev_price_row.iloc[0].get('don_gia_thung', 0)
                    ngay_prev = prev_price_row.iloc[0].get('ngay_nhap_hang', '')
                    diff_thung = gia_thung_curr - gia_thung_prev
                    
                    if diff_thung > 0:
                        trang_thai = f"🔴 TĂNG {diff_thung:,.0f}đ/thùng"
                    elif diff_thung < 0:
                        trang_thai = f"🟢 GIẢM {abs(diff_thung):,.0f}đ/thùng"
                    else:
                        trang_thai = "⚪ Không đổi"
                else:
                    gia_thung_prev = 0
                    ngay_prev = "Chưa có"
                    trang_thai = "🆕 Hàng mới nhập"
                    
                result_rows.append({
                    "Tên Hàng Hóa (NPP)": ten_npp,
                    "Số Lượng": row.get('so_luong_thung', 0),
                    "Đơn Giá Thùng Hiện Tại": gia_thung_curr,
                    "Đơn Giá Lẻ Hiện Tại": gia_le_curr,
                    "Đơn Giá Thùng Trước": gia_thung_prev,
                    "Ngày Nhập Trước": ngay_prev,
                    "Tăng/Giảm Giá Thùng": trang_thai,
                    "Thành Tiền": row.get('tong_tien', 0)
                })
                
            df_res = pd.DataFrame(result_rows)
            st.dataframe(df_res, use_container_width=True)

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
            
            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("💰 Tổng tiền nhập", f"{tong_tien_nhap:,.0f} đ")
            col_m2.metric("📄 Tổng số HD", f"{tong_so_hd} hóa đơn")
            col_m3.metric("📦 Tổng mặt hàng", f"{tong_mat_hang} món")
            
            st.markdown("---")
            st.markdown("##### 📈 Biểu Đồ Nhập Hàng Theo Ngày")
            
            if 'ngay_nhap_hang' in df_bc_filtered.columns:
                df_chart = df_bc_filtered.groupby(['ngay_nhap_hang', 'ngay_dt'], as_index=False)['tong_tien'].sum()
                df_chart = df_chart.sort_values(by='ngay_dt')
                
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
            st.info("Không có dữ liệu trong khoảng thời gian đã chọn.")
    else:
        st.info("Chưa có dữ liệu để lập báo cáo.")
