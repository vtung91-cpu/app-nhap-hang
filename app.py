import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
import plotly.express as px
from datetime import datetime

# 1. CẤU HÌNH MÀN HÌNH
st.set_page_config(page_title="App Nhập Hàng", page_icon="📦", layout="centered")

st.title("📦 QUẢN LÝ NHẬP HÀNG (LƯU ĐÁM MÂY SUPABASE)")

# KẾT NỐI VỚI CƠ SỞ DỮ LIỆU POSTGRESQL SUPABASE
try:
    DB_URL = st.secrets["postgres"]["url"]
    engine = create_engine(DB_URL, pool_pre_ping=True)
except Exception as e:
    st.error("⚠️ Chưa cấu hình kết nối Supabase trong Streamlit Secrets! Vui lòng kiểm tra lại Secrets.")
    st.stop()

# TỰ ĐỘNG KHỞI TẠO BẢNG DỮ LIỆU TRÊN SUPABASE (NẾU CHƯA CÓ)
def init_db():
    with engine.begin() as conn:
        # Bảng Lịch sử nhập hàng
        conn.execute(text('''
            CREATE TABLE IF NOT EXISTS lich_su (
                id SERIAL PRIMARY KEY,
                ngay_hd TEXT,
                so_hd TEXT,
                ten_ncc TEXT,
                ten_phu_npp TEXT,
                ten_sp_chuan TEXT,
                quy_cach DOUBLE PRECISION,
                so_luong DOUBLE PRECISION,
                don_gia_thung DOUBLE PRECISION,
                gia_nhap_le DOUBLE PRECISION
            );
        '''))
        # Bảng Ánh xạ tên
        conn.execute(text('''
            CREATE TABLE IF NOT EXISTS anh_xa (
                ten_phu TEXT PRIMARY KEY,
                ten_chuan TEXT
            );
        '''))

init_db()

# Hàm làm sạch tên (viết hoa chữ cái đầu, xóa khoảng trắng thừa)
def clean_name(name_str):
    if not isinstance(name_str, str) or not name_str.strip():
        return "Chưa Rõ"
    return " ".join(name_str.strip().split()).title()

# Hàm định dạng tiền tệ (VD: 100.000)
def format_money(val):
    try:
        val = float(val)
        if val < 0:
            return f"-{abs(val):,.0f}".replace(",", ".")
        return f"{val:,.0f}".replace(",", ".")
    except:
        return "0"

# Hàm chuẩn hóa chuỗi Ngày / Tháng / Năm đồng nhất (VD: 26/07/2026)
def format_date_str(date_str):
    if not date_str or str(date_str) == "nan":
        return ""
    try:
        dt = pd.to_datetime(date_str, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")
    except:
        pass
    return str(date_str).split(" ")[0]

# Lấy dữ liệu từ Supabase Database
def load_data():
    with engine.connect() as conn:
        df_ls = pd.read_sql_query(text("SELECT * FROM lich_su ORDER BY id ASC"), conn)
        df_ax = pd.read_sql_query(text("SELECT * FROM anh_xa"), conn)
    
    if not df_ls.empty:
        df_ls["ten_ncc"] = df_ls["ten_ncc"].astype(str).apply(clean_name)
        df_ls["ten_sp_chuan"] = df_ls["ten_sp_chuan"].astype(str).apply(clean_name)
        df_ls["gia_nhap_le"] = pd.to_numeric(df_ls["gia_nhap_le"], errors='coerce').fillna(0)
        df_ls["quy_cach"] = pd.to_numeric(df_ls["quy_cach"], errors='coerce').fillna(1)
        df_ls["don_gia_thung"] = pd.to_numeric(df_ls["don_gia_thung"], errors='coerce').fillna(0)
        df_ls["so_luong"] = pd.to_numeric(df_ls["so_luong"], errors='coerce').fillna(0)
        
    return df_ls, df_ax

df_lich_su, df_anh_xa = load_data()

# Tạo từ điển bộ nhớ ánh xạ tên
map_anh_xa = {}
if not df_anh_xa.empty:
    for _, r in df_anh_xa.iterrows():
        if pd.notna(r.get("ten_phu")) and pd.notna(r.get("ten_chuan")):
            map_anh_xa[clean_name(str(r["ten_phu"]))] = clean_name(str(r["ten_chuan"]))

# Danh sách tên chuẩn tổng hợp
danh_sach_ten_chuandaco = sorted(list(set(df_lich_su["ten_sp_chuan"].dropna().astype(str).unique()).union(set(map_anh_xa.values())))) if not df_lich_su.empty else sorted(list(set(map_anh_xa.values())))

# TẠO 5 TAB CHỨC NĂNG
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📥 Nhập Hóa Đơn", 
    "🧾 Danh Sách & Tìm HD", 
    "🛠 Quản Lý Tên Chuẩn", 
    "📜 Lịch Sử Chi Tiết",
    "📊 Biểu Đồ Thống Kê"
])

# ---------------------------------------------------------
# TAB 1: NHẬP HÓA ĐƠN
# ---------------------------------------------------------
with tab1:
    st.subheader("📥 Upload File Excel Hóa Đơn")
    uploaded_file = st.file_uploader("Chọn file Excel từ điện thoại", type=["xlsx", "xls"], key="uploader")
    
    if uploaded_file is not None:
        try:
            df_raw = pd.read_excel(uploaded_file)
            ten_ncc_raw = str(df_raw.iloc[1, 1]) if pd.notna(df_raw.iloc[1, 1]) else "Chưa Rõ"
            ten_ncc = clean_name(ten_ncc_raw)
            
            ngay_hd_raw = str(df_raw.iloc[3, 1]) if pd.notna(df_raw.iloc[3, 1]) else ""
            ngay_hd = format_date_str(ngay_hd_raw)
            
            so_hd = str(df_raw.iloc[4, 1]) if pd.notna(df_raw.iloc[4, 1]) else ""
            
            st.success(f"📌 **NCC:** {ten_ncc} | **Số HD:** {so_hd} | **Ngày:** {ngay_hd}")
            
            start_row = -1
            for idx, row in df_raw.iterrows():
                if pd.notna(row.iloc[1]) and "tên hàng" in str(row.iloc[1]).lower():
                    start_row = idx + 1
                    break
            
            if start_row != -1:
                df_items = df_raw.iloc[start_row:].dropna(subset=[df_raw.columns[1]]).copy()
                items_to_save = []
                st.write("---")
                st.subheader("🔍 Khớp Tên & Cảnh Báo Giá")
                
                with st.form("form_khop_ten"):
                    for idx, row in df_items.iterrows():
                        ten_phu = str(row.iloc[1]).strip()
                        if "tổng cộng" in ten_phu.lower() or "số tiền" in ten_phu.lower():
                            continue
                            
                        quy_cach = float(row.iloc[3]) if pd.notna(row.iloc[3]) and float(row.iloc[3]) > 0 else 1.0
                        so_luong = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0.0
                        don_gia_thung = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0.0
                        gia_nhap_le = don_gia_thung / quy_cach
                        
                        ten_chuan_default = map_anh_xa.get(clean_name(ten_phu), clean_name(ten_phu))
                        
                        # CẢNH BÁO GIÁ LÊN / XUỐNG
                        canh_bao_str = ""
                        if not df_lich_su.empty:
                            df_cu = df_lich_su[df_lich_su["ten_sp_chuan"].str.lower() == ten_chuan_default.lower()]
                            if not df_cu.empty:
                                gia_cu_gan_nhat = float(df_cu.iloc[-1]["gia_nhap_le"])
                                chenh_lech = gia_nhap_le - gia_cu_gan_nhat
                                if chenh_lech > 0:
                                    canh_bao_str = f" 🔴 **TĂNG {format_money(chenh_lech)}** *(Đợt trước: {format_money(gia_cu_gan_nhat)})*"
                                elif chenh_lech < 0:
                                    canh_bao_str = f" 🟢 **GIẢM {format_money(abs(chenh_lech))}** *(Đợt trước: {format_money(gia_cu_gan_nhat)})*"
                                else:
                                    canh_bao_str = " ⚪ **Không đổi**"
                        
                        st.markdown(f"**Tên NPP:** `{ten_phu}`")
                        st.markdown(f"Quy cách: {quy_cach:.0f} | Giá thùng: {format_money(don_gia_thung)} ➔ **Giá lẻ: {format_money(gia_nhap_le)}**{canh_bao_str}")
                        
                        options_goi_y = ["-- [Dùng tên mặc định]: " + ten_chuan_default] + danh_sach_ten_chuandaco + ["➕ Nhập tên mới chưa có trong danh sách..."]
                        
                        chon_ten = st.selectbox(
                            f"🔍 Chọn/Tìm tên gợi ý cho món {idx+1}:",
                            options=options_goi_y,
                            key=f"sb_{idx}"
                        )
                        
                        if chon_ten == "➕ Nhập tên mới chưa có trong danh sách...":
                            ten_chuan_user = st.text_input("Gõ tên chuẩn mới:", value="", key=f"inp_{idx}")
                        elif chon_ten.startswith("-- [Dùng tên mặc định]:"):
                            ten_chuan_user = ten_chuan_default
                        else:
                            ten_chuan_user = chon_ten
                            
                        ten_chuan_user = clean_name(ten_chuan_user) if ten_chuan_user.strip() else clean_name(ten_phu)
                        st.write("---")
                        
                        items_to_save.append({
                            "ngay_hd": ngay_hd,
                            "so_hd": so_hd,
                            "ten_ncc": ten_ncc,
                            "ten_phu_npp": ten_phu,
                            "ten_sp_chuan": ten_chuan_user,
                            "quy_cach": quy_cach,
                            "so_luong": so_luong,
                            "don_gia_thung": don_gia_thung,
                            "gia_nhap_le": gia_nhap_le
                        })
                    
                    submitted = st.form_submit_button("💾 LƯU DỮ LIỆU HÓA ĐƠN")
                    if submitted:
                        with engine.begin() as conn:
                            # Lưu lịch sử
                            for item in items_to_save:
                                conn.execute(text('''
                                    INSERT INTO lich_su (ngay_hd, so_hd, ten_ncc, ten_phu_npp, ten_sp_chuan, quy_cach, so_luong, don_gia_thung, gia_nhap_le)
                                    VALUES (:ngay_hd, :so_hd, :ten_ncc, :ten_phu_npp, :ten_sp_chuan, :quy_cach, :so_luong, :don_gia_thung, :gia_nhap_le)
                                '''), item)
                                
                                # Lưu ánh xạ tên
                                conn.execute(text('''
                                    INSERT INTO anh_xa (ten_phu, ten_chuan)
                                    VALUES (:ten_phu, :ten_chuan)
                                    ON CONFLICT (ten_phu) DO UPDATE SET ten_chuan = EXCLUDED.ten_chuan
                                '''), {"ten_phu": clean_name(item["ten_phu_npp"]), "ten_chuan": item["ten_sp_chuan"]})
                        
                        st.balloons()
                        st.success("✅ ĐÃ LƯU THÀNH CÔNG VÀO SUPABASE!")
                        st.rerun()
        except Exception as e:
            st.error(f"Lỗi xử lý file: {e}")

# ---------------------------------------------------------
# TAB 2: DANH SÁCH & TÌM KIẾM HÓA ĐƠN
# ---------------------------------------------------------
with tab2:
    st.subheader("🧾 Danh Sách & Tìm Kiếm Hóa Đơn")
    
    if not df_lich_su.empty:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            ds_ncc = ["Tất cả NPP"] + sorted(list(df_lich_su["ten_ncc"].unique()))
            ncc_selected = st.selectbox("Lọc theo NPP:", ds_ncc)
        with col_f2:
            sap_xep = st.selectbox("Sắp xếp thời gian:", ["Mới nhất trước", "Cũ nhất trước"])
            
        tu_khoa = st.text_input("🔍 Tìm nhanh theo Số HD / Tên món / NPP:", "")
        
        df_hd_grouped = df_lich_su.groupby(["so_hd", "ten_ncc", "ngay_hd"], sort=False).size().reset_index(name="Tong_Mat_Hang")
        
        if sap_xep == "Mới nhất trước":
            df_hd_grouped = df_hd_grouped.iloc[::-1].reset_index(drop=True)
            
        if ncc_selected != "Tất cả NPP":
            df_hd_grouped = df_hd_grouped[df_hd_grouped["ten_ncc"] == ncc_selected]
            
        st.write("---")
        
        count_hd = 0
        for idx, row in df_hd_grouped.iterrows():
            so_hd_cur = row['so_hd']
            ncc_cur = row['ten_ncc']
            ngay_cur = format_date_str(row['ngay_hd'])
            
            mask_hd = (df_lich_su["so_hd"] == so_hd_cur) & (df_lich_su["ten_ncc"] == ncc_cur)
            df_hd_sub = df_lich_su[mask_hd].copy()
            
            if tu_khoa.strip() != "":
                kw = tu_khoa.lower().strip()
                in_so_hd = kw in str(so_hd_cur).lower()
                in_ncc = kw in str(ncc_cur).lower()
                in_sp = df_hd_sub["ten_phu_npp"].astype(str).str.lower().str.contains(kw).any() or df_hd_sub["ten_sp_chuan"].astype(str).str.lower().str.contains(kw).any()
                if not (in_so_hd or in_ncc or in_sp):
                    continue
            
            count_hd += 1
            with st.expander(f"📄 HD: {so_hd_cur} | NCC: {ncc_cur} ({ngay_cur})"):
                st.write(f"**Số mặt hàng:** {row['Tong_Mat_Hang']} món")
                
                list_chuech_lech = []
                for _, r_item in df_hd_sub.iterrows():
                    sp_c = r_item["ten_sp_chuan"]
                    gia_c = r_item["gia_nhap_le"]
                    idx_c = r_item.name
                    
                    df_truoc = df_lich_su[(df_lich_su["ten_sp_chuan"].str.lower() == str(sp_c).lower()) & (df_lich_su.index < idx_c)]
                    if not df_truoc.empty:
                        gia_truoc = float(df_truoc.iloc[-1]["gia_nhap_le"])
                        diff = gia_c - gia_truoc
                    else:
                        diff = 0.0
                    list_chuech_lech.append(diff)
                
                df_hd_sub["Biến Động Giá"] = list_chuech_lech
                
                df_view = pd.DataFrame()
                df_view["Tên Hàng NPP"] = df_hd_sub["ten_phu_npp"]
                df_view["Tên Chuẩn"] = df_hd_sub["ten_sp_chuan"]
                df_view["Quy Cách"] = df_hd_sub["quy_cach"].astype(int)
                df_view["Số Lượng"] = df_hd_sub["so_luong"].astype(int)
                df_view["Giá Thùng"] = df_hd_sub["don_gia_thung"].apply(format_money)
                df_view["Giá Nhập Lẻ"] = df_hd_sub["gia_nhap_le"].apply(format_money)
                df_view["Tăng/Giảm (Tiền)"] = df_hd_sub["Biến Động Giá"].apply(lambda x: f"+{format_money(x)}" if x > 0 else (format_money(x) if x < 0 else "0"))
                
                def style_row(r):
                    val_str = str(r["Tăng/Giảm (Tiền)"])
                    if val_str.startswith("+"):
                        return ['background-color: #ffcccc; color: #990000; font-weight: bold'] * len(r)
                    elif val_str.startswith("-"):
                        return ['background-color: #d4edda; color: #155724; font-weight: bold'] * len(r)
                    return [''] * len(r)
                
                st.dataframe(df_view.style.apply(style_row, axis=1), use_container_width=True)
                
                c_btn1, c_btn2 = st.columns(2)
                
                # SỬA SẢN PHẨM TRONG HÓA ĐƠN
                with c_btn1:
                    with st.popover("✏️ Sửa chi tiết HD"):
                        st.write("Chỉnh sửa Tên chuẩn / Quy cách:")
                        with st.form(f"form_edit_{idx}"):
                            edited_items = []
                            for sub_idx, sub_row in df_hd_sub.iterrows():
                                item_id = int(sub_row["id"])
                                st.caption(f"📌 **{sub_row['ten_phu_npp']}**")
                                new_qc = st.number_input("Quy cách:", value=float(sub_row["quy_cach"]), min_value=1.0, key=f"qc_{sub_idx}")
                                
                                cur_tc = str(sub_row["ten_sp_chuan"])
                                opts = ["-- Dùng tên cũ: " + cur_tc] + danh_sach_ten_chuandaco + ["➕ Gõ tên mới..."]
                                chon_tc = st.selectbox(f"Chọn tên chuẩn gợi ý:", options=opts, key=f"sb_edit_{sub_idx}")
                                
                                if chon_tc == "➕ Gõ tên mới...":
                                    final_tc = st.text_input("Nhập tên mới:", value="", key=f"inp_edit_{sub_idx}")
                                elif chon_tc.startswith("-- Dùng tên cũ:"):
                                    final_tc = cur_tc
                                else:
                                    final_tc = chon_tc
                                    
                                edited_items.append({
                                    "id": item_id, 
                                    "quy_cach": new_qc, 
                                    "ten_sp_chuan": clean_name(final_tc), 
                                    "don_gia_thung": sub_row["don_gia_thung"], 
                                    "ten_phu_npp": clean_name(sub_row["ten_phu_npp"])
                                })
                                st.write("---")
                            
                            btn_save_edit = st.form_submit_button("💾 Cập Nhật")
                            if btn_save_edit:
                                with engine.begin() as conn:
                                    for item in edited_items:
                                        gia_nhap_le_moi = item["don_gia_thung"] / item["quy_cach"]
                                        conn.execute(text('''
                                            UPDATE lich_su 
                                            SET quy_cach = :quy_cach, ten_sp_chuan = :ten_sp_chuan, gia_nhap_le = :gia_nhap_le
                                            WHERE id = :id
                                        '''), {
                                            "quy_cach": item["quy_cach"],
                                            "ten_sp_chuan": item["ten_sp_chuan"],
                                            "gia_nhap_le": gia_nhap_le_moi,
                                            "id": item["id"]
                                        })
                                        
                                        conn.execute(text('''
                                            INSERT INTO anh_xa (ten_phu, ten_chuan)
                                            VALUES (:ten_phu, :ten_chuan)
                                            ON CONFLICT (ten_phu) DO UPDATE SET ten_chuan = EXCLUDED.ten_chuan
                                        '''), {"ten_phu": item["ten_phu_npp"], "ten_chuan": item["ten_sp_chuan"]})
                                
                                st.success("✅ Đã cập nhật xong!")
                                st.rerun()

                # XÓA HÓA ĐƠN
                with c_btn2:
                    with st.popover("🗑 Xóa Hóa Đơn"):
                        st.warning(f"Bạn có chắc muốn XÓA hẳn HD **{so_hd_cur}**?")
                        if st.button("🔴 Xác nhận Xóa", key=f"btn_del_hd_{idx}"):
                            with engine.begin() as conn:
                                conn.execute(text("DELETE FROM lich_su WHERE so_hd = :so_hd AND ten_ncc = :ten_ncc"), {"so_hd": so_hd_cur, "ten_ncc": ncc_cur})
                            st.success(f"🗑 Đã xóa toàn bộ hóa đơn {so_hd_cur}!")
                            st.rerun()

        if count_hd == 0:
            st.warning("Không tìm thấy Hóa đơn nào phù hợp!")
    else:
        st.info("Chưa có hóa đơn nào được lưu.")

# ---------------------------------------------------------
# TAB 3: QUẢN LÝ TÊN CHUẨN
# ---------------------------------------------------------
with tab3:
    st.subheader("🛠 Quản Lý Danh Sách Tên Chuẩn")
    st.caption("Đổi tên sản phẩm hàng loạt hoặc xóa vĩnh viễn tên sai/thừa.")
    
    if danh_sach_ten_chuandaco:
        sp_chon_ql = st.selectbox("🎯 Chọn Tên Chuẩn cần Sửa hoặc Xóa:", danh_sach_ten_chuandaco)
        
        st.write("---")
        col_act1, col_act2 = st.columns(2)
        
        with col_act1:
            st.markdown("##### ✏️ Đổi Tên Chuẩn")
            ten_moi_input = st.text_input("Nhập tên chuẩn mới thay thế:", value=sp_chon_ql)
            if st.button("💾 Đổi Tên Hàng Loạt"):
                ten_moi_clean = clean_name(ten_moi_input)
                if ten_moi_clean and ten_moi_clean != sp_chon_ql:
                    with engine.begin() as conn:
                        conn.execute(text("UPDATE lich_su SET ten_sp_chuan = :ten_moi WHERE ten_sp_chuan = :ten_cu"), {"ten_moi": ten_moi_clean, "ten_cu": sp_chon_ql})
                        conn.execute(text("UPDATE anh_xa SET ten_chuan = :ten_moi WHERE ten_chuan = :ten_cu"), {"ten_moi": ten_moi_clean, "ten_cu": sp_chon_ql})
                    st.success(f"✅ Đã đổi tên '{sp_chon_ql}' ➔ '{ten_moi_clean}'!")
                    st.rerun()
        
        with col_act2:
            st.markdown("##### 🗑 Xóa Tên Chuẩn Vĩnh Viễn")
            st.warning(f"Xóa `{sp_chon_ql}` khỏi gợi ý & đổi các mặt hàng tên này về tên góc NPP.")
            if st.button("❌ Xác Nhận Xóa Tên Này"):
                with engine.begin() as conn:
                    conn.execute(text("DELETE FROM anh_xa WHERE ten_chuan = :ten_cu"), {"ten_cu": sp_chon_ql})
                    conn.execute(text("UPDATE lich_su SET ten_sp_chuan = ten_phu_npp WHERE ten_sp_chuan = :ten_cu"), {"ten_cu": sp_chon_ql})
                st.success(f"🗑 Đã xóa hoàn toàn tên '{sp_chon_ql}'!")
                st.rerun()
    else:
        st.info("Chưa có danh sách Tên chuẩn nào.")

# ---------------------------------------------------------
# TAB 4: LỊCH SỬ CHI TIẾT
# ---------------------------------------------------------
with tab4:
    st.subheader("📜 Toàn bộ lịch sử nhập hàng")
    if not df_lich_su.empty:
        df_all = df_lich_su.copy()
        df_all["ngay_hd"] = df_all["ngay_hd"].apply(format_date_str)
        df_all["don_gia_thung"] = df_all["don_gia_thung"].apply(format_money)
        df_all["gia_nhap_le"] = df_all["gia_nhap_le"].apply(format_money)
        st.dataframe(df_all, use_container_width=True)
    else:
        st.info("Lịch sử trống.")

# ---------------------------------------------------------
# TAB 5: BIỂU ĐỒ & THỐNG KÊ (LOGIC GOM THỜI GIAN THÔNG MINH)
# ---------------------------------------------------------
with tab5:
    st.subheader("📊 Báo Cáo & Biểu Đồ Nhập Hàng")
    
    if not df_lich_su.empty:
        df_stat = df_lich_su.copy()
        df_stat["Tong_Tien_Dong"] = df_stat["so_luong"] * df_stat["don_gia_thung"]
        
        # Chuyển chuỗi Ngày về kiểu Date chuẩn
        df_stat["Date_Obj"] = pd.to_datetime(df_stat["ngay_hd"], dayfirst=True, errors='coerce')
        df_stat = df_stat.dropna(subset=["Date_Obj"])
        
        if not df_stat.empty:
            now = datetime.now()
            
            luat_chon = st.selectbox(
                "📅 Chọn khoảng thời gian báo cáo:",
                ["Tháng này", "Tháng trước", "Năm nay", "Năm trước", "Tùy chỉnh ngày"]
            )
            
            group_mode = "day" # Mặc định nhóm theo Ngày
            
            if luat_chon == "Tháng này":
                start_d = datetime(now.year, now.month, 1)
                end_d = now
                group_mode = "day"
            elif luat_chon == "Tháng trước":
                first_this_month = datetime(now.year, now.month, 1)
                end_d = first_this_month - pd.Timedelta(days=1)
                start_d = datetime(end_d.year, end_d.month, 1)
                group_mode = "day"
            elif luat_chon == "Năm nay":
                start_d = datetime(now.year, 1, 1)
                end_d = now
                group_mode = "month"
            elif luat_chon == "Năm trước":
                start_d = datetime(now.year - 1, 1, 1)
                end_d = datetime(now.year - 1, 12, 31)
                group_mode = "month"
            else: # Tùy chỉnh ngày
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    start_input = st.date_input("Từ ngày:", datetime(now.year, now.month, 1))
                with col_d2:
                    end_input = st.date_input("Đến ngày:", now)
                start_d = datetime.combine(start_input, datetime.min.time())
                end_d = datetime.combine(end_input, datetime.max.time())
                
                days_diff = (end_d - start_d).days
                if days_diff <= 31:
                    group_mode = "day"
                else:
                    group_mode = "month"

            mask_time = (df_stat["Date_Obj"] >= start_d) & (df_stat["Date_Obj"] <= end_d)
            df_filtered = df_stat[mask_time].copy()

            st.write("---")

            if not df_filtered.empty:
                # 1. METRICS TỔNG QUAN
                tong_chi_phi = df_filtered["Tong_Tien_Dong"].sum()
                tong_hd = df_filtered["so_hd"].nunique()
                tong_mon = len(df_filtered)

                col_m1, col_m2, col_m3 = st.columns(3)
                col_m1.metric("💰 Tổng tiền nhập", f"{format_money(tong_chi_phi)} đ")
                col_m2.metric("📄 Tổng số HD", f"{tong_hd} hóa đơn")
                col_m3.metric("📦 Tổng mặt hàng", f"{tong_mon} món")

                st.write("---")

                # 2. BIỂU ĐỒ SỐ TIỀN NHẬP THEO THỜI GIAN
                if group_mode == "day":
                    st.markdown("##### 📈 Biểu Đồ Nhập Hàng Theo Từng Ngày")
                    df_filtered["Time_Key"] = df_filtered["Date_Obj"].dt.strftime("%d/%m/%Y")
                    df_by_time = df_filtered.groupby(["Time_Key", "Date_Obj"])["Tong_Tien_Dong"].sum().reset_index()
                    df_by_time = df_by_time.sort_values(by="Date_Obj")
                    x_label = "Ngày nhập"
                else:
                    st.markdown("##### 📈 Biểu Đồ Nhập Hàng Theo Từng Tháng")
                    df_filtered["Time_Key"] = df_filtered["Date_Obj"].dt.strftime("Tháng %m/%Y")
                    df_filtered["YearMonth"] = df_filtered["Date_Obj"].dt.to_period('M')
                    df_by_time = df_filtered.groupby(["Time_Key", "YearMonth"])["Tong_Tien_Dong"].sum().reset_index()
                    df_by_time = df_by_time.sort_values(by="YearMonth")
                    x_label = "Tháng nhập"
                
                df_by_time.columns = [x_label, "Sort_Key", "Tổng Tiền (Đồng)"]
                
                fig_time = px.bar(
                    df_by_time, 
                    x=x_label, 
                    y="Tổng Tiền (Đồng)",
                    text_auto=True,
                    color_discrete_sequence=['#1f77b4']
                )
                fig_time.update_traces(texttemplate='%{y:,.0f} đ', textposition='outside')
                fig_time.update_layout(xaxis_title=x_label, yaxis_title="Số tiền (VNĐ)", height=420)
                st.plotly_chart(fig_time, use_container_width=True)

                # 3. BIỂU ĐỒ SỐ TIỀN NHẬP THEO NHÀ CUNG CẤP (NPP)
                st.markdown("##### 🏢 Tỷ Trọng Nhập Hàng Theo Nhà Cung Cấp (NPP)")
                df_by_ncc = df_filtered.groupby("ten_ncc")["Tong_Tien_Dong"].sum().reset_index().sort_values(by="Tong_Tien_Dong", ascending=False)
                df_by_ncc.columns = ["Nhà Cung Cấp", "Tổng Tiền (Đồng)"]

                fig_ncc = px.bar(
                    df_by_ncc, 
                    x="Nhà Cung Cấp", 
                    y="Tổng Tiền (Đồng)",
                    color="Nhà Cung Cấp",
                    text_auto=True
                )
                fig_ncc.update_traces(texttemplate='%{y:,.0f} đ', textposition='outside')
                fig_ncc.update_layout(xaxis_title="Nhà Cung Cấp", yaxis_title="Số tiền (VNĐ)", height=420, showlegend=False)
                st.plotly_chart(fig_ncc, use_container_width=True)

            else:
                st.warning("⚠️ Không có dữ liệu hóa đơn nào trong khoảng thời gian này!")
        else:
            st.info("Chưa có ngày hợp lệ trong dữ liệu.")
    else:
        st.info("Chưa có dữ liệu lịch sử để thống kê biểu đồ.")
