import streamlit as st
import pandas as pd
import sqlite3
import plotly.express as px
from datetime import datetime

# 1. CẤU HÌNH MÀN HÌNH
st.set_page_config(page_title="App Nhập Hàng", page_icon="📦", layout="centered")

st.title("📦 QUẢN LÝ NHẬP HÀNG")

# KẾT NỐI VỚI CƠ SỞ DỮ LIỆU SQLITE (Lưu dữ liệu nội bộ vĩnh viễn)
DB_FILE = "nhap_hang.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Bảng Lịch sử nhập hàng
    c.execute('''
        CREATE TABLE IF NOT EXISTS lich_su (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Ngay_HD TEXT,
            So_HD TEXT,
            Ten_NCC TEXT,
            Ten_Phu_NPP TEXT,
            Ten_Sp_Chuan TEXT,
            Quy_Cach REAL,
            So_Luong REAL,
            Don_Gia_Thung REAL,
            Gia_Nhap_Le REAL
        )
    ''')
    # Bảng Ánh xạ tên
    c.execute('''
        CREATE TABLE IF NOT EXISTS anh_xa (
            Ten_Phu TEXT PRIMARY KEY,
            Ten_Chuan TEXT
        )
    ''')
    conn.commit()
    conn.close()

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

# Lấy dữ liệu từ SQLite Database
def load_data():
    conn = sqlite3.connect(DB_FILE)
    df_ls = pd.read_sql_query("SELECT * FROM lich_su", conn)
    df_ax = pd.read_sql_query("SELECT * FROM anh_xa", conn)
    conn.close()
    
    if not df_ls.empty:
        df_ls["Ten_NCC"] = df_ls["Ten_NCC"].astype(str).apply(clean_name)
        df_ls["Ten_Sp_Chuan"] = df_ls["Ten_Sp_Chuan"].astype(str).apply(clean_name)
        df_ls["Gia_Nhap_Le"] = pd.to_numeric(df_ls["Gia_Nhap_Le"], errors='coerce').fillna(0)
        df_ls["Quy_Cach"] = pd.to_numeric(df_ls["Quy_Cach"], errors='coerce').fillna(1)
        df_ls["Don_Gia_Thung"] = pd.to_numeric(df_ls["Don_Gia_Thung"], errors='coerce').fillna(0)
        df_ls["So_Luong"] = pd.to_numeric(df_ls["So_Luong"], errors='coerce').fillna(0)
        
    return df_ls, df_ax

df_lich_su, df_anh_xa = load_data()

# Tạo từ điển bộ nhớ ánh xạ tên
map_anh_xa = {}
if not df_anh_xa.empty:
    for _, r in df_anh_xa.iterrows():
        if pd.notna(r.get("Ten_Phu")) and pd.notna(r.get("Ten_Chuan")):
            map_anh_xa[clean_name(str(r["Ten_Phu"]))] = clean_name(str(r["Ten_Chuan"]))

# Danh sách tên chuẩn tổng hợp
danh_sach_ten_chuandaco = sorted(list(set(df_lich_su["Ten_Sp_Chuan"].dropna().astype(str).unique()).union(set(map_anh_xa.values())))) if not df_lich_su.empty else sorted(list(set(map_anh_xa.values())))

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
    st.subheader(" Upload File Excel Hóa Đơn")
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
                st.subheader(" Khớp Tên & Cảnh Báo Giá")
                
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
                            df_cu = df_lich_su[df_lich_su["Ten_Sp_Chuan"].str.lower() == ten_chuan_default.lower()]
                            if not df_cu.empty:
                                gia_cu_gan_nhat = float(df_cu.iloc[-1]["Gia_Nhap_Le"])
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
                        
                        items_to_save.append((
                            ngay_hd, so_hd, ten_ncc, ten_phu, ten_chuan_user, quy_cach, so_luong, don_gia_thung, gia_nhap_le
                        ))
                    
                    submitted = st.form_submit_button("💾 LƯU DỮ LIỆU HÓA ĐƠN")
                    if submitted:
                        conn = sqlite3.connect(DB_FILE)
                        c = conn.cursor()
                        # Lưu lịch sử
                        c.executemany('''
                            INSERT INTO lich_su (Ngay_HD, So_HD, Ten_NCC, Ten_Phu_NPP, Ten_Sp_Chuan, Quy_Cach, So_Luong, Don_Gia_Thung, Gia_Nhap_Le)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', items_to_save)
                        
                        # Lưu ánh xạ tên
                        for item in items_to_save:
                            c.execute('''
                                INSERT OR REPLACE INTO anh_xa (Ten_Phu, Ten_Chuan)
                                VALUES (?, ?)
                            ''', (clean_name(item[3]), item[4]))
                        
                        conn.commit()
                        conn.close()
                        
                        st.balloons()
                        st.success("✅ ĐÃ LƯU THÀNH CÔNG!")
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
            ds_ncc = ["Tất cả NPP"] + sorted(list(df_lich_su["Ten_NCC"].unique()))
            ncc_selected = st.selectbox("Lọc theo NPP:", ds_ncc)
        with col_f2:
            sap_xep = st.selectbox("Sắp xếp thời gian:", ["Mới nhất trước", "Cũ nhất trước"])
            
        tu_khoa = st.text_input("🔍 Tìm nhanh theo Số HD / Tên món / NPP:", "")
        
        df_hd_grouped = df_lich_su.groupby(["So_HD", "Ten_NCC", "Ngay_HD"], sort=False).size().reset_index(name="Tong_Mat_Hang")
        
        if sap_xep == "Mới nhất trước":
            df_hd_grouped = df_hd_grouped.iloc[::-1].reset_index(drop=True)
            
        if ncc_selected != "Tất cả NPP":
            df_hd_grouped = df_hd_grouped[df_hd_grouped["Ten_NCC"] == ncc_selected]
            
        st.write("---")
        
        count_hd = 0
        for idx, row in df_hd_grouped.iterrows():
            so_hd_cur = row['So_HD']
            ncc_cur = row['Ten_NCC']
            ngay_cur = format_date_str(row['Ngay_HD'])
            
            mask_hd = (df_lich_su["So_HD"] == so_hd_cur) & (df_lich_su["Ten_NCC"] == ncc_cur)
            df_hd_sub = df_lich_su[mask_hd].copy()
            
            if tu_khoa.strip() != "":
                kw = tu_khoa.lower().strip()
                in_so_hd = kw in str(so_hd_cur).lower()
                in_ncc = kw in str(ncc_cur).lower()
                in_sp = df_hd_sub["Ten_Phu_NPP"].astype(str).str.lower().str.contains(kw).any() or df_hd_sub["Ten_Sp_Chuan"].astype(str).str.lower().str.contains(kw).any()
                if not (in_so_hd or in_ncc or in_sp):
                    continue
            
            count_hd += 1
            with st.expander(f"📄 HD: {so_hd_cur} | NCC: {ncc_cur} ({ngay_cur})"):
                st.write(f"**Số mặt hàng:** {row['Tong_Mat_Hang']} món")
                
                list_chuech_lech = []
                for _, r_item in df_hd_sub.iterrows():
                    sp_c = r_item["Ten_Sp_Chuan"]
                    gia_c = r_item["Gia_Nhap_Le"]
                    idx_c = r_item.name
                    
                    df_truoc = df_lich_su[(df_lich_su["Ten_Sp_Chuan"].str.lower() == str(sp_c).lower()) & (df_lich_su.index < idx_c)]
                    if not df_truoc.empty:
                        gia_truoc = float(df_truoc.iloc[-1]["Gia_Nhap_Le"])
                        diff = gia_c - gia_truoc
                    else:
                        diff = 0.0
                    list_chuech_lech.append(diff)
                
                df_hd_sub["Biến Động Giá"] = list_chuech_lech
                
                df_view = pd.DataFrame()
                df_view["Tên Hàng NPP"] = df_hd_sub["Ten_Phu_NPP"]
                df_view["Tên Chuẩn"] = df_hd_sub["Ten_Sp_Chuan"]
                df_view["Quy Cách"] = df_hd_sub["Quy_Cach"].astype(int)
                df_view["Số Lượng"] = df_hd_sub["So_Luong"].astype(int)
                df_view["Giá Thùng"] = df_hd_sub["Don_Gia_Thung"].apply(format_money)
                df_view["Giá Nhập Lẻ"] = df_hd_sub["Gia_Nhap_Le"].apply(format_money)
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
                
                # KHU VỰC SỬA SẢN PHẨM TRONG HÓA ĐƠN
                with c_btn1:
                    with st.popover("✏️ Sửa chi tiết HD"):
                        st.write("Chỉnh sửa Tên chuẩn / Quy cách:")
                        with st.form(f"form_edit_{idx}"):
                            edited_items = []
                            for sub_idx, sub_row in df_hd_sub.iterrows():
                                item_id = sub_row["id"]
                                st.caption(f"📌 **{sub_row['Ten_Phu_NPP']}**")
                                new_qc = st.number_input("Quy cách:", value=float(sub_row["Quy_Cach"]), min_value=1.0, key=f"qc_{sub_idx}")
                                
                                cur_tc = str(sub_row["Ten_Sp_Chuan"])
                                opts = ["-- Dùng tên cũ: " + cur_tc] + danh_sach_ten_chuandaco + ["➕ Gõ tên mới..."]
                                chon_tc = st.selectbox(f"Chọn tên chuẩn gợi ý:", options=opts, key=f"sb_edit_{sub_idx}")
                                
                                if chon_tc == "➕ Gõ tên mới...":
                                    final_tc = st.text_input("Nhập tên mới:", value="", key=f"inp_edit_{sub_idx}")
                                elif chon_tc.startswith("-- Dùng tên cũ:"):
                                    final_tc = cur_tc
                                else:
                                    final_tc = chon_tc
                                    
                                edited_items.append((item_id, new_qc, clean_name(final_tc), sub_row["Don_Gia_Thung"], sub_row["Ten_Phu_NPP"]))
                                st.write("---")
                            
                            btn_save_edit = st.form_submit_button("💾 Cập Nhật")
                            if btn_save_edit:
                                conn = sqlite3.connect(DB_FILE)
                                c = conn.cursor()
                                for row_id, qc, tc, gia_thung, ten_phu_c in edited_items:
                                    gia_nhap_le_moi = gia_thung / qc
                                    c.execute('''
                                        UPDATE lich_su 
                                        SET Quy_Cach = ?, Ten_Sp_Chuan = ?, Gia_Nhap_Le = ?
                                        WHERE id = ?
                                    ''', (qc, tc, gia_nhap_le_moi, row_id))
                                    
                                    c.execute('''
                                        INSERT OR REPLACE INTO anh_xa (Ten_Phu, Ten_Chuan)
                                        VALUES (?, ?)
                                    ''', (clean_name(ten_phu_c), tc))
                                
                                conn.commit()
                                conn.close()
                                st.success("✅ Đã cập nhật xong!")
                                st.rerun()

                # KHU VỰC XÓA HÓA ĐƠN TRÙNG / SAI
                with c_btn2:
                    with st.popover("🗑 Xóa Hóa Đơn"):
                        st.warning(f"Bạn có chắc muốn XÓA hẳn HD **{so_hd_cur}**?")
                        if st.button("🔴 Xác nhận Xóa", key=f"btn_del_hd_{idx}"):
                            conn = sqlite3.connect(DB_FILE)
                            c = conn.cursor()
                            c.execute("DELETE FROM lich_su WHERE So_HD = ? AND Ten_NCC = ?", (so_hd_cur, ncc_cur))
                            conn.commit()
                            conn.close()
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
                    conn = sqlite3.connect(DB_FILE)
                    c = conn.cursor()
                    c.execute("UPDATE lich_su SET Ten_Sp_Chuan = ? WHERE Ten_Sp_Chuan = ?", (ten_moi_clean, sp_chon_ql))
                    c.execute("UPDATE anh_xa SET Ten_Chuan = ? WHERE Ten_Chuan = ?", (ten_moi_clean, sp_chon_ql))
                    conn.commit()
                    conn.close()
                    st.success(f"✅ Đã đổi tên '{sp_chon_ql}' ➔ '{ten_moi_clean}'!")
                    st.rerun()
        
        with col_act2:
            st.markdown("##### 🗑 Xóa Tên Chuẩn Vĩnh Viễn")
            st.warning(f"Xóa `{sp_chon_ql}` khỏi gợi ý & đổi các mặt hàng tên này về tên góc NPP.")
            if st.button("❌ Xác Nhận Xóa Tên Này"):
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("DELETE FROM anh_xa WHERE Ten_Chuan = ?", (sp_chon_ql,))
                c.execute("UPDATE lich_su SET Ten_Sp_Chuan = Ten_Phu_NPP WHERE Ten_Sp_Chuan = ?", (sp_chon_ql,))
                conn.commit()
                conn.close()
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
        df_all["Ngay_HD"] = df_all["Ngay_HD"].apply(format_date_str)
        df_all["Don_Gia_Thung"] = df_all["Don_Gia_Thung"].apply(format_money)
        df_all["Gia_Nhap_Le"] = df_all["Gia_Nhap_Le"].apply(format_money)
        st.dataframe(df_all, use_container_width=True)
    else:
        st.info("Lịch sử trống.")

# ---------------------------------------------------------
# TAB 5: BIỂU ĐỒ & THỐNG KÊ (MỚI THÊM)
# ---------------------------------------------------------
with tab5:
    st.subheader("📊 Báo Cáo & Biểu Đồ Nhập Hàng")
    
    if not df_lich_su.empty:
        # Chuẩn hóa cột ngày tháng để tính toán
        df_stat = df_lich_su.copy()
        df_stat["Tong_Tien_Dong"] = df_stat["So_Luong"] * df_stat["Don_Gia_Thung"]
        
        # Chuyển chuỗi Ngày về kiểu Date chuẩn
        df_stat["Date_Obj"] = pd.to_datetime(df_stat["Ngay_HD"], dayfirst=True, errors='coerce')
        df_stat = df_stat.dropna(subset=["Date_Obj"]) # Bỏ các dòng ngày sai
        
        if not df_stat.empty:
            now = datetime.now()
            
            # CHỌN KHOẢNG THỜI GIAN
            luat_chon = st.selectbox(
                "📅 Chọn khoảng thời gian báo cáo:",
                ["Tháng này", "Tháng trước", "Năm nay", "Năm trước", "Tùy chỉnh ngày"]
            )
            
            # Tính toán Ngày Bắt Đầu và Ngày Kết Thúc
            if luat_chon == "Tháng này":
                start_d = datetime(now.year, now.month, 1)
                end_d = now
            elif luat_chon == "Tháng trước":
                first_this_month = datetime(now.year, now.month, 1)
                end_d = first_this_month - pd.Timedelta(days=1)
                start_d = datetime(end_d.year, end_d.month, 1)
            elif luat_chon == "Năm nay":
                start_d = datetime(now.year, 1, 1)
                end_d = now
            elif luat_chon == "Năm trước":
                start_d = datetime(now.year - 1, 1, 1)
                end_d = datetime(now.year - 1, 12, 31)
            else: # Tùy chỉnh ngày
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    start_input = st.date_input("Từ ngày:", datetime(now.year, now.month, 1))
                with col_d2:
                    end_input = st.date_input("Đến ngày:", now)
                start_d = datetime.combine(start_input, datetime.min.time())
                end_d = datetime.combine(end_input, datetime.max.time())

            # Lọc dữ liệu theo thời gian đã chọn
            mask_time = (df_stat["Date_Obj"] >= start_d) & (df_stat["Date_Obj"] <= end_d)
            df_filtered = df_stat[mask_time].copy()

            st.write("---")

            if not df_filtered.empty:
                # 1. CÁC THÔNG SỐ TỔNG QUAN (METRICS)
                tong_chi_phi = df_filtered["Tong_Tien_Dong"].sum()
                tong_hd = df_filtered["So_HD"].nunique()
                tong_mon = len(df_filtered)

                col_m1, col_m2, col_m3 = st.columns(3)
                col_m1.metric("💰 Tổng tiền nhập", f"{format_money(tong_chi_phi)} đ")
                col_m2.metric("📄 Tổng số HD", f"{tong_hd} hóa đơn")
                col_m3.metric("📦 Tổng mặt hàng", f"{tong_mon} món")

                st.write("---")

                # 2. BIỂU ĐỒ SỐ TIỀN NHẬP THEO THỜI GIAN
                st.markdown("##### 📈 Biểu Đồ Nhập Hàng Theo Thời Gian")
                
                # Gom nhóm theo ngày
                df_by_date = df_filtered.groupby(df_filtered["Date_Obj"].dt.strftime("%d/%m/%Y"))["Tong_Tien_Dong"].sum().reset_index()
                df_by_date.columns = ["Ngày", "Tổng Tiền (Đồng)"]
                
                fig_date = px.bar(
                    df_by_date, 
                    x="Ngày", 
                    y="Tổng Tiền (Đồng)",
                    text_auto=True,
                    color_discrete_sequence=['#1f77b4']
                )
                fig_date.update_traces(texttemplate='%{y:,.0f} đ', textposition='outside')
                fig_date.update_layout(xaxis_title="Ngày nhập", yaxis_title="Số tiền (VNĐ)", height=400)
                st.plotly_chart(fig_date, use_container_width=True)

                # 3. BIỂU ĐỒ SỐ TIỀN NHẬP THEO NHÀ CUNG CẤP (NPP)
                st.markdown("##### 🏢 Tỷ Trọng Nhập Hàng Theo Nhà Cung Cấp (NPP)")
                df_by_ncc = df_filtered.groupby("Ten_NCC")["Tong_Tien_Dong"].sum().reset_index().sort_values(by="Tong_Tien_Dong", ascending=False)
                df_by_ncc.columns = ["Nhà Cung Cấp", "Tổng Tiền (Đồng)"]

                fig_ncc = px.bar(
                    df_by_ncc, 
                    x="Nhà Cung Cấp", 
                    y="Tổng Tiền (Đồng)",
                    color="Nhà Cung Cấp",
                    text_auto=True
                )
                fig_ncc.update_traces(texttemplate='%{y:,.0f} đ', textposition='outside')
                fig_ncc.update_layout(xaxis_title="Nhà Cung Cấp", yaxis_title="Số tiền (VNĐ)", height=400, showlegend=False)
                st.plotly_chart(fig_ncc, use_container_width=True)

            else:
                st.warning("⚠️ Không có dữ liệu hóa đơn nào trong khoảng thời gian này!")
        else:
            st.info("Chưa có ngày hợp lệ trong dữ liệu.")
    else:
        st.info("Chưa có dữ liệu lịch sử để thống kê biểu đồ.")
