import streamlit as st
import pandas as pd
import plotly.express as px
import os
import json

# 1. CẤU HÌNH MÀN HÌNH ĐIỆN THOẠI
st.set_page_config(page_title="App Nhập Hàng", page_icon="📦", layout="centered")

st.title("📦 QUẢN LÝ NHẬP HÀNG")

FILE_LICH_SU = "lich_su_nhap.csv"
FILE_ANH_XA = "anh_xa_ten.json"

# Hàm định dạng tiền tệ chuẩn: 1.250.000 (không dính chữ đ/VND)
def format_money(val):
    try:
        return f"{float(val):,.0f}".replace(",", ".")
    except:
        return "0"

# Nạp bộ nhớ ánh xạ tên
if os.path.exists(FILE_ANH_XA):
    with open(FILE_ANH_XA, "r", encoding="utf-8") as f:
        map_anh_xa = json.load(f)
else:
    map_anh_xa = {}

# Nạp dữ liệu lịch sử
if os.path.exists(FILE_LICH_SU):
    df_lich_su = pd.read_csv(FILE_LICH_SU, encoding="utf-8-sig")
else:
    df_lich_su = pd.DataFrame(columns=[
        "Ngay_HD", "So_HD", "Ten_NCC", "Ten_Phu_NPP", "Ten_Sp_Chuan", 
        "Quy_Cach", "So_Luong", "Don_Gia_Thung", "Gia_Nhap_Le"
    ])

# 4 TAB CHỨC NĂNG
tab1, tab2, tab3, tab4 = st.tabs(["📥 Nhập Hóa Đơn", "🧾 Danh Sách HD", "📊 Biểu Đồ Giá", "📜 Lịch Sử Chi Tiết"])

# ---------------------------------------------------------
# TAB 1: NHẬP HÓA ĐƠN
# ---------------------------------------------------------
with tab1:
    st.subheader(" Upload File Excel Hóa Đơn")
    uploaded_file = st.file_uploader("Chọn file Excel từ điện thoại", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            df_raw = pd.read_excel(uploaded_file)
            ten_ncc = str(df_raw.iloc[1, 1]) if pd.notna(df_raw.iloc[1, 1]) else "Chưa rõ"
            ngay_hd = str(df_raw.iloc[3, 1]) if pd.notna(df_raw.iloc[3, 1]) else ""
            so_hd = str(df_raw.iloc[4, 1]) if pd.notna(df_raw.iloc[4, 1]) else ""
            
            st.success(f"📌 **NCC:** {ten_ncc} | **Số HD:** {so_hd}")
            
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
                        
                        ten_chuan_default = map_anh_xa.get(ten_phu, ten_phu)
                        
                        # TÍNH CẢNH BÁO GIÁ TĂNG / GIẢM SO VỚI LẦN TRƯỚC
                        canh_bao_str = ""
                        if not df_lich_su.empty:
                            df_cu = df_lich_su[df_lich_su["Ten_Sp_Chuan"] == ten_chuan_default]
                            if not df_cu.empty:
                                gia_cu_gan_nhat = float(df_cu.iloc[-1]["Gia_Nhap_Le"])
                                chenh_lech = gia_nhap_le - gia_cu_gan_nhat
                                if chenh_lech > 0:
                                    canh_bao_str = f" 🔴 **TĂNG {format_money(chenh_lech)}** *(Cũ: {format_money(gia_cu_gan_nhat)})*"
                                elif chenh_lech < 0:
                                    canh_bao_str = f" 🟢 **GIẢM {format_money(abs(chenh_lech))}** *(Cũ: {format_money(gia_cu_gan_nhat)})*"
                                else:
                                    canh_bao_str = " ⚪ **Giá không đổi**"
                        
                        st.markdown(f"**Tên phụ:** `{ten_phu}`")
                        st.markdown(f"Quy cách: {quy_cach:.0f} | Giá thùng: {format_money(don_gia_thung)} ➔ **Giá lẻ: {format_money(gia_nhap_le)}**{canh_bao_str}")
                        
                        ten_chuan_user = st.text_input(
                            label="Tên chính:",
                            value=ten_chuan_default,
                            key=f"item_{idx}"
                        )
                        st.write("---")
                        
                        items_to_save.append({
                            "Ngay_HD": ngay_hd,
                            "So_HD": so_hd,
                            "Ten_NCC": ten_ncc,
                            "Ten_Phu_NPP": ten_phu,
                            "Ten_Sp_Chuan": ten_chuan_user,
                            "Quy_Cach": quy_cach,
                            "So_Luong": so_luong,
                            "Don_Gia_Thung": don_gia_thung,
                            "Gia_Nhap_Le": gia_nhap_le
                        })
                    
                    submitted = st.form_submit_button("💾 LƯU HÓA ĐƠN VÀO HỆ THỐNG")
                    if submitted:
                        for item in items_to_save:
                            map_anh_xa[item["Ten_Phu_NPP"]] = item["Ten_Sp_Chuan"]
                        
                        with open(FILE_ANH_XA, "w", encoding="utf-8") as f:
                            json.dump(map_anh_xa, f, ensure_ascii=False, indent=4)
                        
                        df_new = pd.DataFrame(items_to_save)
                        df_updated = pd.concat([df_lich_su, df_new], ignore_index=True)
                        df_updated.to_csv(FILE_LICH_SU, index=False, encoding="utf-8-sig")
                        
                        st.balloons()
                        st.success("✅ ĐÃ LƯU HÓA ĐƠN THÀNH CÔNG!")
        except Exception as e:
            st.error(f"Lỗi xử lý file: {e}")

# ---------------------------------------------------------
# TAB 2: DANH SÁCH HÓA ĐƠN ĐÃ NHẬP (BẤM XEM CHI TIẾT)
# ---------------------------------------------------------
with tab2:
    st.subheader("🧾 Danh Sách Hóa Đơn Đã Nhập")
    if not df_lich_su.empty:
        # Nhóm danh sách Hóa đơn
        df_hd_list = df_lich_su.groupby(["So_HD", "Ten_NCC", "Ngay_HD"]).size().reset_index(name="Tong_Mat_Hang")
        
        for idx, row in df_hd_list.iterrows():
            with st.expander(f"📄 HD: {row['So_HD']} | NCC: {row['Ten_NCC']} ({row['Ngay_HD']})"):
                st.write(f"**Tổng số mặt hàng:** {row['Tong_Mat_Hang']} món")
                st.write("---")
                
                # Chi tiết từng mặt hàng thuộc Hóa đơn này
                df_ct = df_lich_su[(df_lich_su["So_HD"] == row['So_HD']) & (df_lich_su["Ten_NCC"] == row['Ten_NCC'])].copy()
                df_ct["Giá Thùng"] = df_ct["Don_Gia_Thung"].apply(format_money)
                df_ct["Giá Nhập Lẻ"] = df_ct["Gia_Nhap_Le"].apply(format_money)
                
                st.dataframe(
                    df_ct[["Ten_Sp_Chuan", "Ten_Phu_NPP", "Quy_Cach", "So_Luong", "Giá Thùng", "Giá Nhập Lẻ"]],
                    use_container_width=True
                )
    else:
        st.info("Chưa có hóa đơn nào được lưu.")

# ---------------------------------------------------------
# TAB 3: BIỂU ĐỒ GIÁ
# ---------------------------------------------------------
with tab3:
    st.subheader("📈 Biểu Đồ Biến Động Giá Nhập Lẻ")
    if not df_lich_su.empty:
        danh_sach_sp = df_lich_su["Ten_Sp_Chuan"].unique()
        sp_chon = st.selectbox("🎯 Chọn sản phẩm cần xem giá:", danh_sach_sp)
        
        df_filtered = df_lich_su[df_lich_su["Ten_Sp_Chuan"] == sp_chon].copy()
        
        if not df_filtered.empty:
            fig = px.line(
                df_filtered, 
                x="Ngay_HD", 
                y="Gia_Nhap_Le", 
                markers=True,
                title=f"Lịch sử giá lẻ: {sp_chon}",
                labels={"Ngay_HD": "Ngày Nhập", "Gia_Nhap_Le": "Giá Nhập Lẻ"},
                hover_data=["Ten_NCC", "Don_Gia_Thung", "Quy_Cach"]
            )
            fig.update_traces(line_color="#1f77b4", line_width=3, marker_size=8)
            st.plotly_chart(fig, use_container_width=True)
            
            df_display = df_filtered[["Ngay_HD", "Ten_NCC", "Don_Gia_Thung", "Quy_Cach", "Gia_Nhap_Le"]].copy()
            df_display["Don_Gia_Thung"] = df_display["Don_Gia_Thung"].apply(format_money)
            df_display["Gia_Nhap_Le"] = df_display["Gia_Nhap_Le"].apply(format_money)
            st.dataframe(df_display, use_container_width=True)
    else:
        st.info("Chưa có dữ liệu lịch sử.")

# ---------------------------------------------------------
# TAB 4: LỊCH SỬ CHI TIẾT
# ---------------------------------------------------------
with tab4:
    st.subheader("📜 Toàn bộ lịch sử")
    if not df_lich_su.empty:
        df_all = df_lich_su.copy()
        df_all["Don_Gia_Thung"] = df_all["Don_Gia_Thung"].apply(format_money)
        df_all["Gia_Nhap_Le"] = df_all["Gia_Nhap_Le"].apply(format_money)
        st.dataframe(df_all, use_container_width=True)
    else:
        st.info("Lịch sử trống.")
