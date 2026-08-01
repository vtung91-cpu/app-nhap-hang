import streamlit as st
import pandas as pd
import os
import json

# 1. CẤU HÌNH MÀN HÌNH ĐIỆN THOẠI
st.set_page_config(page_title="App Nhập Hàng", page_icon="📦", layout="centered")

st.title("📦 QUẢN LÝ NHẬP HÀNG")

FILE_LICH_SU = "lich_su_nhap.csv"
FILE_ANH_XA = "anh_xa_ten.json"

# Hàm định dạng tiền tệ: 1.250.000
def format_money(val):
    try:
        val = float(val)
        if val < 0:
            return f"-{abs(val):,.0f}".replace(",", ".")
        return f"{val:,.0f}".replace(",", ".")
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

# Danh sách tên chuẩn đã từng có trong hệ thống
danh_sach_ten_chuandaco = list(set(df_lich_su["Ten_Sp_Chuan"].dropna().unique()).union(set(map_anh_xa.values())))
danh_sach_ten_chuandaco.sort()

# TẠO 3 TAB CHỨC NĂNG
tab1, tab2, tab3 = st.tabs(["📥 Nhập Hóa Đơn", "🧾 Danh Sách & Sửa HD", "📜 Lịch Sử Chi Tiết"])

# ---------------------------------------------------------
# TAB 1: NHẬP HÓA ĐƠN
# ---------------------------------------------------------
with tab1:
    st.subheader(" Upload File Excel Hóa Đơn")
    uploaded_file = st.file_uploader("Chọn file Excel từ điện thoại", type=["xlsx", "xls"], key="uploader")
    
    if uploaded_file is not None:
        try:
            df_raw = pd.read_excel(uploaded_file)
            ten_ncc = str(df_raw.iloc[1, 1]) if pd.notna(df_raw.iloc[1, 1]) else "Chưa rõ"
            ngay_hd = str(df_raw.iloc[3, 1]) if pd.notna(df_raw.iloc[3, 1]) else ""
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
                        
                        ten_chuan_default = map_anh_xa.get(ten_phu, ten_phu)
                        
                        # CẢNH BÁO GIÁ SO VỚI LẦN TRƯỚC GẦN NHẤT
                        canh_bao_str = ""
                        if not df_lich_su.empty:
                            df_cu = df_lich_su[df_lich_su["Ten_Sp_Chuan"] == ten_chuan_default]
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
                        
                        # TẠO DANH SÁCH LỰA CHỌN CÓ TÌM KIẾM GỢI Ý
                        options_list = [ten_chuan_default] + [x for x in danh_sach_ten_chuandaco if x != ten_chuan_default] + ["➕ Nhập tên mới..."]
                        
                        selected_option = st.selectbox(
                            f"Chọn/Gõ tìm tên chuẩn cho món {idx+1}:",
                            options=options_list,
                            key=f"select_{idx}"
                        )
                        
                        if selected_option == "➕ Nhập tên mới...":
                            ten_chuan_user = st.text_input("Gõ tên chuẩn mới:", value="", key=f"input_new_{idx}")
                        else:
                            ten_chuan_user = selected_option
                            
                        st.write("---")
                        
                        items_to_save.append({
                            "Ngay_HD": ngay_hd,
                            "So_HD": so_hd,
                            "Ten_NCC": ten_ncc,
                            "Ten_Phu_NPP": ten_phu,
                            "Ten_Sp_Chuan": ten_chuan_user if ten_chuan_user.strip() != "" else ten_phu,
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
                        st.rerun()
        except Exception as e:
            st.error(f"Lỗi xử lý file: {e}")

# ---------------------------------------------------------
# TAB 2: DANH SÁCH HÓA ĐƠN & SỬA LỖI NHẬP SAI
# ---------------------------------------------------------
with tab2:
    st.subheader("🧾 Danh Sách Hóa Đơn & Chỉnh Sửa")
    if not df_lich_su.empty:
        df_hd_grouped = df_lich_su.groupby(["So_HD", "Ten_NCC", "Ngay_HD"], sort=False).size().reset_index(name="Tong_Mat_Hang")
        
        for idx, row in df_hd_grouped.iterrows():
            so_hd_cur = row['So_HD']
            ncc_cur = row['Ten_NCC']
            ngay_cur = row['Ngay_HD']
            
            with st.expander(f"📄 HD: {so_hd_cur} | NCC: {ncc_cur} ({ngay_cur})"):
                st.write(f"**Số mặt hàng:** {row['Tong_Mat_Hang']} món")
                
                mask_hd = (df_lich_su["So_HD"] == so_hd_cur) & (df_lich_su["Ten_NCC"] == ncc_cur)
                df_hd_sub = df_lich_su[mask_hd].copy()
                
                # TÍNH BIẾN ĐỘNG GIÁ SO VỚI LẦN NHẬP TRƯỚC LIỀN KỀ
                list_chuech_lech = []
                for _, r_item in df_hd_sub.iterrows():
                    sp_c = r_item["Ten_Sp_Chuan"]
                    gia_c = r_item["Gia_Nhap_Le"]
                    idx_c = r_item.name
                    
                    df_truoc = df_lich_su[(df_lich_su["Ten_Sp_Chuan"] == sp_c) & (df_lich_su.index < idx_c)]
                    if not df_truoc.empty:
                        gia_truoc = float(df_truoc.iloc[-1]["Gia_Nhap_Le"])
                        diff = gia_c - gia_truoc
                    else:
                        diff = 0.0
                    list_chuech_lech.append(diff)
                
                df_hd_sub["Biến Động Giá"] = list_chuech_lech
                
                # BẢNG HIỂN THỊ (ĐÃ BỎ CỘT TÊN CHUẨN)
                df_view = pd.DataFrame()
                df_view["Tên Hàng NPP"] = df_hd_sub["Ten_Phu_NPP"]
                df_view["Quy Cách"] = df_hd_sub["Quy_Cach"].astype(int)
                df_view["Số Lượng"] = df_hd_sub["So_Luong"].astype(int)
                df_view["Giá Thùng"] = df_hd_sub["Don_Gia_Thung"].apply(format_money)
                df_view["Giá Nhập Lẻ"] = df_hd_sub["Gia_Nhap_Le"].apply(format_money)
                df_view["Tăng/Giảm (Tiền)"] = df_hd_sub["Biến Động Giá"].apply(lambda x: f"+{format_money(x)}" if x > 0 else (format_money(x) if x < 0 else "0"))
                
                # HÀM TÔ MÀU ĐỎ/XANH
                def style_row(row):
                    val_str = str(row["Tăng/Giảm (Tiền)"])
                    if val_str.startswith("+"):
                        return ['background-color: #ffcccc; color: #990000; font-weight: bold'] * len(row)
                    elif val_str.startswith("-"):
                        return ['background-color: #d4edda; color: #155724; font-weight: bold'] * len(row)
                    return [''] * len(row)
                
                st.dataframe(df_view.style.apply(style_row, axis=1), use_container_width=True)
                
                # KHU VỰC SỬA LỖI SAI
                with st.popover(f"✏️ Sửa Quy Cách / Tên Chuẩn cho HD {so_hd_cur}"):
                    st.write("Chỉnh sửa thông tin:")
                    with st.form(f"form_edit_{idx}"):
                        edited_items = []
                        for sub_idx, sub_row in df_hd_sub.iterrows():
                            st.caption(f"📌 **{sub_row['Ten_Phu_NPP']}**")
                            c1, c2 = st.columns(2)
                            with c1:
                                new_qc = st.number_input("Quy cách:", value=float(sub_row["Quy_Cach"]), min_value=1.0, key=f"qc_{sub_idx}")
                            with c2:
                                new_tc = st.text_input("Tên chuẩn:", value=str(sub_row["Ten_Sp_Chuan"]), key=f"tc_{sub_idx}")
                            
                            edited_items.append((sub_idx, new_qc, new_tc))
                        
                        btn_save_edit = st.form_submit_button("💾 Cập Nhật Thay Đổi")
                        if btn_save_edit:
                            for index_to_update, qc, tc in edited_items:
                                df_lich_su.loc[index_to_update, "Quy_Cach"] = qc
                                df_lich_su.loc[index_to_update, "Ten_Sp_Chuan"] = tc
                                gia_thung_cur = df_lich_su.loc[index_to_update, "Don_Gia_Thung"]
                                df_lich_su.loc[index_to_update, "Gia_Nhap_Le"] = gia_thung_cur / qc
                                
                                ten_phu_c = df_lich_su.loc[index_to_update, "Ten_Phu_NPP"]
                                map_anh_xa[ten_phu_c] = tc
                            
                            df_lich_su.to_csv(FILE_LICH_SU, index=False, encoding="utf-8-sig")
                            with open(FILE_ANH_XA, "w", encoding="utf-8") as f:
                                json.dump(map_anh_xa, f, ensure_ascii=False, indent=4)
                                
                            st.success("✅ Đã cập nhật xong!")
                            st.rerun()
    else:
        st.info("Chưa có hóa đơn nào được lưu.")

# ---------------------------------------------------------
# TAB 3: LỊCH SỬ CHI TIẾT
# ---------------------------------------------------------
with tab3:
    st.subheader("📜 Toàn bộ lịch sử")
    if not df_lich_su.empty:
        df_all = df_lich_su.copy()
        df_all["Don_Gia_Thung"] = df_all["Don_Gia_Thung"].apply(format_money)
        df_all["Gia_Nhap_Le"] = df_all["Gia_Nhap_Le"].apply(format_money)
        st.dataframe(df_all, use_container_width=True)
    else:
        st.info("Lịch sử trống.")
