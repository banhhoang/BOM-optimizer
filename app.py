

import streamlit as st
import pandas as pd

st.set_page_config(page_title="Tối ưu BOM & Báo giá", layout="wide")
st.title("🛠️ Web App Tối ưu hóa BOM & Báo giá tự động")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    st.subheader("📦 1. Tải lên File Data Gốc (Master)")
    master_file = st.file_uploader("Kéo thả file Master (.xlsx, .xlsm)", type=['xlsx', 'xlsm'], key="master")

with col2:
    st.subheader("📄 2. Tải lên File BOM List")
    bom_file = st.file_uploader("Kéo thả file BOM (.xlsx, .xls)", type=['xlsx', 'xls'], key="bom")

if master_file and bom_file:
    try:
        # BƯỚC ĐỘT PHÁ: Đọc TOÀN BỘ CÁC SHEET trong file Master
        all_master_sheets = pd.read_excel(master_file, sheet_name=None)
        
        # Gom tất cả các sheet lại thành 1 bảng Master duy nhất, thêm cột 'Tên_Sheet_Size' để nhớ Size
        df_master_list = []
        for sheet_name, df in all_master_sheets.items():
            df['Tên_Sheet_Size'] = str(sheet_name).strip().upper() # Lưu tên Sheet (chính là Size: 0402, 0603...)
            df_master_list.append(df)
            
        df_master = pd.concat(df_master_list, ignore_index=True)
        
        # Đọc file BOM (mặc định BOM thường chỉ có 1 sheet)
        df_bom = pd.read_excel(bom_file)
        
        st.success("✅ Đã tải và gộp TẤT CẢ các Sheet thành công!")
        st.markdown("---")
        
        st.subheader("⚙️ Bước 3: Chỉ định các cột dữ liệu")
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Bảng Data Gốc (Master)**")
            m_cols = df_master.columns.astype(str).tolist()
            
            m_pn = st.selectbox("Cột 'P/N của HSX':", m_cols, index=m_cols.index("P/N của HSX (tham khảo)") if "P/N của HSX (tham khảo)" in m_cols else 0)
            m_chon = st.selectbox("Cột đánh dấu 'Chọn' (VD: Ghi chú):", m_cols, index=m_cols.index("Ghi chú") if "Ghi chú" in m_cols else 0)
            m_val = st.selectbox("Cột 'đồng bộ giá trị':", m_cols, index=m_cols.index("đồng bộ giá trị") if "đồng bộ giá trị" in m_cols else 0)
            st.info("💡 Không cần chọn cột Size nữa, hệ thống sẽ tự động dùng tên Sheet (0402, 0603...) làm Size.")
            
        with c2:
            st.markdown("**Bảng BOM List cần mua**")
            b_cols = df_bom.columns.astype(str).tolist()
            
            b_pn = st.selectbox("Cột 'P/N của HSX' (BOM):", b_cols, index=b_cols.index("P/N của HSX (tham khảo)") if "P/N của HSX (tham khảo)" in b_cols else 0)
            b_qty = st.selectbox("Cột 'Số lượng cần mua' (VD: SL cần):", b_cols, index=b_cols.index("SL cần") if "SL cần" in b_cols else 0)
            
        st.markdown("---")
        
        if st.button("🚀 BẮT ĐẦU ĐỐI CHIẾU BOM", type="primary"):
            with st.spinner("Đang quét toàn bộ các Sheet để đối chiếu..."):
                
                # Làm sạch dữ liệu P/N
                df_master[m_pn] = df_master[m_pn].astype(str).str.upper().str.strip()
                df_bom[b_pn] = df_bom[b_pn].astype(str).str.upper().str.strip()

                results = []

                for index, row in df_bom.iterrows():
                    pn_bom = str(row[b_pn]).strip()
                    sl_mua = row[b_qty]
                    
                    if pd.isna(row[b_pn]) or pn_bom == 'NAN': # Bỏ qua dòng trống
                        continue
                        
                    # BƯỚC 1: Tìm P/N của BOM trong file Master (đã gộp tất cả sheet)
                    match_df = df_master[df_master[m_pn] == pn_bom]
                    
                    if match_df.empty:
                        # Không có -> Mã mới
                        results.append({
                            "P/N gốc (BOM)": pn_bom,
                            "Số lượng": sl_mua,
                            "Trạng thái": "❌ MÃ MỚI TINH",
                            "P/N Đề xuất": "---",
                            "Lời khuyên": "Mã mới, cần tìm Supplier và xin giá"
                        })
                    else:
                        # Có trong Master -> Kiểm tra cột Ghi chú
                        matched_row = match_df.iloc[0]
                        is_chon = "chọn" in str(matched_row[m_chon]).lower()
                        
                        if is_chon:
                            # 2.1 CÓ "Chọn"
                            results.append({
                                "P/N gốc (BOM)": pn_bom,
                                "Số lượng": sl_mua,
                                "Trạng thái": "✅ ĐÚNG MÃ ƯU TIÊN",
                                "P/N Đề xuất": pn_bom,
                                "Lời khuyên": "Tuyệt vời, mua mã này!"
                            })
                        else:
                            # 2.2 KHÔNG CÓ "Chọn" -> Đi tìm mã thay thế cùng Giá trị và cùng Sheet(Size)
                            val_goc = str(matched_row[m_val]).strip().upper()
                            sheet_size_goc = str(matched_row['Tên_Sheet_Size']).strip().upper()
                            
                            temp_master = df_master.copy()
                            temp_master[m_val] = temp_master[m_val].astype(str).str.strip().str.upper()
                            
                            # Điều kiện lọc: Cùng Giá Trị + Cùng Sheet (Size) + Có chữ CHỌN
                            alt_df = temp_master[(temp_master[m_val] == val_goc) & 
                                               (temp_master['Tên_Sheet_Size'] == sheet_size_goc) & 
                                               (temp_master[m_chon].astype(str).str.lower().str.contains("chọn", na=False))]
                            
                            if not alt_df.empty:
                                pn_thaythe = alt_df.iloc[0][m_pn]
                                results.append({
                                    "P/N gốc (BOM)": pn_bom,
                                    "Số lượng": sl_mua,
                                    "Trạng thái": "⚠️ CÓ MÃ THAY THẾ TỐT HƠN",
                                    "P/N Đề xuất": pn_thaythe,
                                    "Lời khuyên": f"Đổi sang mã ưu tiên: {pn_thaythe}"
                                })
                            else:
                                results.append({
                                    "P/N gốc (BOM)": pn_bom,
                                    "Số lượng": sl_mua,
                                    "Trạng thái": "⚠️ KHÔNG TÌM THẤY MÃ THAY THẾ",
                                    "P/N Đề xuất": pn_bom,
                                    "Lời khuyên": "Giữ nguyên mã gốc vì không có mã ưu tiên nào cùng thông số"
                                })
                
                df_ketqua = pd.DataFrame(results)
                st.success("🎉 Đã hoàn thành đối chiếu P/N trên toàn bộ kho Data!")
                st.dataframe(df_ketqua, use_container_width=True)

    except Exception as e:
        st.error(f"Lỗi: {e}")
else:
    st.info("👈 Vui lòng tải lên cả 2 file để bắt đầu.")