import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="BOM Checker - Tụ & Trở", layout="wide")
st.title("🚀 Đối chiếu BOM: Ưu tiên mã 'Chọn' & So sánh giá")

# --- 1. CÁC HÀM XỬ LÝ DỮ LIỆU ---

def is_valid_tech_type(type_val):
    t = str(type_val).upper()
    return "CAP" in t or "RES" in t

def clean_sync_value(val):
    if pd.isna(val): return None
    s = str(val).strip().upper()
    invalid_marks = ["#VALUE!", "#N/A", "NAN", "---", "", "NONE", "0", "#REF!"]
    if s in invalid_marks: return None
    return s

def clean_price(value):
    if pd.isna(value) or str(value).strip() in ["---", "", "NAN"]: 
        return float('inf')
    text = str(value).replace('$', '').replace(',', '').strip()
    try: return float(text)
    except: return float('inf')

def extract_specs(desc):
    desc = str(desc).upper()
    specs = {'volt': 0.0, 'size': ''}
    size_match = re.search(r'(0201|0402|0603|0805|1206|1210|2010|2512)', desc)
    if size_match: specs['size'] = size_match.group(1)
    v_match = re.search(r'(\d+\.?\d*)\s*V', desc)
    if v_match: specs['volt'] = float(v_match.group(1))
    return specs

# --- 2. GIAO DIỆN ---

master_file = st.sidebar.file_uploader("1. Nạp Master Data", type=['xlsx', 'xlsm'])
bom_file = st.sidebar.file_uploader("2. Nạp BOM List", type=['xlsx', 'xls'])

if master_file and bom_file:
    dict_master = pd.read_excel(master_file, sheet_name=None)
    for s_name in dict_master:
        # Vừa In Hoa, Vừa xóa khoảng trắng 2 đầu, Vừa thay dấu xuống dòng thành dấu cách
        dict_master[s_name].columns = [str(c).replace('\n', ' ').strip().upper() for c in dict_master[s_name].columns]

    df_bom = pd.read_excel(bom_file)
    # Tương tự cho file BOM
    df_bom.columns = [str(c).replace('\n', ' ').strip().upper() for c in df_bom.columns]

    st.subheader("⚙️ Cấu hình cột")
    sample_cols = list(dict_master.values())[0].columns.tolist()
    
    col1, col2 = st.columns(2)
    with col1:
        st.info("Cấu hình Master")
        m_pn = st.selectbox("Cột P/N (Master):", sample_cols)
        m_val = st.selectbox("Cột Giá trị đồng bộ (Master):", sample_cols)
        m_desc = st.selectbox("Cột Mô tả (Master):", sample_cols)
        m_price = st.selectbox("Cột Giá 1000pcs (Master):", sample_cols)
        m_type = st.selectbox("Cột Loại hàng hóa (Master):", sample_cols)
        m_note = st.selectbox("Cột Ghi chú (Master):", sample_cols)
    with col2:
        st.info("Cấu hình BOM")
        b_cols = df_bom.columns.tolist()
        b_pn = st.selectbox("Cột P/N (BOM):", b_cols)
        b_qty = st.selectbox("Cột SL cần mua (BOM):", b_cols) # Cột mới
        b_val = st.selectbox("Cột Giá trị đồng bộ (BOM):", b_cols)
        b_desc = st.selectbox("Cột Mô tả (BOM):", b_cols)
        b_type = st.selectbox("Cột Loại hàng hóa (BOM):", b_cols)

    if st.button("🚀 BẮT ĐẦU ĐỐI CHIẾU", type="primary"):
        results = []
        
        for _, row in df_bom.iterrows():
            raw_type_bom = row.get(b_type, "")
            raw_val_bom = row.get(b_val, "")
            qty_bom = row.get(b_qty, 0) # Lấy số lượng
            clean_v_bom = clean_sync_value(raw_val_bom)
            
            if not is_valid_tech_type(raw_type_bom) or clean_v_bom is None:
                continue
            
            pn_bom = str(row[b_pn]).strip().upper() if not pd.isna(row[b_pn]) else "N/A"
            specs_bom = extract_specs(row[b_desc])
            bom_size = specs_bom['size']

            target_df = dict_master.get(bom_size, pd.DataFrame())
            
            if target_df.empty:
                results.append({"P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, "Trạng thái": "❌ THIẾU SHEET", "Đề xuất": "---", "Giá Đề Xuất": "---", "Lý do": f"Không có sheet {bom_size}"})
                continue

            # --- TÌM GIÁ GỐC CỦA MÃ BOM TRONG MASTER (ĐỂ SO SÁNH) ---
            original_item = target_df[target_df[m_pn].astype(str).str.upper().str.strip() == pn_bom]
            original_price = original_item.iloc[0][m_price] if not original_item.empty else "N/A"

            # --- LOGIC 1: KIỂM TRA CHỮ "CHỌN" ---
            mask_select = (target_df[m_pn].astype(str).str.upper().str.strip() == pn_bom) & \
                          (target_df[m_note].astype(str).str.lower().str.contains("chọn", na=False))
            selected_item = target_df[mask_select]

            if not selected_item.empty:
                results.append({
                    "P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, "Giá 1000pcs (Gốc)": original_price,
                    "Trạng thái": "✅ ƯU TIÊN", "Đề xuất": pn_bom, "Giá 1000pcs Đề Xuất": selected_item.iloc[0][m_price], "Lý do": "Đã duyệt 'Chọn'"
                })
            else:
                # --- LOGIC 2: TÌM MÃ THAY THẾ RẺ NHẤT ---
                potential = target_df[
                    (target_df[m_type].apply(is_valid_tech_type)) & 
                    (target_df[m_val].apply(clean_sync_value) == clean_v_bom)
                ]
                
                valid_list = []
                for _, m_row in potential.iterrows():
                    m_specs = extract_specs(m_row[m_desc])
                    match = True
                    if "CAP" in str(raw_type_bom).upper():
                        match = (m_specs['volt'] >= specs_bom['volt'])
                    
                    if match:
                        item = m_row.to_dict()
                        item['p_num'] = clean_price(m_row[m_price])
                        valid_list.append(item)
                
                if valid_list:
                    best = pd.DataFrame(valid_list).sort_values('p_num').iloc[0]
                    results.append({
                        "P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, "Giá 1000pcs (Gốc)": original_price,
                        "Trạng thái": "⚠️ CÓ MÃ THAY THẾ", "Đề xuất": best[m_pn], "Giá 1000pcs Đề Xuất": best[m_price], "Lý do": "Mã thay thế rẻ nhất& đạt yêu cầu kỹ thuật"
                    })
                else:
                    results.append({
                        "P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, "Giá 1000pcs (Gốc)": original_price,
                        "Trạng thái": "✨ Mã MỚI", "Đề xuất": "Mã mới tinh", "Giá 1000pcs Đề Xuất": "---", "Lý do": "Chưa có trong Master"
                    })

        if results:
            df_final = pd.DataFrame(results)
            st.success("Đã đối chiếu xong!")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False)
            st.download_button("📥 Tải Kết Quả", output.getvalue(), "Ket_qua_BOM_Full.xlsx")
