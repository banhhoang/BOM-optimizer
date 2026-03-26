import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="BOM Checker - Tụ & Trở", layout="wide")
st.title("🚀 Đối chiếu BOM: Nhập giá trực tiếp & So sánh linh kiện")

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
    if pd.isna(value) or str(value).strip() in ["---", "", "NAN", "N/A"]: 
        return float('inf')
    text = str(value).replace('$', '').replace(',', '').strip()
    try: return float(text)
    except: return float('inf')

def parse_watt(watt_str):
    """Chuyển đổi 1/10W, 1/16W thành số thực để so sánh"""
    if not watt_str: return 0.0
    try:
        if '/' in str(watt_str):
            num, den = str(watt_str).split('/')
            return float(num) / float(den)
        return float(watt_str)
    except: return 0.0

def extract_specs(desc):
    desc = str(desc).upper()
    specs = {'volt': 0.0, 'size': '', 'tol': 100.0, 'watt': 0.0}
    
    # Bóc Size
    size_match = re.search(r'(0201|0402|0603|0805|1206|1210|2010|2512)', desc)
    if size_match: specs['size'] = size_match.group(1)
    
    # Bóc Volt
    v_match = re.search(r'(\d+\.?\d*)\s*V', desc)
    if v_match: specs['volt'] = float(v_match.group(1))
    
    # Bóc Sai số (%)
    tol_match = re.search(r'(\d+\.?\d*)\s*%', desc)
    if tol_match: specs['tol'] = float(tol_match.group(1))
    
    # Bóc Công suất (W)
    w_match = re.search(r'(\d+/\d+|\d+\.?\d*)\s*W', desc)
    if w_match: specs['watt'] = parse_watt(w_match.group(1))
    
    return specs

# --- 2. GIAO DIỆN ---

master_file = st.sidebar.file_uploader("1. Nạp Master Data", type=['xlsx', 'xlsm'])
bom_file = st.sidebar.file_uploader("2. Nạp BOM List", type=['xlsx', 'xls'])

if master_file and bom_file:
    dict_master = pd.read_excel(master_file, sheet_name=None)
    for s_name in dict_master:
        dict_master[s_name].columns = [str(c).strip() for c in dict_master[s_name].columns]

    df_bom = pd.read_excel(bom_file)
    df_bom.columns = [str(c).strip() for c in df_bom.columns]

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
        b_qty = st.selectbox("Cột SL cần mua (BOM):", b_cols)
        b_val = st.selectbox("Cột Giá trị đồng bộ (BOM):", b_cols)
        b_desc = st.selectbox("Cột Mô tả (BOM):", b_cols)
        b_type = st.selectbox("Cột Loại hàng hóa (BOM):", b_cols)

    # Nút kiểm tra mã mới (Trường hợp 3)
    if st.button("🔍 1. KIỂM TRA MÃ MỚI", type="secondary"):
        new_codes = []
        for _, row in df_bom.iterrows():
            specs = extract_specs(row[b_desc])
            if specs['size'] in ["0402", "0603"]:
                target_df = dict_master.get(specs['size'], pd.DataFrame())
                exists = target_df[target_df[m_pn].astype(str).str.upper() == str(row[b_pn]).upper()]
                if exists.empty:
                    new_codes.append({
                        "P/N BOM": row[b_pn],
                        "Size": specs['size'],
                        "Giá 1000pcs": 0.0,
                        "Sai số (%)": specs['tol'],
                        "Watt/Volt": specs['watt'] if "RES" in str(row[b_type]).upper() else specs['volt']
                    })
        if new_codes:
            st.session_state.df_new_codes = pd.DataFrame(new_codes).drop_duplicates("P/N BOM")
        else:
            st.session_state.df_new_codes = pd.DataFrame()
            st.info("Không có mã mới nào cần nhập giá.")

    # Bảng nhập giá cho mã mới
    price_map = {}
    if 'df_new_codes' in st.session_state and not st.session_state.df_new_codes.empty:
        st.warning("📋 Nhập giá 1000pcs cho các mã mới để so sánh tối ưu:")
        edited_df = st.data_editor(st.session_state.df_new_codes, use_container_width=True, hide_index=True)
        price_map = edited_df.set_index("P/N BOM").to_dict('index')

    if st.button("🚀 2. BẮT ĐẦU ĐỐI CHIẾU TOÀN BỘ", type="primary"):
        results = []
        for _, row in df_bom.iterrows():
            raw_type_bom = row.get(b_type, "")
            raw_val_bom = row.get(b_val, "")
            qty_bom = row.get(b_qty, 0)
            clean_v_bom = clean_sync_value(raw_val_bom)
            
            if not is_valid_tech_type(raw_type_bom) or clean_v_bom is None:
                continue
            
            pn_bom = str(row[b_pn]).strip().upper() if not pd.isna(row[b_pn]) else "N/A"
            specs_bom = extract_specs(row[b_desc])
            bom_size = specs_bom['size']

            # Giữ nguyên logic Size
            if bom_size not in ["0402", "0603"]:
                results.append({
                    "P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, 
                    "Giá 1000pcs (Gốc)": "---", "Trạng thái": "⏩ GIỮ NGUYÊN", 
                    "Đề xuất": pn_bom, "Giá Đề Xuất": "---", "Lý do": "Chỉ ưu tiên check size 0402 & 0603"
                })
                continue

            target_df = dict_master.get(bom_size, pd.DataFrame())
            if target_df.empty:
                results.append({"P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, "Trạng thái": "❌ THIẾU SHEET", "Đề xuất": "---", "Giá Đề Xuất": "---", "Lý do": f"Không có sheet {bom_size}"})
                continue

            # Lấy thông tin giá gốc và thông số từ Master hoặc từ bảng nhập web
            original_item = target_df[target_df[m_pn].astype(str).str.upper().str.strip() == pn_bom]
            
            if not original_item.empty:
                original_price = original_item.iloc[0][m_price]
                price_to_compare = clean_price(original_price)
            elif pn_bom in price_map:
                original_price = price_map[pn_bom]['Giá 1000pcs']
                price_to_compare = float(original_price)
                # Cập nhật thông số từ web nếu người dùng sửa
                specs_bom['tol'] = price_map[pn_bom]['Sai số (%)']
                if "RES" in str(raw_type_bom).upper(): specs_bom['watt'] = price_map[pn_bom]['Watt/Volt']
                else: specs_bom['volt'] = price_map[pn_bom]['Watt/Volt']
            else:
                original_price = "N/A"
                price_to_compare = float('inf')

            # --- LOGIC 1: CHỮ "CHỌN" (GIỮ NGUYÊN) ---
            mask_select = (target_df[m_pn].astype(str).str.upper().str.strip() == pn_bom) & \
                          (target_df[m_note].astype(str).str.lower().str.contains("chọn", na=False))
            selected_item = target_df[mask_select]

            if not selected_item.empty:
                results.append({
                    "P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, "Giá 1000pcs (Gốc)": original_price,
                    "Trạng thái": "✅ ƯU TIÊN", "Đề xuất": pn_bom, "Giá Đề Xuất": original_price, "Lý do": "Đã duyệt 'Chọn'"
                })
            else:
                # --- LOGIC 2 & 3: TÌM MÃ THAY THẾ RẺ NHẤT (CÓ SO SÁNH TOL/WATT/VOLT) ---
                potential = target_df[
                    (target_df[m_type].apply(is_valid_tech_type)) & 
                    (target_df[m_val].apply(clean_sync_value) == clean_v_bom)
                ]
                
                valid_list = []
                for _, m_row in potential.iterrows():
                    m_specs = extract_specs(m_row[m_desc])
                    is_ok = False
                    if "CAP" in str(raw_type_bom).upper():
                        is_ok = (m_specs['volt'] >= specs_bom['volt'])
                    else:
                        is_ok = (m_specs['tol'] <= specs_bom['tol']) and (m_specs['watt'] >= specs_bom['watt'])
                    
                    if is_ok:
                        item = m_row.to_dict()
                        item['p_num'] = clean_price(m_row[m_price])
                        valid_list.append(item)
                
                if valid_list:
                    best = pd.DataFrame(valid_list).sort_values('p_num').iloc[0]
                    # So sánh với giá gốc (hoặc giá vừa nhập)
                    if best['p_num'] < price_to_compare:
                        results.append({
                            "P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, "Giá 1000pcs (Gốc)": original_price,
                            "Trạng thái": "⚠️ CÓ MÃ THAY THẾ", "Đề xuất": best[m_pn], "Giá Đề Xuất": best[m_price], "Lý do": "Mã Master rẻ hơn & đạt kỹ thuật"
                        })
                    else:
                        results.append({
                            "P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, "Giá 1000pcs (Gốc)": original_price,
                            "Trạng thái": "✅ GIỮ NGUYÊN", "Đề xuất": pn_bom, "Giá Đề Xuất": original_price, "Lý do": "Giá hiện tại là rẻ nhất"
                        })
                else:
                    results.append({
                        "P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, "Giá 1000pcs (Gốc)": original_price,
                        "Trạng thái": "✨ Mã MỚI", "Đề xuất": pn_bom, "Giá Đề Xuất": original_price, "Lý do": "Không có mã thay thế đạt kỹ thuật trong Master"
                    })

        if results:
            df_final = pd.DataFrame(results)
            st.success("Đã đối chiếu xong!")
            st.dataframe(df_final, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False)
            st.download_button("📥 Tải Kết Quả", output.getvalue(), "Ket_qua_BOM_Toi_Uu.xlsx")
