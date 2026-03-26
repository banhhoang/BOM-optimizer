import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="BOM Checker Pro", layout="wide")
st.title("🚀 Đối chiếu BOM: Nhập giá trực tiếp & Tối ưu kỹ thuật")

# --- 1. CÁC HÀM XỬ LÝ (GIỮ NGUYÊN & NÂNG CẤP) ---
def is_valid_tech_type(type_val):
    t = str(type_val).upper()
    return "CAP" in t or "RES" in t

def clean_sync_value(val):
    if pd.isna(val): return None
    s = str(val).strip().upper()
    return s if s not in ["#VALUE!", "#N/A", "NAN", "---", "", "0"] else None

def parse_watt(watt_str):
    if not watt_str: return 0.0
    try:
        if '/' in str(watt_str):
            n, d = str(watt_str).split('/')
            return float(n) / float(d)
        return float(watt_str)
    except: return 0.0

def extract_specs(desc):
    desc = str(desc).upper()
    specs = {'volt': 0.0, 'size': '', 'tol': 100.0, 'watt': 0.0}
    sz = re.search(r'(0201|0402|0603|0805|1206|1210|2010|2512)', desc)
    if sz: specs['size'] = sz.group(1)
    v = re.search(r'(\d+\.?\d*)\s*V', desc)
    if v: specs['volt'] = float(v.group(1))
    t = re.search(r'(\d+\.?\d*)\s*%', desc)
    if t: specs['tol'] = float(t.group(1))
    w = re.search(r'(\d+/\d+|\d+\.?\d*)\s*W', desc)
    if w: specs['watt'] = parse_watt(w.group(1))
    return specs

# --- 2. GIAO DIỆN NẠP FILE ---
with st.sidebar:
    master_file = st.file_uploader("1. Master Data", type=['xlsx', 'xlsm'])
    bom_file = st.file_uploader("2. BOM List", type=['xlsx', 'xls'])

if master_file and bom_file:
    dict_master = pd.read_excel(master_file, sheet_name=None)
    df_bom = pd.read_excel(bom_file)
    
    # Chuẩn hóa tên cột
    for s in dict_master: dict_master[s].columns = [str(c).strip() for c in dict_master[s].columns]
    df_bom.columns = [str(c).strip() for c in df_bom.columns]

    st.subheader("⚙️ Bước 1: Cấu hình cột")
    sample_cols = list(dict_master.values())[0].columns.tolist()
    b_cols = df_bom.columns.tolist()
    
    c1, c2 = st.columns(2)
    with c1:
        m_pn = st.selectbox("P/N (Master)", sample_cols, index=0)
        m_val = st.selectbox("Giá trị đồng bộ (Master)", sample_cols, index=1)
        m_price = st.selectbox("Giá 1000pcs (Master)", sample_cols, index=2)
        m_desc = st.selectbox("Mô tả (Master)", sample_cols, index=3)
        m_type = st.selectbox("Loại hàng (Master)", sample_cols, index=4)
        m_note = st.selectbox("Ghi chú (Master)", sample_cols, index=5)
    with c2:
        b_pn = st.selectbox("P/N (BOM)", b_cols)
        b_val = st.selectbox("Giá trị đồng bộ (BOM)", b_cols)
        b_desc = st.selectbox("Mô tả (BOM)", b_cols)
        b_type = st.selectbox("Loại hàng (BOM)", b_cols)
        b_qty = st.selectbox("SL cần", b_cols)

    st.divider()

    # --- BƯỚC 2: TÌM MÃ MỚI CẦN NHẬP GIÁ ---
    if st.button("🔍 KIỂM TRA MÃ MỚI (TRƯỜNG HỢP 3)"):
        new_items = []
        for _, row in df_bom.iterrows():
            specs = extract_specs(row[b_desc])
            clean_v = clean_sync_value(row[b_val])
            if specs['size'] in ["0402", "0603"] and is_valid_tech_type(row[b_type]) and clean_v:
                target_df = dict_master.get(specs['size'], pd.DataFrame())
                # Kiểm tra xem mã này đã có giá trong Master chưa
                exists = target_df[target_df[m_pn].astype(str).str.upper() == str(row[b_pn]).upper()]
                if exists.empty or pd.isna(exists.iloc[0][m_price]):
                    new_items.append({
                        "P/N BOM": row[b_pn],
                        "Giá trị": clean_v,
                        "Size": specs['size'],
                        "Loại": row[b_type],
                        "Giá 1000pcs": 0.0,
                        "Sai số (%)": specs['tol'],
                        "Công suất/Áp": specs['watt'] if "RES" in str(row[b_type]).upper() else specs['volt']
                    })
        
        if new_items:
            st.session_state.new_codes_df = pd.DataFrame(new_items).drop_duplicates(subset=['P/N BOM'])
        else:
            st.info("Không có mã mới nào cần nhập giá. Bạn có thể tiến hành đối chiếu ngay.")
            st.session_state.new_codes_df = pd.DataFrame()

    # --- BƯỚC 3: BẢNG NHẬP GIÁ TRỰC TIẾP ---
    if 'new_codes_df' in st.session_state and not st.session_state.new_codes_df.empty:
        st.warning("📋 Danh sách mã mới (Trường hợp 3). Vui lòng điền giá và thông số chuẩn để so sánh:")
        edited_new_codes = st.data_editor(
            st.session_state.new_codes_df,
            key="editor_new_codes",
            use_container_width=True,
            hide_index=True
        )
        
        if st.button("🚀 GHI NHẬN & BẮT ĐẦU SO SÁNH TẤT CẢ", type="primary"):
            # Chuyển đổi bảng đã sửa thành từ điển để tra cứu nhanh
            price_map = edited_new_codes.set_index("P/N BOM").to_dict('index')
            
            results = []
            for _, row in df_bom.iterrows():
                p_bom = str(row[b_pn]).upper()
                specs_bom = extract_specs(row[b_desc])
                clean_v_bom = clean_sync_value(row[b_val])
                
                # Cập nhật thông số từ bảng editor nếu có
                if p_bom in price_map:
                    price_bom = price_map[p_bom]['Giá 1000pcs']
                    specs_bom['tol'] = price_map[p_bom]['Sai số (%)']
                    if "RES" in str(row[b_type]).upper(): specs_bom['watt'] = price_map[p_bom]['Công suất/Áp']
                    else: specs_bom['volt'] = price_map[p_bom]['Công suất/Áp']
                else:
                    price_bom = float('inf') # Sẽ được xử lý ở logic Master bên dưới

                # LOGIC SO SÁNH CHÍNH
                if specs_bom['size'] not in ["0402", "0603"] or not is_valid_tech_type(row[b_type]):
                    results.append({"P/N BOM": p_bom, "Trạng thái": "⏩ GIỮ NGUYÊN", "Đề xuất": p_bom, "Giá Đề Xuất": "---", "Lý do": "Size khác"})
                    continue

                target_df = dict_master.get(specs_bom['size'], pd.DataFrame())
                
                # Tìm ứng viên trong Master
                potential = target_df[
                    (target_df[m_type].apply(is_valid_tech_type)) & 
                    (target_df[m_val].apply(clean_sync_value) == clean_v_bom)
                ]
                
                valid_candidates = []
                for _, m_row in potential.iterrows():
                    m_specs = extract_specs(m_row[m_desc])
                    m_price_val = float(str(m_row[m_price]).replace('$','').replace(',','')) if pd.notna(m_row[m_price]) else float('inf')
                    
                    # Điều kiện kỹ thuật
                    if "CAP" in str(row[b_type]).upper():
                        is_ok = (m_specs['volt'] >= specs_bom['volt'])
                    else:
                        is_ok = (m_specs['tol'] <= specs_bom['tol']) and (m_specs['watt'] >= specs_bom['watt'])
                    
                    if is_ok:
                        valid_candidates.append({'pn': m_row[m_pn], 'price': m_price_val})

                # So sánh giá giữa BOM và Master
                if valid_candidates:
                    best_m = sorted(valid_candidates, key=lambda x: x['price'])[0]
                    if best_m['price'] < price_bom:
                        results.append({"P/N BOM": p_bom, "Trạng thái": "⚠️ THAY THẾ RẺ HƠN", "Đề xuất": best_m['pn'], "Giá Đề Xuất": best_m['price'], "Lý do": f"Master rẻ hơn giá bạn nhập ({best_m['price']} < {price_bom})"})
                    else:
                        results.append({"P/N BOM": p_bom, "Trạng thái": "✅ GIỮ NGUYÊN", "Đề xuất": p_bom, "Giá Đề Xuất": price_bom, "Lý do": "Giá bạn nhập là tốt nhất"})
                else:
                    results.append({"P/N BOM": p_bom, "Trạng thái": "✨ MÃ MỚI", "Đề xuất": p_bom, "Giá Đề Xuất": price_bom, "Lý do": "Không có mã Master nào đạt kỹ thuật"})

            st.session_state.final_df = pd.DataFrame(results)

    # --- BƯỚC 4: XUẤT KẾT QUẢ ---
    if 'final_df' in st.session_state:
        st.success("✅ Đã đối chiếu xong dựa trên giá bạn cung cấp!")
        st.dataframe(st.session_state.final_df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            st.session_state.final_df.to_excel(writer, index=False)
        st.download_button("📥 TẢI KẾT QUẢ CUỐI CÙNG", output.getvalue(), "Ket_qua_BOM_Toi_Uu.xlsx")
