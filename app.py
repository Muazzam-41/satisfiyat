import streamlit as st
import pandas as pd
import io
import re
from difflib import SequenceMatcher

# Sayfa Yapılandırması
st.set_page_config(page_title="Excel İşlem Merkezi PRO", layout="wide")

# --- YARDIMCI FONKSİYONLAR ---
def clean_code(text):
    if not text or pd.isna(text): return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).upper()

def clean_name(text):
    if not text or pd.isna(text): return ""
    text = str(text).lower()
    # Sembolleri (●, ▼ vb.) ve gereksiz boşlukları siler
    cleaned = re.sub(r'[^a-z0-9ğüşıöç ]', '', text).strip()
    return cleaned

def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

def calculate_net_from_list(price, disc_str):
    try:
        if isinstance(price, str):
            price = price.replace('₺', '').replace('.', '').replace(',', '.').strip()
        val = float(price)
        if not disc_str or disc_str == "0": return val
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return val
    except: return 0

# Manuel eşleşmeleri saklamak için Session State
if 'manual_matches' not in st.session_state: st.session_state.manual_matches = []

st.title("🛡️ Kurumsal Excel ve Fiyatlandırma Paneli")

# ==============================================================================
# 1. BÖLÜM: MUHASEBE ENTEGRE FİYAT HESAPLAMA (RESTORE EDİLDİ)
# ==============================================================================
st.header("1️⃣ Muhasebe Entegre Fiyat Hesaplama")
st.info("Kendi listenizdeki kodlarla SVR listesini eşleştirip baremleri hesaplar.")

f1_col1, f1_col2 = st.columns(2)
with f1_col1:
    ref_f1 = st.file_uploader("Kendi Listeniz (Excel)", type="xlsx", key="f1_ref")
with f1_col2:
    svr_f1 = st.file_uploader("SVR Fiyat Listesi (Excel)", type="xlsx", key="f1_svr")

if ref_f1 and svr_f1:
    df_ref_f1 = pd.read_excel(ref_f1)
    df_svr_f1 = pd.read_excel(svr_f1)
    
    st.write("⚙️ Sütunları Seçin")
    c1, c2, c3, c4 = st.columns(4)
    with c1: ref_code_col = st.selectbox("Kendi Listeniz: Malzeme Kodu", df_ref_f1.columns)
    with c2: svr_code_col = st.selectbox("SVR: Malzeme Kodu", df_svr_f1.columns)
    with c3: svr_name_col = st.selectbox("SVR: Malzeme Adı", df_svr_f1.columns)
    with c4: svr_price_col = st.selectbox("SVR: Birim Fiyatı (Tanımlı)", df_svr_f1.columns)
    
    disc_input = st.text_input("📉 İskonto (Örn: 50+15)", value="50+15")

    if st.button("🚀 Hesaplamayı Başlat"):
        price_map = {}
        for _, row in df_svr_f1.iterrows():
            c = clean_code(row[svr_code_col])
            if c: price_map[c] = (row[svr_price_col], row[svr_name_col])
        
        results_f1 = []
        for _, row in df_ref_f1.iterrows():
            my_c = clean_code(row[ref_code_col])
            if my_c in price_map:
                l_price, u_name = price_map[my_c]
                net = calculate_net_price = calculate_net_from_list(l_price, disc_input)
                results_f1.append({
                    "Malzeme Kodu": row[ref_code_col], "Ürün Adı": u_name,
                    "Tanımlı Fiyat": round(float(l_price), 2), "Net Fiyat": round(net, 4),
                    "Barem %12": round(net / 0.88, 4), "40+12 Liste": round(net / 0.528, 4)
                })
        
        if results_f1:
            st.dataframe(pd.DataFrame(results_f1), use_container_width=True)
            out1 = io.BytesIO()
            pd.DataFrame(results_f1).to_excel(out1, index=False)
            st.download_button("📥 Fiyat Excel'ini İndir", out1.getvalue(), "muhasebe_fiyat_listesi.xlsx")

st.divider()

# ==============================================================================
# 2. BÖLÜM: STOK KARŞILAŞTIRMA VE MANUEL ASİSTAN (BOZULMADI)
# ==============================================================================
st.header("2️⃣ Stok Karşılaştırma ve Manuel Seçim")
# ... (Bu kısım isteğiniz üzerine bozulmadan korundu ve manuel asistan özellikleri eklendi)
col_s1, col_s2 = st.columns(2)
with col_s1: file_s1 = st.file_uploader("Kendi Stok Excel'iniz", type="xlsx", key="s_ref")
with col_s2: file_s2 = st.file_uploader("Karşılaştırılacak Excel", type="xlsx", key="s_comp")

if file_s1 and file_s2:
    df_s1, df_s2 = pd.read_excel(file_s1), pd.read_excel(file_s2)
    sc1, sc2, sc3, sc4 = st.columns(4)
    with sc1: r_c = st.selectbox("Stok: Kod", df_s1.columns, key="sc1")
    with sc2: r_n = st.selectbox("Stok: İsim", df_s1.columns, key="sc2")
    with sc3: c_c = st.selectbox("Karşı: Kod", df_s2.columns, key="sc3")
    with sc4: c_n = st.selectbox("Karşı: İsim", df_s2.columns, key="sc4")

    if st.button("🚀 Otomatik Eşleştir"):
        comp_by_code = {clean_code(r[c_c]): r for _, r in df_s2.iterrows() if pd.notna(r[c_c])}
        comp_by_name = {clean_name(r[c_n]): r for _, r in df_s2.iterrows() if pd.notna(r[c_n])}
        auto_m, not_f = [], []
        for _, row in df_s1.iterrows():
            code_c, name_c = clean_code(row[r_c]), clean_name(row[r_n])
            if code_c in comp_by_code: auto_m.append({**row.to_dict(), "Tip": "KOD", **comp_by_code[code_c].add_prefix("Karşı_")})
            elif name_c in comp_by_name: auto_m.append({**row.to_dict(), "Tip": "İSİM", **comp_by_name[name_c].add_prefix("Karşı_")})
            else: not_f.append(row.to_dict())
        st.session_state.auto_matched = auto_m
        st.session_state.not_found = not_f
        st.session_state.comp_list = df_s2.to_dict('records')

    if 'not_found' in st.session_state:
        t1, t2 = st.tabs(["✅ Eşleşenler", "🔍 Manuel Asistan"])
        with t1:
            all_m = st.session_state.get('auto_matched', []) + st.session_state.manual_matches
            if all_m:
                st.dataframe(pd.DataFrame(all_m), use_container_width=True)
        with t2:
            if st.session_state.not_found:
                sel = st.selectbox("Ürün Seç:", st.session_state.not_found, format_func=lambda x: f"{x[r_n]}")
                sugg = sorted(st.session_state.comp_list, key=lambda x: similarity(clean_name(sel[r_n]), clean_name(x[c_n])), reverse=True)[:3]
                for s in sugg:
                    if st.button(f"Eşleştir: {s[c_n]}", key=f"m_{s[c_c]}"):
                        st.session_state.manual_matches.append({**sel, "Tip": "MANUEL", **pd.Series(s).add_prefix("Karşı_")})
                        st.session_state.not_found.remove(sel)
                        st.rerun()

st.divider()

# ==============================================================================
# 3. BÖLÜM: KARLILIK ANALİZİ VE YENİDEN FİYATLANDIRMA (YENİ)
# ==============================================================================
st.header("3️⃣ Karlılık Analizi ve Akıllı Fiyatlandırma")
st.info("Kıyaslama yapar: Karşı Net / Tanımlı Fiyat <= 1.16 ise fiyatı 1.17 ile çarpar.")

col_k1, col_k2 = st.columns(2)
with col_k1:
    file_k1 = st.file_uploader("1. Excel (Tanımlı Fiyat Olan)", type="xlsx", key="k_ref")
with col_k2:
    file_k2 = st.file_uploader("2. Excel (Karşı Net Fiyat Olan)", type="xlsx", key="k_comp")

if file_k1 and file_k2:
    df_k1, df_k2 = pd.read_excel(file_k1), pd.read_excel(file_k2)
    sk1, sk2, sk3, sk4 = st.columns(4)
    with sk1: k1_code = st.selectbox("Excel 1: Malzeme Kodu", df_k1.columns, key="k1_c")
    with sk2: k1_price = st.selectbox("Excel 1: Birim Fiyatı (Tanımlı)", df_k1.columns, key="k1_p")
    with sk3: k2_code = st.selectbox("Excel 2: Malzeme Kodu", df_k2.columns, key="k2_c")
    with sk4: k2_price = st.selectbox("Excel 2: Karşı_Net Fiyat", df_k2.columns, key="k2_p")

    if st.button("🚀 Karlılığı Hesapla"):
        # Excel 2'yi haritalandır (Kod -> Karşı_Net Fiyat)
        cost_map = {clean_code(r[k2_code]): r[k2_price] for _, r in df_k2.iterrows() if pd.notna(r[k2_code])}
        
        k_results = []
        for _, row in df_k1.iterrows():
            c = clean_code(row[k1_code])
            if c in cost_map:
                p1_tanimli = float(row[k1_price])
                p2_karsi_net = float(cost_map[c])
                
                # Oran Hesabı
                ratio = p2_karsi_net / p1_tanimli if p1_tanimli > 0 else 0
                
                # Kural Uygulama
                if ratio <= 1.16:
                    new_net = p1_tanimli * 1.17
                    durum = "Yeniden Fiyatlandı (x1.17)"
                else:
                    new_net = p1_tanimli
                    durum = "Normal"
                
                k_results.append({
                    "Malzeme Kodu": row[k1_code],
                    "Tanımlı Fiyat (Eski)": round(p1_tanimli, 2),
                    "Karşı Net Fiyat": round(p2_karsi_net, 2),
                    "Oran": round(ratio, 4),
                    "YENİ NET FİYAT": round(new_net, 4),
                    "Barem %12": round(new_net / 0.88, 4),
                    "40+12 Liste": round(new_net / 0.528, 4),
                    "Analiz Notu": durum
                })

        if k_results:
            st.success("Analiz Tamamlandı!")
            df_k_final = pd.DataFrame(k_results)
            st.dataframe(df_k_final, use_container_width=True)
            
            out_k = io.BytesIO()
            df_k_final.to_excel(out_k, index=False)
            st.download_button("📥 Karlılık Analizini İndir", out_k.getvalue(), "karlilik_analizi.xlsx")
