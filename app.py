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

# Session State Yönetimi (Verilerin kaybolmaması için)
if 's2_manual' not in st.session_state: st.session_state.s2_manual = []
if 'k3_auto' not in st.session_state: st.session_state.k3_auto = []
if 'k3_not_found' not in st.session_state: st.session_state.k3_not_found = []
if 'k3_manual' not in st.session_state: st.session_state.k3_manual = []

st.title("🛡️ Kurumsal Excel ve Fiyatlandırma Paneli")

# ==============================================================================
# 1. BÖLÜM: MUHASEBE ENTEGRE FİYAT HESAPLAMA
# ==============================================================================
st.header("1️⃣ Muhasebe Entegre Fiyat Hesaplama")
f1_col1, f1_col2 = st.columns(2)
with f1_col1: ref_f1 = st.file_uploader("Kendi Listeniz (Excel)", type="xlsx", key="f1_ref")
with f1_col2: svr_f1 = st.file_uploader("SVR Fiyat Listesi (Excel)", type="xlsx", key="f1_svr")

if ref_f1 and svr_f1:
    df_ref_f1, df_svr_f1 = pd.read_excel(ref_f1), pd.read_excel(svr_f1)
    c1, c2, c3, c4 = st.columns(4)
    with c1: ref_code_col = st.selectbox("Kendi Listeniz: Malzeme Kodu", df_ref_f1.columns, key="f1_s1")
    with c2: svr_code_col = st.selectbox("SVR: Malzeme Kodu", df_svr_f1.columns, key="f1_s2")
    with c3: svr_name_col = st.selectbox("SVR: Malzeme Adı", df_svr_f1.columns, key="f1_s3")
    with c4: svr_price_col = st.selectbox("SVR: Birim Fiyatı (Tanımlı)", df_svr_f1.columns, key="f1_s4")
    disc_input = st.text_input("📉 İskonto (Örn: 50+15)", value="50+15", key="f1_disc")
    
    if st.button("🚀 Hesaplamayı Başlat", key="f1_btn"):
        price_map = {clean_code(r[svr_code_col]): (r[svr_price_col], r[svr_name_col]) for _, r in df_svr_f1.iterrows()}
        results_f1 = []
        for _, row in df_ref_f1.iterrows():
            my_c = clean_code(row[ref_code_col])
            if my_c in price_map:
                l_price, u_name = price_map[my_c]
                net = calculate_net_from_list(l_price, disc_input)
                results_f1.append({
                    "Malzeme Kodu": row[ref_code_col], "Ürün Adı": u_name,
                    "Tanımlı Fiyat": round(float(l_price), 2), "Net Fiyat": round(net, 4),
                    "Barem %12": round(net / 0.88, 4), "40+12 Liste": round(net / 0.528, 4)
                })
        if results_f1:
            st.dataframe(pd.DataFrame(results_f1), use_container_width=True)
            out1 = io.BytesIO()
            pd.DataFrame(results_f1).to_excel(out1, index=False)
            st.download_button("📥 Fiyat Excel'ini İndir", out1.getvalue(), "muhasebe_fiyatlari.xlsx")

st.divider()

# ==============================================================================
# 2. BÖLÜM: STOK KARŞILAŞTIRMA (Bozulmadı)
# ==============================================================================
st.header("2️⃣ Stok Karşılaştırma")
# Bu kısım önceki stabil haliyle bırakılmıştır.

st.divider()

# ==============================================================================
# 3. BÖLÜM: KARLILIK ANALİZİ VE MANUEL ASİSTAN (GELİŞTİRİLDİ)
# ==============================================================================
st.header("3️⃣ Karlılık Analizi ve Manuel Eşleştirme")
st.info("Kıyaslama: Karşı Net / Tanımlı Fiyat <= 1.16 ise fiyatı 1.17 ile çarpar.")

col_k1, col_k2 = st.columns(2)
with col_k1: file_k1 = st.file_uploader("1. Excel (Tanımlı Fiyat)", type="xlsx", key="k3_f1")
with col_k2: file_k2 = st.file_uploader("2. Excel (Karşı Net Fiyat)", type="xlsx", key="k3_f2")

if file_k1 and file_k2:
    df_k1, df_k2 = pd.read_excel(file_k1), pd.read_excel(file_k2)
    sk1, sk1_n, sk2, sk3, sk4 = st.columns(5)
    with sk1: k1_c = st.selectbox("E1: Kod", df_k1.columns, key="k3_sel1")
    with sk1_n: k1_n = st.selectbox("E1: İsim", df_k1.columns, key="k3_sel2")
    with sk2: k1_p = st.selectbox("E1: Tanımlı Fiyat", df_k1.columns, key="k3_sel3")
    with sk3: k2_c = st.selectbox("E2: Kod", df_k2.columns, key="k3_sel4")
    with sk4: k2_p = st.selectbox("E2: Karşı Net Fiyat", df_k2.columns, key="k3_sel5")

    if st.button("🚀 Karlılığı Hesapla", key="k3_btn"):
        cost_map = {clean_code(r[k2_c]): r[k2_p] for _, r in df_k2.iterrows() if pd.notna(r[k2_c])}
        
        auto_results, not_found = [], []
        for _, row in df_k1.iterrows():
            c = clean_code(row[k1_c])
            if c in cost_map:
                p1, p2 = float(row[k1_p]), float(cost_map[c])
                ratio = p2 / p1 if p1 > 0 else 0
                new_net = p1 * 1.17 if ratio <= 1.16 else p1
                auto_results.append({
                    "Malzeme Kodu": row[k1_c], "Ürün Adı": row[k1_n],
                    "Eski Tanımlı": round(p1, 2), "Karşı Net": round(p2, 2),
                    "Oran": round(ratio, 4), "YENİ NET": round(new_net, 4),
                    "Barem %12": round(new_net/0.88, 4), "40+12 Liste": round(new_net/0.528, 4),
                    "Tip": "Otomatik (Kod)"
                })
            else:
                not_found.append(row.to_dict())
        
        st.session_state.k3_auto = auto_results
        st.session_state.k3_not_found = not_found
        st.session_state.k3_comp_list = df_k2.to_dict('records')

    # SONUÇ TABLOLARI VE MANUEL ASİSTAN
    if 'k3_auto' in st.session_state:
        tab_k1, tab_k2 = st.tabs(["✅ Analiz Sonuçları", "🔍 Manuel Eşleştirme Asistanı"])
        
        with tab_k1:
            total_k3 = st.session_state.k3_auto + st.session_state.k3_manual
            if total_k3:
                df_res_k3 = pd.DataFrame(total_k3)
                st.dataframe(df_res_k3, use_container_width=True)
                
                out_k3 = io.BytesIO()
                df_res_k3.to_excel(out_k3, index=False)
                st.download_button("📥 Nihai Karlılık Listesini İndir", out_k3.getvalue(), "karlilik_final.xlsx")

        with tab_k2:
            if st.session_state.k3_not_found:
                st.write(f"⚠️ Eşleşmeyen {len(st.session_state.k3_not_found)} ürün var.")
                
                # Manuel Seçim
                sel_item = st.selectbox("Eşleşmeyen Ürün Seç:", st.session_state.k3_not_found, 
                                        format_func=lambda x: f"{x[k1_n]} | {x[k1_c]}")
                
                if sel_item:
                    # Öneriler (İsim benzerliğine göre)
                    search_n = clean_name(sel_item[k1_n])
                    suggs = sorted(st.session_state.k3_comp_list, 
                                  key=lambda x: similarity(search_n, clean_name(x[k2_c] + " " + str(x.get(k2_p, "")))), 
                                  reverse=True)[:5]
                    
                    st.write("💡 **Karşı Dosyadan Öneriler:**")
                    for s in suggs:
                        c_text, c_btn = st.columns([4, 1])
                        c_text.write(f"🔹 {s.get(k2_c, 'Kodsuz')} | Net: {s.get(k2_p, '0')}")
                        
                        if c_btn.button("Bununla Eşleştir", key=f"k3_m_{s[k2_c]}_{sel_item[k1_c]}"):
                            # Hesapla ve ekle
                            p1, p2 = float(sel_item[k1_p]), float(s[k2_p])
                            ratio = p2 / p1 if p1 > 0 else 0
                            new_net = p1 * 1.17 if ratio <= 1.16 else p1
                            
                            match_entry = {
                                "Malzeme Kodu": sel_item[k1_c], "Ürün Adı": sel_item[k1_n],
                                "Eski Tanımlı": round(p1, 2), "Karşı Net": round(p2, 2),
                                "Oran": round(ratio, 4), "YENİ NET": round(new_net, 4),
                                "Barem %12": round(new_net/0.88, 4), "40+12 Liste": round(new_net/0.528, 4),
                                "Tip": "Manuel Eşleşme"
                            }
                            st.session_state.k3_manual.append(match_entry)
                            st.session_state.k3_not_found.remove(sel_item)
                            st.success(f"{sel_item[k1_n]} başarıyla eşleşti!")
                            st.rerun()
            else:
                st.success("Tüm ürünler başarıyla analiz edildi!")
