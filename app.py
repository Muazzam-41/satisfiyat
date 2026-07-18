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

# Hafıza Alanları (Session State)
state_keys = [
    's2_auto', 's2_not_found', 's2_manual', 's2_comp_list',
    'k3_auto', 'k3_not_found', 'k3_unused_e2', 'k3_manual', 'k3_comp_list'
]
for key in state_keys:
    if key not in st.session_state: st.session_state[key] = []

st.title("🛡️ Kurumsal Excel ve Fiyatlandırma Paneli")

# ==============================================================================
# 1. BÖLÜM: MUHASEBE ENTEGRE FİYAT HESAPLAMA
# ==============================================================================
st.header("1️⃣ Muhasebe Entegre Fiyat Hesaplama")
f1_c1, f1_c2 = st.columns(2)
with f1_c1: f1_ref = st.file_uploader("Kendi Listeniz (Excel)", type="xlsx", key="f1_r")
with f1_c2: f1_svr = st.file_uploader("SVR Fiyat Listesi (Excel)", type="xlsx", key="f1_s")

if f1_ref and f1_svr:
    df1_r, df1_s = pd.read_excel(f1_ref), pd.read_excel(f1_svr)
    c1_1, c1_2, c1_3, c1_4 = st.columns(4)
    with c1_1: f1_rc = st.selectbox("Kendi Kod", df1_r.columns, key="f1_s1")
    with c1_2: f1_sc = st.selectbox("SVR Kod", df1_s.columns, key="f1_s2")
    with c1_3: f1_sn = st.selectbox("SVR İsim", df1_s.columns, key="f1_s3")
    with c1_4: f1_sp = st.selectbox("SVR Birim Fiyatı (Tanımlı)", df1_s.columns, key="f1_s4")
    f1_disc = st.text_input("📉 İskonto (Örn: 50+15)", value="50+15", key="f1_d")
    
    if st.button("🚀 Fiyatları Hesapla", key="f1_btn"):
        s_map = {clean_code(r[f1_sc]): (r[f1_sp], r[f1_sn]) for _, r in df1_s.iterrows()}
        f1_res = []
        for _, row in df1_r.iterrows():
            code = clean_code(row[f1_rc])
            if code in s_map:
                p, n = s_map[code]
                net = calculate_net_from_list(p, f1_disc)
                f1_res.append({
                    "Kod": row[f1_rc], "Ürün Adı": n, "Liste Fiyatı": round(float(p), 2),
                    "Net Fiyat": round(net, 4), "Barem %12": round(net/0.88, 4), "40+12 Liste": round(net/0.528, 4)
                })
        if f1_res:
            st.dataframe(pd.DataFrame(f1_res), use_container_width=True)
            out1 = io.BytesIO(); pd.DataFrame(f1_res).to_excel(out1, index=False)
            st.download_button("📥 Muhasebe Excel İndir", out1.getvalue(), "muhasebe_fiyatlari.xlsx")

st.divider()

# ==============================================================================
# 2. BÖLÜM: STOK KARŞILAŞTIRMA (Hibrit & Fiyat Bilgili)
# ==============================================================================
st.header("2️⃣ Stok Karşılaştırma (Fiyat Bilgili Hibrit)")
f2_c1, f2_c2 = st.columns(2)
with f2_c1: f2_ref = st.file_uploader("Kendi Reel Stok Excel'iniz", type="xlsx", key="f2_r")
with f2_c2: f2_comp = st.file_uploader("Karşılaştırılacak Excel", type="xlsx", key="f2_s")

if f2_ref and f2_comp:
    df2_r, df2_s = pd.read_excel(f2_ref), pd.read_excel(f2_comp)
    st.write("⚙️ **Sütun Ayarları**")
    s2_1, s2_2, s2_3, s2_4 = st.columns(4)
    with s2_1: r2_c = st.selectbox("Reel: Kod", df2_r.columns, key="f2_sc1")
    with s2_2: r2_n = st.selectbox("Reel: İsim", df2_r.columns, key="f2_sc2")
    with s2_3: c2_c = st.selectbox("Karşı: Kod", df2_s.columns, key="f2_sc3")
    with s2_4: c2_n = st.selectbox("Karşı: İsim", df2_s.columns, key="f2_sc4")
    
    st.write("💰 **Karşı Dosya Fiyat Sütunları**")
    p2_1, p2_2 = st.columns(2)
    with p2_1: c2_lp = st.selectbox("Karşı: Liste Fiyatı", df2_s.columns, key="f2_lp")
    with p2_2: c2_np = st.selectbox("Karşı: Net Fiyat", df2_s.columns, key="f2_np")

    if st.button("🚀 Stokları Karşılaştır", key="f2_btn"):
        c_by_code = {clean_code(r[c2_c]): r for _, r in df2_s.iterrows() if pd.notna(r[c2_c])}
        c_by_name = {clean_name(r[c2_n]): r for _, r in df2_s.iterrows() if pd.notna(r[c2_n])}
        a_m, n_f = [], []
        for _, row in df2_r.iterrows():
            cc, cn = clean_code(row[r2_c]), clean_name(row[r2_n])
            match_row = c_by_code.get(cc) or c_by_name.get(cn)
            if match_row is not None:
                a_m.append({
                    "Reel Kod": row[r2_c], "Reel İsim": row[r2_n],
                    "Karşı Liste": match_row[c2_lp], "Karşı Net": match_row[c2_np],
                    **match_row.add_prefix("K_")
                })
            else: n_f.append(row.to_dict())
        st.session_state.s2_auto, st.session_state.s2_not_found = a_m, n_f
        st.session_state.s2_comp_list = df2_s.to_dict('records')

    if st.session_state.s2_auto or st.session_state.s2_not_found:
        t2_1, t2_2, t2_3 = st.tabs(["✅ Eşleşenler", "🔍 Manuel Stok Asistanı", "❌ Bulunamayanlar"])
        with t2_1:
            full_s2 = st.session_state.s2_auto + st.session_state.s2_manual
            if full_s2: 
                df_s2 = pd.DataFrame(full_s2); st.dataframe(df_s2, use_container_width=True)
                out2 = io.BytesIO(); df_s2.to_excel(out2, index=False)
                st.download_button("📥 Eşleşmiş Stokları İndir", out2.getvalue(), "stok_eslesme.xlsx")
        with t2_2:
            if st.session_state.s2_not_found:
                sel = st.selectbox("Ürün Seç:", st.session_state.s2_not_found, format_func=lambda x: f"{x[r2_n]}", key="s2_sb")
                suggs = sorted(st.session_state.s2_comp_list, key=lambda x: similarity(clean_name(sel[r2_n]), clean_name(x[c2_n])), reverse=True)[:3]
                for s in suggs:
                    if st.button(f"Eşleştir: {s[c2_n]}", key=f"s2_m_{s[c2_c]}_{sel[r2_c]}"):
                        st.session_state.s2_manual.append({"Reel Kod": sel[r2_c], "Reel İsim": sel[r2_n], "Karşı Liste": s[c2_lp], "Karşı Net": s[c2_np], **pd.Series(s).add_prefix("K_")})
                        st.session_state.s2_not_found.remove(sel); st.rerun()
        with t2_3:
            if st.session_state.s2_not_found:
                df_fail2 = pd.DataFrame(st.session_state.s2_not_found); st.dataframe(df_fail2, use_container_width=True)
                out2f = io.BytesIO(); df_fail2.to_excel(out2f, index=False)
                st.download_button("📥 Eksikleri İndir", out2f.getvalue(), "eksik_stoklar.xlsx")

st.divider()

# ==============================================================================
# 3. BÖLÜM: KARLILIK ANALİZİ (1.16 VE 1.17 KURALI)
# ==============================================================================
st.header("3️⃣ Karlılık Analizi (Akıllı Fiyatlandırma)")
st.info("Kural: Karşı Net / Tanımlı <= 1.16 ise, Tanımlı Fiyatı 1.17 ile çarparak YENİ NET oluşturur.")

f3_c1, f3_c2 = st.columns(2)
with f3_c1: f3_ref = st.file_uploader("1. Excel (Tanımlı Fiyatlar)", type="xlsx", key="f3_r")
with f3_c2: f3_comp = st.file_uploader("2. Excel (Karşı Net Fiyatlar)", type="xlsx", key="f3_s")

if f3_ref and f3_comp:
    df3_r, df3_s = pd.read_excel(f3_ref), pd.read_excel(f3_comp)
    c3_sel = st.columns(5)
    with c3_sel[0]: k3_c1 = st.selectbox("E1: Kod", df3_r.columns, key="k3_s1")
    with c3_sel[1]: k3_n1 = st.selectbox("E1: İsim", df3_r.columns, key="k3_s2")
    with c3_sel[2]: k3_p1 = st.selectbox("E1: Tanımlı Fiyat", df3_r.columns, key="k3_s3")
    with c3_sel[3]: k3_c2 = st.selectbox("E2: Kod", df3_s.columns, key="k3_s4")
    with c3_sel[4]: k3_p2 = st.selectbox("E2: Karşı Net Fiyat", df3_s.columns, key="k3_s5")

    if st.button("🚀 Karlılığı Analiz Et", key="f3_btn"):
        # Excel 2 haritası
        cost_map = {clean_code(r[k3_c2]): r[k3_p2] for _, r in df3_s.iterrows() if pd.notna(r[k3_c2])}
        matched_e2 = set()
        a_r, n_f = [], []
        
        for _, row in df3_r.iterrows():
            code = clean_code(row[k3_c1])
            if code in cost_map:
                p1, p2 = float(row[k3_p1]), float(cost_map[code])
                ratio = p2 / p1 if p1 > 0 else 0
                
                # --- KRİTİK KURAL UYGULAMASI ---
                if ratio <= 1.16:
                    new_net = p1 * 1.17
                    notu = "DÜŞÜK KAR: x1.17 uygulandı"
                else:
                    new_net = p1
                    notu = "Kayıtlı Fiyat Korundu"
                
                a_r.append({
                    "Malzeme Kodu": row[k3_c1], "Ürün Adı": row[k3_n1],
                    "Eski Tanımlı": round(p1, 2), "Karşı Net": round(p2, 2),
                    "Oran": round(ratio, 4), "YENİ NET FİYAT": round(new_net, 4),
                    "Barem %12": round(new_net/0.88, 4), "40+12 Liste": round(new_net/0.528, 4),
                    "Durum": notu
                })
                matched_e2.add(code)
            else: n_f.append(row.to_dict())
        
        st.session_state.k3_auto, st.session_state.k3_not_found = a_r, n_f
        st.session_state.k3_unused_e2 = df3_s[~df3_s[k3_c2].apply(clean_code).isin(matched_e2)].to_dict('records')
        st.session_state.k3_comp_list = df3_s.to_dict('records')

    if st.session_state.k3_auto or st.session_state.k3_not_found:
        t3_1, t3_2, t3_3 = st.tabs(["📊 Analiz Sonuçları", "🔍 Manuel Asistan", "❗ Karşıda Olup Eşleşmeyenler"])
        with t3_1:
            full_k3 = st.session_state.k3_auto + st.session_state.k3_manual
            if full_k3:
                df_k3 = pd.DataFrame(full_k3); st.dataframe(df_k3, use_container_width=True)
                out3 = io.BytesIO(); df_k3.to_excel(out3, index=False)
                st.download_button("📥 Karlılık Analizini İndir", out3.getvalue(), "karlilik_analizi.xlsx")
        with t3_2:
            if st.session_state.k3_not_found:
                sel = st.selectbox("Ürün Seç (E1):", st.session_state.k3_not_found, format_func=lambda x: f"{x[k3_n1]}", key="k3_sb")
                sug = sorted(st.session_state.k3_comp_list, key=lambda x: similarity(clean_name(sel[k3_n1]), clean_name(str(x[k3_c2]))), reverse=True)[:3]
                for s in sug:
                    if st.button(f"Eşleştir: {s[k3_c2]}", key=f"k3_bt_{s[k3_c2]}"):
                        p1, p2 = float(sel[k3_p1]), float(s[k3_p2])
                        ratio = p2/p1 if p1>0 else 0
                        nn = p1*1.17 if ratio <= 1.16 else p1
                        st.session_state.k3_manual.append({
                            "Malzeme Kodu": sel[k3_c1], "Ürün Adı": sel[k3_n1], "Eski Tanımlı": round(p1,2),
                            "Karşı Net": round(p2,2), "Oran": round(ratio,4), "YENİ NET FİYAT": round(nn,4),
                            "Barem %12": round(nn/0.88,4), "40+12 Liste": round(nn/0.528,4), "Durum": "Manuel Eşleşme"
                        })
                        st.session_state.k3_not_found.remove(sel); st.rerun()
        with t3_3:
            if st.session_state.k3_unused_e2:
                df_u3 = pd.DataFrame(st.session_state.k3_unused_e2); st.dataframe(df_u3, use_container_width=True)
                out3u = io.BytesIO(); df_u3.to_excel(out3u, index=False)
                st.download_button("📥 Karşı Listeyi İndir", out3u.getvalue(), "karsi_liste_fazlalik.xlsx")
