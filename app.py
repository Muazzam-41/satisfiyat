import streamlit as st
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Gelişmiş Excel İşlem Merkezi", layout="wide")

# --- GENEL YARDIMCI FONKSİYONLAR ---
def clean_code(text):
    if not text or pd.isna(text): return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).upper()

def clean_name(text):
    """
    Ürün ismindeki ●, ▼, ★ gibi tüm sembolleri ve boşlukları temizler.
    Sadece harf ve rakamları bırakarak eşleştirme yapar.
    """
    if not text or pd.isna(text): return ""
    text = str(text).lower()
    # Türkçe karakter dönüşümleri (opsiyonel ama garantici yaklaşım)
    # Regex: Sadece alfanümerik karakterleri (harf ve rakam) tutar
    cleaned = re.sub(r'[^a-z0-9ğüşıöç]', '', text)
    return cleaned

def calculate_net_price(price, disc_str):
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

# --- ARAYÜZ BAŞLIĞI ---
st.title("🛡️ Profesyonel Excel İşlem ve Karşılaştırma Merkezi")

# ==============================================================================
# BÖLÜM 1: FİYAT HESAPLAMA (Değişmedi)
# ==============================================================================
st.header("1️⃣ Muhasebe Entegre Fiyat Hesaplama")
f1_col1, f1_col2 = st.columns(2)
with f1_col1:
    ref_file_fiyat = st.file_uploader("Kendi Ürün Listenizi Yükleyin", type="xlsx", key="fiyat_ref")
with f1_col2:
    svr_file_fiyat = st.file_uploader("SVR Fiyat Excel'ini Yükleyin", type="xlsx", key="fiyat_svr")

if ref_file_fiyat and svr_file_fiyat:
    df_ref_f = pd.read_excel(ref_file_fiyat)
    df_svr_f = pd.read_excel(svr_file_fiyat)
    
    st.write("⚙️ **Fiyat Sütun Ayarları**")
    s1, s2, s3, s4 = st.columns(4)
    with s1: ref_code_f = st.selectbox("Kendi Listeniz: Kod Sütunu", df_ref_f.columns, key="f_s1")
    with s2: svr_code_f = st.selectbox("SVR: Kod Sütunu", df_svr_f.columns, key="f_s2")
    with s3: svr_name_f = st.selectbox("SVR: İsim Sütunu", df_svr_f.columns, key="f_s3")
    with s4: svr_price_f = st.selectbox("SVR: Fiyat Sütunu", df_svr_f.columns, key="f_s4")
    
    disc_f = st.text_input("İskonto (50+15 vb.)", value="50+15", key="f_isc")

    if st.button("🚀 Fiyatları Hesapla"):
        price_map = {clean_code(r[svr_code_f]): (r[svr_price_f], r[svr_name_f]) for _, r in df_svr_f.iterrows()}
        f_results = []
        for _, row in df_ref_f.iterrows():
            c = clean_code(row[ref_code_f])
            if c in price_map:
                l_price, u_name = price_map[c]
                net = calculate_net_price(l_price, disc_f)
                f_results.append({
                    "Kod": row[ref_code_f], "İsim": u_name,
                    "Liste Fiyatı": round(float(l_price), 2), "Net Fiyat": round(net, 4),
                    "Barem %12": round(net / 0.88, 4), "40+12 Liste": round(net / 0.528, 4)
                })
        if f_results:
            st.dataframe(pd.DataFrame(f_results), use_container_width=True)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                pd.DataFrame(f_results).to_excel(writer, index=False)
            st.download_button("📥 Fiyat Excel'ini İndir", output.getvalue(), "fiyat_listesi.xlsx")

st.divider()

# ==============================================================================
# BÖLÜM 2: HİBRİT STOK KARŞILAŞTIRMA (Özel Karakter Korumalı)
# ==============================================================================
st.header("2️⃣ Stok Karşılaştırma (Hibrit Eşleştirme)")
st.info("Bu bölüm ürün başındaki ●●, ▼, ■ gibi sembolleri otomatik olarak temizleyip eşleştirme yapar.")

col_st1, col_st2 = st.columns(2)
with col_st1:
    real_stock_file = st.file_uploader("Kendi Reel Stok Excel'iniz", type="xlsx", key="st_real")
with col_st2:
    compare_stock_file = st.file_uploader("Karşılaştırılacak İkinci Excel", type="xlsx", key="st_comp")

if real_stock_file and compare_stock_file:
    df_real = pd.read_excel(real_stock_file)
    df_comp = pd.read_excel(compare_stock_file)

    st.write("⚙️ **Stok Sütun Ayarları**")
    st1, st2, st3, st4 = st.columns(4)
    with st1: r_code_col = st.selectbox("Reel Stok: Kod Sütunu", df_real.columns, key="st_s1")
    with st2: r_name_col = st.selectbox("Reel Stok: İsim Sütunu", df_real.columns, key="st_s2")
    with st3: c_code_col = st.selectbox("2. Excel: Kod Sütunu", df_comp.columns, key="st_s3")
    with st4: c_name_col = st.selectbox("2. Excel: İsim Sütunu", df_comp.columns, key="st_s4")

    if st.button("🚀 Stokları Karşılaştır"):
        try:
            # 2. Excel'i hafızaya alırken isimleri sembollerden temizleyerek haritalandır
            comp_by_code = {clean_code(r[c_code_col]): r for _, r in df_comp.iterrows() if pd.notna(r[c_code_col])}
            comp_by_name = {clean_name(r[c_name_col]): r for _, r in df_comp.iterrows() if pd.notna(r[c_name_col])}

            matched_stock = []
            not_found_stock = []

            for _, row in df_real.iterrows():
                r_code_raw = row[r_code_col]
                r_name_raw = row[r_name_col]
                
                r_code_c = clean_code(r_code_raw)
                r_name_c = clean_name(r_name_raw)

                # 1. Kod ile eşleştirme
                if r_code_c and r_code_c in comp_by_code:
                    match_row = comp_by_code[r_code_c]
                    res_item = {"Reel Kod": r_code_raw, "Reel İsim": r_name_raw, "Eşleşme Türü": "KOD İLE"}
                    # İkinci Excel'deki tüm sütunları ekle
                    for col in df_comp.columns:
                        res_item[f"Karşı_Dosya_{col}"] = match_row[col]
                    matched_stock.append(res_item)
                
                # 2. İsim ile eşleştirme (Temizlenmiş isimler üzerinden)
                elif r_name_c and r_name_c in comp_by_name:
                    match_row = comp_by_name[r_name_c]
                    res_item = {"Reel Kod": r_code_raw, "Reel İsim": r_name_raw, "Eşleşme Türü": "İSİM İLE (Semboller Temizlendi)"}
                    for col in df_comp.columns:
                        res_item[f"Karşı_Dosya_{col}"] = match_row[col]
                    matched_stock.append(res_item)
                
                # 3. Bulunamadı
                else:
                    not_found_stock.append({"Kod": r_code_raw, "İsim": r_name_raw})

            # Sonuç Tabloları
            tab_ok, tab_fail = st.tabs(["✅ Eşleşenler", "❌ Bulunamayanlar"])
            
            with tab_ok:
                if matched_stock:
                    display_df = pd.DataFrame(matched_stock)
                    st.success(f"✅ Toplam {len(display_df)} ürün sembol hassasiyetiyle eşleşti.")
                    st.dataframe(display_df, use_container_width=True)
                    
                    out_st = io.BytesIO()
                    with pd.ExcelWriter(out_st, engine='openpyxl') as writer:
                        display_df.to_excel(writer, index=False)
                    st.download_button("📥 Eşleşmiş Listeyi İndir", out_st.getvalue(), "hibrit_stok_sonuc.xlsx")
                else:
                    st.warning("Eşleşen ürün bulunamadı.")

            with tab_fail:
                if not_found_stock:
                    fail_df = pd.DataFrame(not_found_stock)
                    st.error(f"❌ {len(fail_df)} ürün hiçbir şekilde eşleşmedi.")
                    st.dataframe(fail_df, use_container_width=True)

        except Exception as e:
            st.error(f"Hata oluştu: {e}")
