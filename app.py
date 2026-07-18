import streamlit as st
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Tanımlı Fiyat Eşleştirici", layout="wide")

def clean_code(text):
    """Kodları eşleşme için temizler: Boşluk, nokta, tire siler."""
    if not text or pd.isna(text):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).upper()

def calculate_net_price(price, disc_str):
    """Birim Fiyatı üzerinden zincir iskontoyu uygular."""
    try:
        if isinstance(price, str):
            price = price.replace('₺', '').replace('.', '').replace(',', '.').strip()
        val = float(price)
        if not disc_str or disc_str == "0":
            return val
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return val
    except:
        return 0

def find_column_index(df, target_names):
    """Verilen alternatif isimlere göre sütun indeksini bulur."""
    for i, col in enumerate(df.columns):
        clean_col = str(col).strip().lower()
        if any(name.lower() in clean_col for name in target_names):
            return i
    return None

st.title("🎯 Nokta Atışı: Tanımlı Fiyatlandırma Sistemi")
st.write("Eşleştirme: **Malzeme Kodu** | Fiyat Kaynağı: **Birim Fiyatı (Tanımlı)**")

col1, col2 = st.columns(2)
with col1:
    ref_file = st.file_uploader("1. Kendi Ürün Listeniz (Excel)", type="xlsx", key="ref")
with col2:
    svr_file = st.file_uploader("2. SVR Fiyat Listesi (Excel)", type="xlsx", key="svr")

discount_input = st.text_input("📉 J Sütununa Uygulanacak Ana İskonto (Örn: 50+15)", value="50+15")

if st.button("🚀 HESAPLAMAYI BAŞLAT"):
    if ref_file and svr_file:
        try:
            with st.spinner('Veriler işleniyor...'):
                # Excel'leri oku
                df_ref = pd.read_excel(ref_file)
                df_svr = pd.read_excel(svr_file)

                # --- SÜTUNLARI TESPİT ET ---
                # Referans Excel'de Malzeme Kodu sütununu bul
                ref_code_idx = find_column_index(df_ref, ["Malzeme Kodu", "Kod"])
                
                # SVR Excel'de Malzeme Kodu, Adı ve "Birim Fiyatı (Tanımlı)" sütunlarını bul
                svr_code_idx = find_column_index(df_svr, ["Malzeme Kodu"])
                svr_name_idx = find_column_index(df_svr, ["Malzeme Adı"])
                # Özellikle sizin istediğiniz sütun ismi:
                svr_price_idx = find_column_index(df_svr, ["Birim Fiyatı (Tanımlı)"])

                # Hata kontrolleri
                if ref_code_idx is None:
                    st.error("HATA: Kendi listenizde 'Malzeme Kodu' sütunu bulunamadı!")
                    st.stop()
                if svr_code_idx is None:
                    st.error("HATA: SVR listesinde 'Malzeme Kodu' sütunu bulunamadı!")
                    st.stop()
                if svr_price_idx is None:
                    st.error("HATA: SVR listesinde 'Birim Fiyatı (Tanımlı)' sütunu bulunamadı! Lütfen sütun ismini kontrol edin.")
                    st.stop()

                # --- SVR VERİLERİNİ HAFIZAYA AL ---
                price_map = {}
                name_map = {}
                for _, row in df_svr.iterrows():
                    try:
                        c_code = clean_code(row.iloc[svr_code_idx])
                        c_price = row.iloc[svr_price_idx]
                        c_name = row.iloc[svr_name_idx] if svr_name_idx is not None else "İsim Bilgisi Yok"
                        
                        if c_code:
                            price_map[c_code] = c_price
                            name_map[c_code] = c_name
                    except: continue

                # --- EŞLEŞTİRME VE HESAPLAMA ---
                final_results = []
                not_found = []

                for _, row in df_ref.iterrows():
                    raw_ref_code = str(row.iloc[ref_code_idx]).strip()
                    clean_ref_code = clean_code(raw_ref_code)

                    if clean_ref_code in price_map:
                        liste_fiyat = price_map[clean_ref_code]
                        net_fiyat = calculate_net_price(liste_fiyat, discount_input)
                        
                        # İstediğiniz Barem Hesaplamaları
                        final_results.append({
                            "Malzeme Kodu": raw_ref_code,
                            "Malzeme Adı": name_map[clean_ref_code],
                            "Birim Fiyatı (Tanımlı)": round(float(liste_fiyat), 2) if pd.notna(liste_fiyat) else 0,
                            "Net Fiyat": round(net_fiyat, 4),
                            "Barem Liste (%12)": round(net_fiyat / 0.88, 4),
                            "40+12 Liste": round(net_fiyat / (0.60 * 0.88), 4)
                        })
                    else:
                        if raw_ref_code != "nan" and raw_ref_code != "":
                            not_found.append({"Kod": raw_ref_code})

                # --- SONUÇLARI GÖSTER ---
                tab1, tab2 = st.tabs(["✅ Eşleşen ve Hesaplananlar", "❌ Bulunamayan Kodlar"])
                
                with tab1:
                    if final_results:
                        res_df = pd.DataFrame(final_results)
                        st.success(f"Başarılı: {len(res_df)} ürün eşleşti.")
                        st.dataframe(res_df, use_container_width=True)
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            res_df.to_excel(writer, index=False)
                        st.download_button("📥 Excel Olarak İndir", output.getvalue(), "guncel_fiyat_listesi.xlsx")
                    else:
                        st.warning("Eşleşen ürün bulunamadı.")

                with tab2:
                    if not_found:
                        st.error(f"{len(not_found)} kod SVR listesinde bulunamadı.")
                        st.table(pd.DataFrame(not_found))

        except Exception as e:
            st.error(f"Sistem Hatası: {e}")
    else:
        st.warning("Lütfen her iki dosyayı da yükleyin.")
