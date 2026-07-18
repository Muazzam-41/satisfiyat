import streamlit as st
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Kod Tabanlı Fiyat Botu", layout="wide")

def clean_code(text):
    """Kodlardaki boşluk, nokta ve tireleri temizler (Tam eşleşme için)."""
    if not text or pd.isna(text):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).upper()

def calculate_net_price(price, disc_str):
    """Zincir iskontoyu (Örn: 50+15) liste fiyatına uygular."""
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

st.title("🛡️ Malzeme Kodu Tabanlı Fiyatlandırma Sistemi")
st.write("Kodlar üzerinden birebir eşleştirme yapar ve J sütunundaki fiyattan iskontolu baremleri hesaplar.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Kendi Ürün Listeniz")
    ref_file = st.file_uploader("Excel'i yükleyin (En sol sütun kod olmalı)", type="xlsx", key="ref")
with col2:
    st.subheader("2. SVR Fiyat Listesi")
    svr_file = st.file_uploader("SVR Teknik Bayi Excel'ini yükleyin", type="xlsx", key="svr")

discount_input = st.text_input("📉 Ana İskonto Oranı (Örn: 50+15)", value="50+15")

if st.button("🚀 EŞLEŞTİRMEYİ VE HESAPLAMAYI BAŞLAT"):
    if ref_file and svr_file:
        try:
            with st.spinner('Kodlar eşleştiriliyor...'):
                # Excel'leri oku
                df_ref = pd.read_excel(ref_file)
                df_svr = pd.read_excel(svr_file)

                # SVR Dosyasını Haritalandır (B Sütunu Kod -> Index 1, J Sütunu Fiyat -> Index 9, C Sütunu İsim -> Index 2)
                # Not: Görselinizde Malzeme Kodu B (index 1) sütununda görünüyor.
                price_map = {}
                name_map = {}
                for idx, row in df_svr.iterrows():
                    try:
                        raw_code = clean_code(row.iloc[1]) # B Sütunu: Malzeme Kodu
                        raw_price = row.iloc[9]            # J Sütunu: Birim Fiyatı
                        raw_name = row.iloc[2]             # C Sütunu: Malzeme Adı
                        
                        if raw_code:
                            price_map[raw_code] = raw_price
                            name_map[raw_code] = raw_name
                    except:
                        continue

                # Kendi listenizle karşılaştırın
                final_results = []
                unmatched = []

                for idx, row in df_ref.iterrows():
                    # Kendi listenizdeki ilk sütun Malzeme Kodu varsayılıyor
                    my_code_raw = row.iloc[0] 
                    my_code_clean = clean_code(my_code_raw)

                    if my_code_clean in price_map:
                        liste_fiyati = price_map[my_code_clean]
                        urun_adi = name_map[my_code_clean]
                        
                        # Hesaplamalar
                        net_fiyat = calculate_net_price(liste_fiyati, discount_input)
                        barem_12 = net_fiyat / 0.88
                        barem_40_12 = net_fiyat / (0.60 * 0.88)

                        final_results.append({
                            "Malzeme Kodu": my_code_raw,
                            "Malzeme Adı": urun_adi,
                            "SVR Liste Fiyatı (J)": round(float(liste_fiyati), 2),
                            "Net Fiyat": round(net_fiyat, 4),
                            "Barem Liste (%12)": round(barem_12, 4),
                            "40+12 Liste": round(barem_40_12, 4)
                        })
                    else:
                        if pd.notna(my_code_raw):
                            unmatched.append({"Kod": my_code_raw})

                # Sonuçları Göster
                tab1, tab2 = st.tabs(["✅ Eşleşen Ürünler", "❌ Eşleşmeyen Kodlar"])

                with tab1:
                    if final_results:
                        res_df = pd.DataFrame(final_results)
                        st.success(f"{len(res_df)} ürün kod üzerinden başarıyla eşleşti.")
                        st.dataframe(res_df, use_container_width=True)
                        
                        out = io.BytesIO()
                        with pd.ExcelWriter(out, engine='openpyxl') as writer:
                            res_df.to_excel(writer, index=False)
                        st.download_button("📥 Hesaplanan Excel'i İndir", out.getvalue(), "hesaplanan_fiyatlar.xlsx")
                    else:
                        st.warning("Hiçbir kod eşleşmedi.")

                with tab2:
                    if unmatched:
                        st.error(f"{len(unmatched)} kod SVR listesinde bulunamadı.")
                        st.dataframe(pd.DataFrame(unmatched))

        except Exception as e:
            st.error(f"Hata: {str(e)}")
    else:
        st.warning("Lütfen her iki dosyayı da yükleyin.")
