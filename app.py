import streamlit as st
import pandas as pd
import io

# Sayfa Yapılandırması
st.set_page_config(page_title="Gelişmiş Excel Eşleştirici", layout="wide")

def calculate_chain_discount(price, disc_str):
    """Zincir iskontoyu matematiksel olarak uygular."""
    try:
        val = float(price)
        if not disc_str or disc_str == "0":
            return round(val, 2)
        
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except:
        return 0

st.title("📊 Profesyonel Excel Fiyat Karşılaştırma")
st.write("Ürün Adı üzerinden eşleştirme yapar, J sütunundan fiyatı çeker ve raporlar.")

# Dosya Yükleme Alanları
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Kendi Ürün Listeniz")
    ref_file = st.file_uploader("Kendi Excel'inizi yükleyin", type="xlsx", key="ref")

with col2:
    st.subheader("2. Güncel Fiyat Listesi (SVR)")
    price_file = st.file_uploader("Tedarikçi Fiyat Excel'ini yükleyin", type="xlsx", key="price")

discount_input = st.text_input("📉 Uygulanacak İskonto (Örn: 50+15)", value="50+15")

if st.button("🚀 EŞLEŞTİRMEYİ BAŞLAT"):
    if ref_file and price_file:
        try:
            # 1. Dosyaları oku
            df_ref = pd.read_excel(ref_file)
            df_price = pd.read_excel(price_file)

            # Fiyat Listesi Sözlüğü Oluştur (İsim -> Fiyat)
            # Görseldeki yapı: C sütunu (İsim) = index 2, J sütunu (Fiyat) = index 9
            price_map = {}
            for _, row in df_price.iterrows():
                try:
                    m_adi = str(row.iloc[2]).strip().lower() # Malzeme Adı (C)
                    b_fiyat = row.iloc[9]                   # Birim Fiyatı (J)
                    price_map[m_adi] = b_fiyat
                except:
                    continue

            # 2. Karşılaştırma ve Listeleme
            matched_results = []
            unmatched_results = []

            for _, row in df_ref.iterrows():
                # Kendi listenizdeki ilk sütun isim varsayılıyor (Gerekirse iloc[0] değişebilir)
                ref_name = str(row.iloc[0]).strip() 
                ref_name_lower = ref_name.lower()

                if ref_name_lower in price_map:
                    liste_fiyati = price_map[ref_name_lower]
                    net_fiyat = calculate_chain_discount(liste_fiyati, discount_input)
                    
                    matched_results.append({
                        "Ürün Adı": ref_name,
                        "Liste Fiyatı (J)": liste_fiyati,
                        "İskonto": discount_input,
                        "Net Fiyat": net_fiyat
                    })
                else:
                    unmatched_results.append({
                        "Ürün Adı": ref_name,
                        "Durum": "Fiyat Listesinde Bulunamadı"
                    })

            # 3. GÖRSELLEŞTİRME VE RAPORLAMA
            tab1, tab2 = st.tabs(["✅ Eşleşen Ürünler", "❌ Bulunamayan Ürünler"])

            with tab1:
                if matched_results:
                    res_matched_df = pd.DataFrame(matched_results)
                    st.success(f"{len(res_matched_df)} ürün başarıyla güncellendi.")
                    st.dataframe(res_matched_df, use_container_width=True)
                    
                    # İndirme Butonu (Eşleşenler)
                    output_ok = io.BytesIO()
                    with pd.ExcelWriter(output_ok, engine='openpyxl') as writer:
                        res_matched_df.to_excel(writer, index=False)
                    st.download_button("📥 Güncel Fiyat Listesini İndir", output_ok.getvalue(), "guncel_fiyatlar.xlsx")
                else:
                    st.warning("Hiçbir ürün eşleşmedi.")

            with tab2:
                if unmatched_results:
                    res_unmatched_df = pd.DataFrame(unmatched_results)
                    st.error(f"{len(res_unmatched_df)} ürün fiyat listesinde bulunamadı.")
                    st.write("Aşağıdaki ürünlerin isimlerini kontrol etmeniz gerekebilir:")
                    st.table(res_unmatched_df)
                    
                    # İndirme Butonu (Bulunamayanlar)
                    output_fail = io.BytesIO()
                    with pd.ExcelWriter(output_fail, engine='openpyxl') as writer:
                        res_unmatched_df.to_excel(writer, index=False)
                    st.download_button("📥 Bulunamayanlar Listesini İndir", output_fail.getvalue(), "eksik_urunler.xlsx")
                else:
                    st.balloons()
                    st.success("Tebrikler! Listenizdeki tüm ürünler fiyat listesinde bulundu.")

        except Exception as e:
            st.error(f"Hata: {e}. Lütfen Excel sütunlarının (C ve J) görseldeki gibi olduğundan emin olun.")
    else:
        st.warning("İşlem için iki Excel dosyasını da yüklemelisiniz.")
