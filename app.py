import streamlit as st
import pandas as pd
import io

# Sayfa Yapılandırması
st.set_page_config(page_title="Hatasız Excel Eşleştirici", layout="wide")

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
st.write("Boş satırları ayıklar, sadece dolu hücreleri karşılaştırır.")

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

            # --- TEMİZLİK AŞAMASI ---
            # İsim sütunu boş olan satırları tamamen sil (NAN engelleme)
            df_ref = df_ref.dropna(subset=[df_ref.columns[0]])
            df_price = df_price.dropna(subset=[df_price.columns[2]])

            # Fiyat Listesi Sözlüğü Oluştur (İsim -> Fiyat)
            # C sütunu (İsim) = index 2, J sütunu (Fiyat) = index 9
            price_map = {}
            for _, row in df_price.iterrows():
                try:
                    val_name = row.iloc[2]
                    val_price = row.iloc[9]
                    
                    # Eğer isim veya fiyat gerçekten boş değilse haritaya ekle
                    if pd.notna(val_name) and pd.notna(val_price):
                        m_adi = str(val_name).strip().lower()
                        price_map[m_adi] = val_price
                except:
                    continue

            # 2. Karşılaştırma
            matched_results = []
            unmatched_results = []

            for _, row in df_ref.iterrows():
                val_ref_name = row.iloc[0]
                
                # Sadece dolu satırları işle
                if pd.notna(val_ref_name):
                    ref_name = str(val_ref_name).strip()
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
                        # Gerçekten bulunamayan ama ismi dolu olan ürün
                        unmatched_results.append({
                            "Ürün Adı": ref_name,
                            "Durum": "Fiyat Listesinde Yok"
                        })

            # 3. GÖRSELLEŞTİRME
            tab1, tab2 = st.tabs(["✅ Eşleşen Ürünler", "❌ Bulunamayan Ürünler"])

            with tab1:
                if matched_results:
                    res_matched_df = pd.DataFrame(matched_results)
                    st.success(f"{len(res_matched_df)} ürün başarıyla eşleşti.")
                    st.dataframe(res_matched_df, use_container_width=True)
                    
                    output_ok = io.BytesIO()
                    with pd.ExcelWriter(output_ok, engine='openpyxl') as writer:
                        res_matched_df.to_excel(writer, index=False)
                    st.download_button("📥 Excel İndir", output_ok.getvalue(), "guncel_fiyatlar.xlsx")
                else:
                    st.warning("Eşleşen ürün bulunamadı.")

            with tab2:
                if unmatched_results:
                    res_unmatched_df = pd.DataFrame(unmatched_results)
                    st.error(f"{len(res_unmatched_df)} ürün listelerde uyuşmuyor.")
                    st.dataframe(res_unmatched_df, use_container_width=True)
                    
                    output_fail = io.BytesIO()
                    with pd.ExcelWriter(output_fail, engine='openpyxl') as writer:
                        res_unmatched_df.to_excel(writer, index=False)
                    st.download_button("📥 Eksikler Listesini İndir", output_fail.getvalue(), "eksik_urunler.xlsx")
                else:
                    st.success("Tüm ürünler başarıyla eşleşti!")

        except Exception as e:
            st.error(f"Hata oluştu: {e}")
    else:
        st.warning("İşlem için iki dosyayı da yükleyin.")
