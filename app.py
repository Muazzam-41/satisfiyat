import streamlit as st
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Excel Fiyat Karşılaştırıcı", layout="wide")

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

st.title("📊 Excel Ürün Adı Eşleştirme & Fiyatlandırma")
st.write("İki Excel dosyasını 'Ürün Adı' üzerinden karşılaştırır ve J sütunundaki fiyatı çekerek iskonto uygular.")

# Dosya Yükleme Alanları
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Kendi Ürün Listeniz")
    ref_file = st.file_uploader("Ürün isimlerinin olduğu Excel'i yükleyin", type="xlsx", key="ref")

with col2:
    st.subheader("2. Güncel Fiyat Listesi (Tedarikçi)")
    price_file = st.file_uploader("Fiyatların olduğu Excel'i (SVR) yükleyin", type="xlsx", key="price")

discount_input = st.text_input("📉 Uygulanacak İskonto Oranı (Örn: 50+15)", value="50+15")

if st.button("🚀 EŞLEŞTİRMEYİ VE HESAPLAMAYI BAŞLAT"):
    if ref_file and price_file:
        try:
            # 1. Dosyaları oku
            # Fiyat listesini (SVR) okurken Malzeme Adı (C) ve Birim Fiyatı (J) sütunlarına odaklanıyoruz
            df_ref = pd.read_excel(ref_file)
            df_price = pd.read_excel(price_file)

            # Görseldeki yapıya göre (C sütunu isim, J sütunu fiyat)
            # Not: Python'da sütunlar 0'dan başlar (A=0, B=1, C=2... J=9)
            # Eğer Excel'de başlık satırı farklıysa sütun isimlerini manuel belirleyelim:
            
            # Fiyat Listesi Sözlüğü Oluştur (İsim -> Fiyat)
            # Harf duyarlılığını ve boşlukları temizleyerek eşleştirme gücünü artırıyoruz
            price_map = {}
            for _, row in df_price.iterrows():
                try:
                    m_adi = str(row.iloc[2]).strip().lower() # C Sütunu (Malzeme Adı)
                    b_fiyat = row.iloc[9]                   # J Sütunu (Birim Fiyatı)
                    price_map[m_adi] = b_fiyat
                except:
                    continue

            # 2. Karşılaştırma ve Yeni Liste Oluşturma
            results = []
            for _, row in df_ref.iterrows():
                ref_name = str(row.iloc[0]).strip() # Kendi listenizdeki ilk sütun isim varsayılıyor
                ref_name_lower = ref_name.lower()

                if ref_name_lower in price_map:
                    liste_fiyati = price_map[ref_name_lower]
                    net_fiyat = calculate_chain_discount(liste_fiyati, discount_input)
                    
                    results.append({
                        "Ürün Adı": ref_name,
                        "Eski Liste Fiyatı (J Sütunu)": liste_fiyati,
                        "İskonto": discount_input,
                        "Yeni İskontolu Fiyat": net_fiyat,
                        "Durum": "Eşleşti"
                    })
                else:
                    results.append({
                        "Ürün Adı": ref_name,
                        "Eski Liste Fiyatı (J Sütunu)": "-",
                        "İskonto": "-",
                        "Yeni İskontolu Fiyat": "-",
                        "Durum": "Bulunamadı"
                    })

            # 3. Sonuçları Göster ve İndir
            res_df = pd.DataFrame(results)
            st.success(f"İşlem Tamamlandı! {len(res_df[res_df['Durum'] == 'Eşleşti'])} ürün başarıyla eşleşti.")
            st.dataframe(res_df, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Hazırlanan Excel'i İndir",
                data=output.getvalue(),
                file_name="guncellenmis_fiyat_listesi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Bir hata oluştu: {e}")
    else:
        st.warning("Lütfen her iki Excel dosyasını da yükleyin.")
