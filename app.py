import streamlit as st
import pandas as pd
import io

# Sayfa Yapılandırması
st.set_page_config(page_title="Muhasebe Entegre Fiyat Botu", layout="wide")

def calculate_net_price(price, disc_str):
    """Girdiğiniz iskontoyu J sütununa uygular ve NET fiyatı bulur."""
    try:
        val = float(price)
        if not disc_str or disc_str == "0":
            return val
        
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return val
    except:
        return 0

st.title("🚀 Muhasebe Entegre: Ters İskonto ve Fiyat Botu")
st.write("Net fiyattan yola çıkarak, belirli iskontolarla aynı sonucu verecek 'Sanal Liste Fiyatları' oluşturur.")

# Dosya Yükleme Alanları
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Kendi Ürün Listeniz")
    ref_file = st.file_uploader("Kendi Excel'inizi yükleyin", type="xlsx", key="ref")

with col2:
    st.subheader("2. Güncel Fiyat Listesi (SVR)")
    price_file = st.file_uploader("Tedarikçi Fiyat Excel'ini (SVR) yükleyin", type="xlsx", key="price")

discount_input = st.text_input("📉 J Sütununa Uygulanacak Ana İskonto (Örn: 50+15)", value="50+15")

if st.button("🚀 HESAPLAMAYI VE TERS EŞLEŞTİRMEYİ BAŞLAT"):
    if ref_file and price_file:
        try:
            # 1. Dosyaları oku ve temizle
            df_ref = pd.read_excel(ref_file)
            df_price = pd.read_excel(price_file)

            # Boş satırları ayıkla
            df_ref = df_ref.dropna(subset=[df_ref.columns[0]])
            df_price = df_price.dropna(subset=[df_price.columns[2]])

            # Fiyat Listesi Sözlüğü (C: İsim, J: Fiyat)
            price_map = {}
            for _, row in df_price.iterrows():
                try:
                    val_name = row.iloc[2]
                    val_price = row.iloc[9]
                    if pd.notna(val_name) and pd.notna(val_price):
                        price_map[str(val_name).strip().lower()] = float(val_price)
                except:
                    continue

            # 2. Hesaplama Döngüsü
            matched_results = []
            unmatched_results = []

            for _, row in df_ref.iterrows():
                ref_name = str(row.iloc[0]).strip()
                ref_name_lower = ref_name.lower()

                if ref_name_lower in price_map:
                    original_j_price = price_map[ref_name_lower]
                    
                    # A - Gerçek Net Fiyatı Bul
                    net_fiyat = calculate_net_price(original_j_price, discount_input)
                    
                    # B - %12 İskonto yapıldığında bu net fiyatı verecek Liste Fiyatı
                    # x * 0.88 = Net -> x = Net / 0.88
                    barem_12_liste = net_fiyat / 0.88
                    
                    # C - %40 + %12 iskonto yapıldığında bu net fiyatı verecek Liste Fiyatı
                    # x * 0.60 * 0.88 = Net -> x = Net / (0.60 * 0.88)
                    barem_40_12_liste = net_fiyat / (0.60 * 0.88)
                    
                    matched_results.append({
                        "Ürün Adı": ref_name,
                        "SVR Liste Fiyatı (J)": round(original_j_price, 4),
                        "Net Fiyat": round(net_fiyat, 4),
                        "Barem Liste Fiyatı (Net / 0.88)": round(barem_12_liste, 4),
                        "40+12 Liste Fiyatı (Net / 0.528)": round(barem_40_12_liste, 4)
                    })
                else:
                    unmatched_results.append({"Ürün Adı": ref_name, "Durum": "Fiyat Listesinde Yok"})

            # 3. GÖRSELLEŞTİRME
            tab1, tab2 = st.tabs(["✅ Hesaplananlar", "❌ Bulunamayanlar"])

            with tab1:
                if matched_results:
                    res_df = pd.DataFrame(matched_results)
                    st.success(f"{len(res_df)} Ürün için ters hesaplama yapıldı.")
                    st.dataframe(res_df, use_container_width=True)
                    
                    # Excel Çıktısı
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        res_df.to_excel(writer, index=False)
                    st.download_button("📥 Muhasebe Excel'ini İndir", output.getvalue(), "muhasebe_fiyat_listesi.xlsx")

            with tab2:
                if unmatched_results:
                    st.error(f"{len(unmatched_results)} ürün eşleşmedi.")
                    st.table(pd.DataFrame(unmatched_results))

        except Exception as e:
            st.error(f"Hata: {e}")
    else:
        st.warning("Lütfen iki dosyayı da yükleyin.")
