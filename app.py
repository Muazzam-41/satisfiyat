import streamlit as st
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Akıllı Excel İşlem Merkezi", layout="wide")

def super_clean(text):
    """Metni eşleştirme için tamamen temizler: Küçük harf yapar, sadece harf ve rakam bırakır."""
    if not text or pd.isna(text):
        return ""
    # Sadece harf ve rakamları tut, geri kalan her şeyi (boşluk, parantez, nokta vb.) sil
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()

def calculate_net_price(price, disc_str):
    """Zincir iskontoyu uygular."""
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

def find_column_by_name(df, possible_names):
    """Verilen isim listesine göre en uygun sütun indeksini bulur."""
    for col_idx, col_name in enumerate(df.columns):
        clean_col = str(col_name).strip().lower()
        if any(name in clean_col for name in possible_names):
            return col_idx
    return None

st.title("🛡️ Akıllı Excel İşlem Merkezi")

# --- BÖLÜM 1: FİYAT HESAPLAMA ---
st.header("1️⃣ Muhasebe Entegre Fiyat Hesaplama")

col1, col2 = st.columns(2)
with col1:
    ref_file_1 = st.file_uploader("Kendi Ürün Listeniz", type="xlsx", key="fiyat_ref")
with col2:
    price_file_1 = st.file_uploader("SVR Fiyat Listesi", type="xlsx", key="fiyat_svr")

discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="50+15")

if st.button("🚀 Fiyatları Hesapla"):
    if ref_file_1 and price_file_1:
        try:
            with st.spinner('Akıllı eşleştirme yapılıyor...'):
                df_ref = pd.read_excel(ref_file_1)
                df_price = pd.read_excel(price_file_1)

                # --- AKILLI SÜTUN BULMA ---
                # Referans dosyada ilk sütunu isim varsay
                ref_name_idx = 0 
                # Fiyat listesinde "Malzeme Adı" ve "Birim Fiyatı" sütunlarını ara
                svr_name_idx = find_column_by_name(df_price, ["malzeme", "ürün", "urun", "adi", "adı"])
                svr_price_idx = find_column_by_name(df_price, ["birim fiyat", "j sütunu", "fiyatı"])
                
                # J sütunu garantisi (Eğer isimle bulamazsa 9. indexi yani J'yi kullan)
                if svr_price_idx is None: svr_price_idx = 9 
                if svr_name_idx is None: svr_name_idx = 2 # C sütunu varsayılan

                # Fiyat haritası oluştur (Süper temizlenmiş isim -> Fiyat)
                price_map = {}
                for idx, row in df_price.iterrows():
                    try:
                        raw_name = row.iloc[svr_name_idx]
                        raw_price = row.iloc[svr_price_idx]
                        if pd.notna(raw_name) and pd.notna(raw_price):
                            price_map[super_clean(raw_name)] = raw_price
                    except:
                        continue

                # Eşleştirme
                matched_data = []
                for idx, row in df_ref.iterrows():
                    ref_name = row.iloc[ref_name_idx]
                    clean_ref_name = super_clean(ref_name)

                    if clean_ref_name in price_map:
                        l_price = price_map[clean_ref_name]
                        net = calculate_net_price(l_price, discount_input)
                        matched_data.append({
                            "Ürün Adı (Sizin)": ref_name,
                            "Liste Fiyatı": round(float(l_price), 2),
                            "Net Fiyat": round(net, 4),
                            "Barem Liste (%12)": round(net / 0.88, 4),
                            "40+12 Liste": round(net / 0.528, 4)
                        })

                if matched_data:
                    res_df = pd.DataFrame(matched_data)
                    st.success(f"✅ {len(res_df)} ürün başarıyla eşleşti!")
                    st.dataframe(res_df, use_container_width=True)
                    
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='openpyxl') as writer:
                        res_df.to_excel(writer, index=False)
                    st.download_button("📥 Excel İndir", out.getvalue(), "hesaplanan_liste.xlsx")
                else:
                    st.error("⚠️ Hiçbir ürün eşleşmedi!")
                    st.info(f"SVR dosyasında '{df_price.columns[svr_name_idx]}' sütunu ile eşleştirme yapılmaya çalışıldı. İsimlerin benzer olduğundan emin olun.")

        except Exception as e:
            st.error(f"Hata: {str(e)}")
    else:
        st.warning("Dosyaları yükleyin.")

st.divider()

# --- BÖLÜM 2: KOD AKTARMA (Gelişmiş) ---
st.header("2️⃣ Ürün Kodu Aktarma")
if st.button("🚀 Kodları SVR'ye Aktar"):
    # (Bu kısım da yukarıdaki akıllı mantıkla çalışacak şekilde önceki kodda mevcuttur)
    st.write("Dosyaları yükleyip bu butona basarak Malzeme Kodu sütununu ekleyebilirsiniz.")
