import streamlit as st
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Garantili Fiyat Botu", layout="wide")

def clean_code(text):
    """Kodları eşleşme için tamamen temizler."""
    if not text or pd.isna(text):
        return ""
    # Sadece harf ve rakamları tut, boşluk/nokta/tire sil
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).upper()

def calculate_net_price(price, disc_str):
    """Tanımlı fiyata zincir iskontoyu uygular."""
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

st.title("🎯 Garantili Fiyat & Kod Eşleştirici")
st.write("Sütunları manuel seçerek yanlış veri çekme hatasını tamamen ortadan kaldırın.")

# 1. DOSYA YÜKLEME
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Kendi Ürün Listeniz")
    ref_file = st.file_uploader("Kendi Excel'inizi yükleyin", type="xlsx", key="ref")
with col2:
    st.subheader("2. SVR Fiyat Listesi")
    svr_file = st.file_uploader("SVR Fiyat Excel'ini yükleyin", type="xlsx", key="svr")

if ref_file and svr_file:
    # Excel dosyalarını önizleme için oku
    df_ref = pd.read_excel(ref_file)
    df_svr = pd.read_excel(svr_file)

    st.divider()
    st.subheader("⚙️ Sütun Ayarları")
    st.write("Lütfen dosyalarınızdaki doğru sütun başlıklarını aşağıdan seçin:")

    c1, c2, c3 = st.columns(3)
    
    with c1:
        # Kendi listesindeki kod sütunu
        ref_code_col = st.selectbox("Kendi Listenizde 'Malzeme Kodu' Sütunu hangisi?", df_ref.columns)
    
    with c2:
        # SVR listesindeki kod sütunu
        svr_code_col = st.selectbox("SVR Listesinde 'Malzeme Kodu' Sütunu hangisi?", df_svr.columns)
        svr_name_col = st.selectbox("SVR Listesinde 'Malzeme Adı' Sütunu hangisi?", df_svr.columns)
        
    with c3:
        # SVR listesindeki fiyat sütunu
        svr_price_col = st.selectbox("SVR Listesinde 'Birim Fiyatı (Tanımlı)' Sütunu hangisi?", df_svr.columns)
        discount_input = st.text_input("Uygulanacak İskonto (Örn: 50+15)", value="50+15")

    if st.button("🚀 EŞLEŞTİRMEYİ VE HESAPLAMAYI BAŞLAT"):
        try:
            with st.spinner('Veriler işleniyor...'):
                # SVR Verilerini Sözlüğe Al
                price_map = {}
                name_map = {}
                
                for _, row in df_svr.iterrows():
                    code_raw = row[svr_code_col]
                    clean_c = clean_code(code_raw)
                    if clean_c:
                        price_map[clean_c] = row[svr_price_col]
                        name_map[clean_c] = row[svr_name_col]

                # Kendi Listenizle Eşleştir
                results = []
                not_found = []

                for _, row in df_ref.iterrows():
                    my_code_raw = row[ref_code_col]
                    my_clean = clean_code(my_code_raw)

                    if my_clean in price_map:
                        liste_fiyat = price_map[my_clean]
                        net_fiyat = calculate_net_price(liste_fiyat, discount_input)
                        
                        results.append({
                            "Malzeme Kodu": my_code_raw,
                            "Malzeme Adı": name_map[my_clean],
                            "Birim Fiyatı (Tanımlı)": round(float(liste_fiyat), 2) if pd.notna(liste_fiyat) else 0,
                            "Net Fiyat": round(net_fiyat, 4),
                            "Barem Liste (%12)": round(net_fiyat / 0.88, 4),
                            "40+12 Liste": round(net_fiyat / (0.60 * 0.88), 4)
                        })
                    else:
                        if pd.notna(my_code_raw):
                            not_found.append({"Kod": my_code_raw})

                # 4. SONUÇLAR
                tab1, tab2 = st.tabs(["✅ Başarılı Eşleşmeler", "❌ Bulunamayan Kodlar"])
                
                with tab1:
                    if results:
                        res_df = pd.DataFrame(results)
                        st.success(f"{len(res_df)} ürün başarıyla eşleşti.")
                        st.dataframe(res_df, use_container_width=True)
                        
                        out = io.BytesIO()
                        with pd.ExcelWriter(out, engine='openpyxl') as writer:
                            res_df.to_excel(writer, index=False)
                        st.download_button("📥 Excel Olarak İndir", out.getvalue(), "guncel_liste.xlsx")
                
                with tab2:
                    if not_found:
                        st.error(f"{len(not_found)} kod fiyat listesinde bulunamadı.")
                        st.dataframe(pd.DataFrame(not_found))

        except Exception as e:
            st.error(f"Hata: {e}")
else:
    st.info("Lütfen işlem yapmak için iki Excel dosyasını da yükleyin.")
