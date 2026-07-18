import streamlit as st
import pandas as pd
import io

# Sayfa Yapılandırması
st.set_page_config(page_title="Profesyonel Excel İşlem Merkezi", layout="wide")

def calculate_net_price(price, disc_str):
    """Zincir iskontoyu uygular."""
    try:
        # Fiyat sayı değilse temizle
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

st.title("🛡️ Profesyonel Excel İşlem Merkezi")

# --- BÖLÜM 1: FİYAT HESAPLAMA ---
st.header("1️⃣ Muhasebe Entegre Fiyat Hesaplama")

col1, col2 = st.columns(2)
with col1:
    ref_file_1 = st.file_uploader("Kendi Ürün Listeniz (Fiyat İçin)", type="xlsx", key="fiyat_ref")
with col2:
    price_file_1 = st.file_uploader("SVR Fiyat Listesi (Fiyat İçin)", type="xlsx", key="fiyat_svr")

discount_input = st.text_input("📉 J Sütununa Uygulanacak İskonto (Örn: 50+15)", value="50+15", key="isc_input")

if st.button("🚀 Fiyatları Hesapla", key="btn_fiyat"):
    if ref_file_1 and price_file_1:
        try:
            with st.spinner('Veriler işleniyor...'):
                # Dosyaları oku
                df_ref = pd.read_excel(ref_file_1)
                df_price = pd.read_excel(price_file_1)

                # Boş satırları temizle (Ürün adı sütunu boş olanları at)
                df_ref = df_ref.dropna(subset=[df_ref.columns[0]])
                
                # Fiyat haritası oluştur (C sütunu isim, J sütunu fiyat)
                price_map = {}
                for idx, row in df_price.iterrows():
                    try:
                        # C sütunu index 2, J sütunu index 9
                        name_val = str(row.iloc[2]).strip().lower()
                        price_val = row.iloc[9]
                        if pd.notna(name_val) and pd.notna(price_val):
                            price_map[name_val] = price_val
                    except:
                        continue

                # Eşleştirme yap
                matched_fiyat = []
                for idx, row in df_ref.iterrows():
                    name_in_ref = str(row.iloc[0]).strip()
                    name_lower = name_in_ref.lower()

                    if name_lower in price_map:
                        liste_fiyati = price_map[name_lower]
                        net = calculate_net_price(liste_fiyati, discount_input)
                        matched_fiyat.append({
                            "Ürün Adı": name_in_ref,
                            "SVR Liste Fiyatı": round(float(liste_fiyati), 2) if pd.notna(liste_fiyati) else 0,
                            "Net Fiyat": round(net, 4),
                            "Barem Liste (%12)": round(net / 0.88, 4),
                            "40+12 Liste": round(net / 0.528, 4)
                        })

                if matched_fiyat:
                    res_df = pd.DataFrame(matched_fiyat)
                    st.success(f"✅ {len(res_df)} ürün eşleşti ve hesaplandı.")
                    st.dataframe(res_df, use_container_width=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        res_df.to_excel(writer, index=False)
                    st.download_button("📥 Hesaplanan Fiyatları İndir", output.getvalue(), "muhasebe_fiyatlari.xlsx")
                else:
                    st.warning("⚠️ Hiçbir ürün ismi eşleşmedi! Lütfen her iki Excel'deki ürün isimlerinin (boşluklar dahil) aynı olduğunu kontrol edin.")
        
        except Exception as e:
            st.error(f"❌ Bir hata oluştu: {str(e)}")
            st.info("İpucu: SVR dosyasında C sütununda isim, J sütununda fiyat olduğundan emin olun.")
    else:
        st.warning("⚠️ Lütfen her iki dosyayı da yükleyin.")

st.divider()

# --- BÖLÜM 2: KOD EŞLEŞTİRME ---
st.header("2️⃣ Ürün Kodu (Malzeme Kodu) Aktarma")
col3, col4 = st.columns(2)
with col3:
    ref_file_2 = st.file_uploader("Kendi Listeniz (Kodların Olduğu)", type="xlsx", key="kod_ref")
with col4:
    target_file_2 = st.file_uploader("SVR Dosyası (Kodların Ekleneceği)", type="xlsx", key="kod_target")

if st.button("🚀 Kodları SVR'ye Aktar", key="btn_kod"):
    if ref_file_2 and target_file_2:
        try:
            df_ref_2 = pd.read_excel(ref_file_2)
            df_svr_2 = pd.read_excel(target_file_2)

            code_map = {}
            for _, row in df_ref_2.iterrows():
                n = str(row.iloc[0]).strip().lower()
                c = str(row.iloc[1]).strip()
                if n != 'nan': code_map[n] = c if c != 'nan' else ""

            def find_code(row):
                return code_map.get(str(row.iloc[2]).strip().lower(), "")

            new_col = df_svr_2.apply(find_code, axis=1)
            df_svr_2.insert(0, 'MALZEME KODU', new_col)

            st.success("✅ Kodlar başarıyla aktarıldı.")
            st.dataframe(df_svr_2.head(10), use_container_width=True)
            
            out_s = io.BytesIO()
            with pd.ExcelWriter(out_s, engine='openpyxl') as writer:
                df_svr_2.to_excel(writer, index=False)
            st.download_button("📥 Kodlu SVR Dosyasını İndir", out_s.getvalue(), "kodlu_svr.xlsx")
        except Exception as e:
            st.error(f"Hata: {e}")
