import streamlit as st
import pandas as pd
import io

# Sayfa Yapılandırması
st.set_page_config(page_title="Gelişmiş Fiyat ve Kod Botu", layout="wide")

def calculate_net_price(price, disc_str):
    """Zincir iskontoyu J sütununa uygular ve NET fiyatı bulur."""
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

# --- ARAYÜZ ---
st.title("🛡️ Profesyonel Excel İşlem Merkezi")

# --- BÖLÜM 1: FİYAT HESAPLAMA ---
st.header("1️⃣ Muhasebe Entegre Fiyat Hesaplama")
st.write("Net fiyattan yola çıkarak sanal liste fiyatları oluşturur.")

col1, col2 = st.columns(2)
with col1:
    ref_file_1 = st.file_uploader("Kendi Ürün Listeniz (Fiyat İçin)", type="xlsx", key="fiyat_ref")
with col2:
    price_file_1 = st.file_uploader("SVR Fiyat Listesi (Fiyat İçin)", type="xlsx", key="fiyat_svr")

discount_input = st.text_input("📉 J Sütununa Uygulanacak İskonto (Örn: 50+15)", value="50+15", key="isc_input")

if st.button("🚀 Fiyatları Hesapla", key="btn_fiyat"):
    if ref_file_1 and price_file_1:
        df_ref = pd.read_excel(ref_file_1).dropna(subset=[pd.read_excel(ref_file_1).columns[0]])
        df_price = pd.read_excel(price_file_1).dropna(subset=[pd.read_excel(price_file_1).columns[2]])
        
        price_map = {}
        for _, row in df_price.iterrows():
            try:
                price_map[str(row.iloc[2]).strip().lower()] = float(row.iloc[9])
            except: continue

        matched_fiyat = []
        for _, row in df_ref.iterrows():
            name = str(row.iloc[0]).strip()
            if name.lower() in price_map:
                net = calculate_net_price(price_map[name.lower()], discount_input)
                matched_fiyat.append({
                    "Ürün Adı": name,
                    "Net Fiyat": round(net, 4),
                    "Barem Liste (%12)": round(net / 0.88, 4),
                    "40+12 Liste": round(net / 0.528, 4)
                })
        
        if matched_fiyat:
            res_df = pd.DataFrame(matched_fiyat)
            st.dataframe(res_df, use_container_width=True)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 Fiyat Excel'ini İndir", output.getvalue(), "muhasebe_fiyatlari.xlsx")

st.divider()

# --- BÖLÜM 2: KOD EŞLEŞTİRME ---
st.header("2️⃣ Ürün Kodu (Malzeme Kodu) Aktarma")
st.write("Kendi listenizdeki 'MALZEME KODU'nu, SVR dosyasındaki ürünün yanına ekler.")

col3, col4 = st.columns(2)
with col3:
    ref_file_2 = st.file_uploader("Kendi Listeniz (Kodların Olduğu)", type="xlsx", key="kod_ref")
    st.caption("Not: İlk sütun İsim, ikinci sütun 'MALZEME KODU' olmalıdır.")
with col4:
    target_file_2 = st.file_uploader("SVR Dosyası (Kodların Ekleneceği)", type="xlsx", key="kod_target")

if st.button("🚀 Kodları Eşleştir ve SVR'ye Aktar", key="btn_kod"):
    if ref_file_2 and target_file_2:
        try:
            # Dosyaları oku
            df_ref_2 = pd.read_excel(ref_file_2)
            df_svr_2 = pd.read_excel(target_file_2)

            # Referans haritası oluştur (İsim -> Malzeme Kodu)
            # İsim index 0, Kod index 1 varsayılıyor
            code_map = {}
            for _, row in df_ref_2.iterrows():
                name_val = str(row.iloc[0]).strip().lower()
                code_val = str(row.iloc[1]).strip()
                if pd.notna(name_val):
                    code_map[name_val] = code_found = code_val if code_val != 'nan' else ""

            # SVR dosyasını işle (İsimler index 2'de yani C sütununda)
            def find_code(row):
                name_in_svr = str(row.iloc[2]).strip().lower()
                return code_map.get(name_in_svr, "")

            # En sola 'MALZEME KODU' sütunu ekle
            new_codes = df_svr_2.apply(find_code, axis=1)
            df_svr_2.insert(0, 'MALZEME KODU', new_codes)

            st.success("Eşleştirme yapıldı. Kodlar SVR dosyasının en soluna eklendi.")
            st.dataframe(df_svr_2.head(20), use_container_width=True) # Önizleme

            # İndirme
            output_svr = io.BytesIO()
            with pd.ExcelWriter(output_svr, engine='openpyxl') as writer:
                df_svr_2.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Kodlu SVR Dosyasını İndir",
                data=output_svr.getvalue(),
                file_name="kod_eklenmis_svr.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Hata: {e}")
    else:
        st.warning("Lütfen iki Excel dosyasını da yükleyin.")
