import streamlit as st
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Hatasız Kod Eşleştirici", layout="wide")

def clean_code(text):
    """Kodları eşleşme için temizler."""
    if not text or pd.isna(text):
        return ""
    # Sadece harf ve rakamları tutar (Boşluk, nokta, tire silinir)
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).upper()

def calculate_net_price(price, disc_str):
    """Zincir iskontoyu hatasız uygular."""
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

def find_column_index(df, target_name):
    """Excel içinde başlığa göre sütun indeksini bulur."""
    target_name = str(target_name).strip().lower()
    for i, col in enumerate(df.columns):
        if target_name in str(col).strip().lower():
            return i
    # Başlıklarda bulamazsa ilk satırlarda ara (Bazı Excel'lerde başlık 1. satırda olmayabilir)
    for i in range(len(df.columns)):
        first_rows = df.iloc[:5, i].astype(str).str.lower()
        if first_rows.str.contains(target_name).any():
            return i
    return None

st.title("🎯 Nokta Atışı: Sütun İsmiyle Eşleştirme")
st.write("Sistem artık sütun sırasına değil, doğrudan 'Malzeme Kodu' başlığına bakar.")

col1, col2 = st.columns(2)
with col1:
    ref_file = st.file_uploader("1. Kendi Ürün Listeniz", type="xlsx", key="ref")
with col2:
    svr_file = st.file_uploader("2. SVR Teknik Bayi Listesi", type="xlsx", key="svr")

discount_input = st.text_input("📉 Uygulanacak İskonto (Örn: 50+15)", value="50+15")

if st.button("🚀 VERİLERİ EŞLEŞTİR VE HESAPLA"):
    if ref_file and svr_file:
        try:
            # Excel'leri oku
            df_ref = pd.read_excel(ref_file)
            df_svr = pd.read_excel(svr_file)

            # --- SÜTUNLARI TESPİT ET ---
            # Kendi listenizde "Malzeme Kodu" sütununu bul
            ref_code_idx = find_column_index(df_ref, "Malzeme Kodu")
            
            # SVR dosyasında gerekli sütunları bul
            svr_code_idx = find_column_index(df_svr, "Malzeme Kodu")
            svr_name_idx = find_column_index(df_svr, "Malzeme Adı")
            svr_price_idx = find_column_index(df_svr, "Birim Fiyatı") # J Sütunu

            # Eksik sütun kontrolü
            if ref_code_idx is None or svr_code_idx is None:
                st.error("HATA: Her iki Excel'de de 'Malzeme Kodu' başlıklı bir sütun bulunamadı!")
                st.stop()
            
            if svr_price_idx is None:
                st.warning("Uyarı: SVR'de 'Birim Fiyatı' başlığı bulunamadı, varsayılan olarak J sütununa (10. sütun) bakılıyor.")
                svr_price_idx = 9

            # --- VERİ HARİTASINI OLUŞTUR (SVR) ---
            price_map = {}
            name_map = {}
            for _, row in df_svr.iterrows():
                try:
                    c_code = clean_code(row.iloc[svr_code_idx])
                    c_price = row.iloc[svr_price_idx]
                    c_name = row.iloc[svr_name_idx] if svr_name_idx is not None else "İsim Yok"
                    
                    if c_code:
                        price_map[c_code] = c_price
                        name_map[c_code] = c_name
                except: continue

            # --- EŞLEŞTİRME ---
            results = []
            failed = []

            for _, row in df_ref.iterrows():
                raw_code = str(row.iloc[ref_code_idx]).strip()
                cleaned_ref_code = clean_code(raw_code)

                if cleaned_ref_code in price_map:
                    l_price = price_map[cleaned_ref_code]
                    net = calculate_net_price(l_price, discount_input)
                    
                    results.append({
                        "Malzeme Kodu": raw_code,
                        "Malzeme Adı": name_map[cleaned_ref_code],
                        "Liste Fiyatı": l_price,
                        "Net Fiyat": round(net, 4),
                        "Barem %12": round(net / 0.88, 4),
                        "40+12 Barem": round(net / 0.528, 4)
                    })
                else:
                    if raw_code != "nan" and raw_code != "":
                        failed.append({"Kod": raw_code})

            # --- ÇIKTI ---
            t1, t2 = st.tabs(["✅ Eşleşenler", "❌ Bulunamayan Kodlar"])
            with t1:
                if results:
                    res_df = pd.DataFrame(results)
                    st.success(f"{len(res_df)} ürün doğru kodla eşleşti.")
                    st.dataframe(res_df)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        res_df.to_excel(writer, index=False)
                    st.download_button("📥 Excel İndir", output.getvalue(), "fiyat_listesi.xlsx")
            
            with t2:
                if failed:
                    st.error(f"{len(failed)} kod SVR'de bulunamadı.")
                    st.dataframe(pd.DataFrame(failed))

        except Exception as e:
            st.error(f"Sistemsel Hata: {e}")
