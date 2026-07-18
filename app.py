import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Konumsal Fiyat Dedektifi", layout="wide")

# 1. FONKSİYON TANIMLARI (Hata almamak için en üstte olmalı)
def clean_string(s):
    """Kodlardaki boşluk ve tireleri temizleyerek eşleşme şansını artırır."""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    """Fiyatı temizler ve zincir iskontoyu hesaplar."""
    try:
        # Fiyat temizleme: ₺, . ve boşlukları kaldır, virgülü noktaya çevir
        num_str = price_str.replace('₺', '').replace('.', '').replace(',', '.').strip()
        val = float(num_str)
        if not disc_str: return round(val, 2)
        
        # 50+10 gibi yapıları işle
        discounts = [float(d.strip()) for d in disc_str.split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except:
        return 0

# 2. ARAYÜZ
st.title("🔍 Akıllı Konumsal Fiyat Eşleştirici")
st.write("Kod nerede olursa olsun CTRL+F mantığıyla bulur ve o sayfadaki fiyatı yakalar.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("1. Referans Excel (İsim & Kod)", type="xlsx")
with col2:
    pdf_file = st.file_uploader("2. PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("İskonto (Örn: 50+10)", value="20")

# 3. ANA MANTIK
if reference_excel and pdf_file:
    # Excel Verilerini Oku
    ref_df = pd.read_excel(reference_excel)
    products_to_search = []
    for _, row in ref_df.iterrows():
        products_to_search.append({
            "name": str(row.iloc[0]).strip(),
            "code": str(row.iloc[1]).strip()
        })

    st.info(f"Excel'den {len(products_to_search)} ürün yüklendi. PDF taranıyor...")

    found_data = []
    not_found = []
    
    # Fiyat Regex: ₺25,09 veya 1.250,00 veya 50,65 formatı
    price_pattern = re.compile(r'₺?\s?\d{1,3}(?:\.\d{3})*,\d{2}')

    with pdfplumber.open(pdf_file) as pdf:
        progress_bar = st.progress(0)
        
        for idx, target in enumerate(products_to_search):
            search_code = target['code']
            if not search_code or search_code == 'nan' or search_code == "":
                continue
            
            is_found_for_this_code = False
            clean_target_code = clean_string(search_code)
            
            for page in pdf.pages:
                # Sayfa metnini ve kelimeleri al
                page_text = page.extract_text()
                if not page_text: continue
                
                # Kod o sayfada geçiyor mu? (Esnek arama)
                if search_code in page_text or clean_target_code in clean_string(page_text):
                    
                    # Sayfadaki tüm fiyatları bul
                    all_prices = price_pattern.findall(page_text)
                    
                    if all_prices:
                        # ANO ÇITASI mantığı: Kod sayfadaysa, sayfadaki fiyatlara bak.
                        # Genelde ürün kodu ile fiyat aynı sayfada/bloktadır.
                        # Birden fazla fiyat varsa, 'Fiyat' kelimesine en yakın olanı veya ilkini alıyoruz.
                        p_found = all_prices[0] # İlk bulunan fiyatı al

                        found_data.append({
                            "Ürün İsmi": target['name'],
                            "Ürün Kodu": search_code,
                            "Liste Fiyatı": p_found,
                            "İskontolu Fiyat": calculate_discount(p_found, discount_input),
                            "Sayfa": page.page_number
                        })
                        is_found_for_this_code = True
                        break # Bu kod bulundu, diğer sayfalara bakma
            
            if not is_found_for_this_code:
                not_found.append(target)
            
            progress_bar.progress((idx + 1) / len(products_to_search))

    # 4. SONUÇLARIN GÖSTERİLMESİ
    st.divider()
    if found_data:
        res_df = pd.DataFrame(found_data).drop_duplicates(subset=['Ürün Kodu'])
        st.success(f"✅ {len(res_df)} ürün başarıyla eşleşti!")
        st.dataframe(res_df, use_container_width=True)
        
        # Excel İndirme
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            res_df.to_excel(writer, index=False)
        st.download_button("📥 Güncel Excel'i İndir", out.getvalue(), "guncel_fiyat_listesi.xlsx")
    
    if not_found:
        with st.expander(f"❌ Bulunamayan Ürünler ({len(not_found)})"):
            st.write("Bu ürünlerin kodları PDF içinde metin olarak bulunamadı. Kodları kontrol edin.")
            st.table(pd.DataFrame(not_found))
