import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Konumsal Fiyat Dedektifi", layout="wide")

st.title("🔍 Akıllı Konumsal Fiyat Eşleştirici")
st.write("Kod nerede olursa olsun CTRL+F mantığıyla bulur ve o sayfadaki en alakalı fiyatı eşleştirir.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("1. Referans Excel (İsim & Kod)", type="xlsx")
with col2:
    pdf_file = st.file_uploader("2. PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("İskonto (Örn: 50+10)", value="20")

def calculate_discount(price_str, disc_str):
    try:
        num_str = price_str.replace('₺', '').replace('.', '').replace(',', '.').strip()
        val = float(num_str)
        if not disc_str: return round(val, 2)
        discounts = [float(d.strip()) for d in disc_str.split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

if reference_excel and pdf_file:
    # Excel Verilerini Oku
    ref_df = pd.read_excel(reference_excel)
    products_to_search = []
    for _, row in ref_df.iterrows():
        products_to_search.append({
            "name": str(row.iloc[0]).strip(),
            "code": str(row.iloc[1]).strip()
        })

    st.info(f"Excel'den {len(products_to_search)} ürün alındı. PDF taranıyor...")

    found_data = []
    not_found = []
    
    # Fiyat Regex: ₺25,09 veya 1.250,00 formatı
    price_pattern = re.compile(r'₺?\s?\d{1,3}(?:\.\d{3})*,\d{2}')

    with pdfplumber.open(pdf_file) as pdf:
        progress_bar = st.progress(0)
        
        for idx, target in enumerate(products_to_search):
            search_code = target['code']
            if not search_code or search_code == 'nan':
                continue
            
            is_found_for_this_code = False
            
            for page in pdf.pages:
                # 1. ADIM: KODU BUL (Konumsal olarak)
                # Sayfadaki tüm kelimeleri ve koordinatlarını al
                words = page.extract_words()
                
                # Kodun o sayfada geçtiği yerleri bul
                code_instances = [w for w in words if search_code in w['text'] or clean_string(search_code) in clean_string(w['text'])]
                
                if code_instances:
                    # Kodu bulduk! Şimdi bu sayfanın metnini alıp fiyat arayalım
                    page_text = page.extract_text()
                    
                    # Sayfa metninde tüm fiyatları bul
                    all_prices = list(price_pattern.finditer(page_text))
                    
                    if all_prices:
                        # ANO ÇITASI örneğindeki gibi kod yukarıda fiyat aşağıda olabilir.
                        # Kodun geçtiği yerden SONRAKİ ilk fiyatı veya sayfadaki en mantıklı fiyatı alalım.
                        # Basit ve etkili mantık: Kod o sayfadaysa, o sayfadaki fiyatlara bak.
                        
                        # Eğer sayfada tek bir fiyat varsa direkt onu al
                        if len(all_prices) == 1:
                            p_found = all_prices[0].group()
                        else:
                            # Birden fazla fiyat varsa, kodun geçtiği metin bloğuna en yakın olanı bulmaya çalışırız
                            # Şimdilik sayfadaki ilk fiyatı alıyoruz (Genelde o bloktaki fiyattır)
                            p_found = all_prices[0].group() 

                        found_data.append({
                            "Ürün İsmi": target['name'],
                            "Ürün Kodu": search_code,
                            "Liste Fiyatı": p_found,
                            "İskontolu Fiyat": calculate_discount(p_found, discount_input),
                            "Sayfa": page.page_number
                        })
                        is_found_for_this_code = True
                        break # Diğer sayfalara bakma
            
            if not is_found_for_this_code:
                not_found.append(target)
            
            progress_bar.progress((idx + 1) / len(products_to_search))

    # Sonuçlar
    st.divider()
    if found_data:
        res_df = pd.DataFrame(found_data).drop_duplicates(subset=['Ürün Kodu'])
        st.success(f"✅ {len(res_df)} ürün başarıyla eşleşti!")
        st.dataframe(res_df)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            res_df.to_excel(writer, index=False)
        st.download_button("📥 Excel'i İndir", out.getvalue(), "guncel_fiyatlar.xlsx")
    
    if not_found:
        with st.expander(f"❌ Bulunamayan Ürünler ({len(not_found)})"):
            st.write(pd.DataFrame(not_found))

def clean_string(s):
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()
