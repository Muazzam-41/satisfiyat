import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import math

st.set_page_config(page_title="Nokta Atışı Fiyat Eşleştirici PRO", layout="wide")

def clean_string(s):
    if not s or s == 'nan': return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    try:
        clean_p = re.sub(r'[^\d,]', '', price_str).replace(',', '.')
        val = float(clean_p)
        if not disc_str or disc_str == "0": return round(val, 2)
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

def get_distance(w1, w2):
    """İki koordinat arası mesafe (Fiyat kodun altında veya sağındaysa mesafe küçük çıkar)"""
    dx = w2['x0'] - w1['x0']
    dy = w2['top'] - w1['top']
    # Fiyat yukarıdaysa uzaklığı artır (yanlış eşleşmeyi önle)
    if dy < -5: dy = abs(dy) * 5 
    return math.sqrt(dx**2 + dy**2)

st.title("🎯 Nokta Atışı Fiyat Eşleştirici (Hızlı Sürüm)")
st.write("PDF tek seferde indekslenir ve 660 ürün saniyeler içinde eşleştirilir.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="50+15")

if reference_excel and pdf_file:
    if st.button("🚀 TARAMAYI BAŞLAT"):
        # 1. Excel'i hazırlat
        ref_df = pd.read_excel(reference_excel)
        excel_products = []
        excel_codes_map = {} # Hızlı arama için temizlenmiş kod sözlüğü
        
        for _, row in ref_df.iterrows():
            orig_name = str(row.iloc[0]).strip()
            orig_code = str(row.iloc[1]).strip()
            clean_c = clean_string(orig_code)
            if clean_c:
                excel_products.append({"name": orig_name, "code": orig_code, "clean": clean_c})
                excel_codes_map[clean_c] = {"name": orig_name, "orig_code": orig_code}

        # 2. PDF'i Tek Geçişte İndeksle
        st.info("PDF analiz ediliyor, lütfen bekleyin...")
        pdf_index = [] # Her sayfadaki kod ve fiyatların yerini tutacak
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            total_pages = len(pdf.pages)
            progress_bar = st.progress(0)
            
            for p_idx, page in enumerate(pdf.pages):
                words = page.extract_words()
                full_text = page.extract_text() or ""
                
                page_codes = []
                page_prices = []
                
                # Sayfadaki tüm kelimeleri tara
                for w in words:
                    # Bu kelime Excel'deki bir kod mu?
                    w_clean = clean_string(w['text'])
                    if w_clean in excel_codes_map:
                        page_codes.append({"clean": w_clean, "coords": w})
                
                # Sayfadaki tüm fiyatları ve yerlerini bul
                for match in price_regex.finditer(full_text):
                    p_str = match.group()
                    # Fiyatın koordinatını bulmak için basit bir arama (kuruş kısmından)
                    kuruş = p_str.split(',')[-1]
                    for w in words:
                        if kuruş in w['text'] and abs(w['top'] - page.height) > 0:
                            page_prices.append({"text": p_str, "coords": w})
                            break
                
                pdf_index.append({"codes": page_codes, "prices": page_prices, "page": p_idx + 1})
                progress_bar.progress((p_idx + 1) / total_pages)

        # 3. Eşleştirme Yap (Hafızada)
        st.info("Eşleştirme yapılıyor...")
        results = []
        found_codes_set = set()

        for page_data in pdf_index:
            p_codes = page_data['codes']
            p_prices = page_data['prices']
            
            if p_codes and p_prices:
                for c_item in p_codes:
                    # Bu koda en yakın fiyatı bul
                    best_price = min(p_prices, key=lambda p: get_distance(c_item['coords'], p['coords']))
                    
                    price_val = best_price['text']
                    net_val = calculate_discount(price_val, discount_input)
                    
                    ref_info = excel_codes_map[c_item['clean']]
                    
                    results.append({
                        "Ürün İsmi": ref_info['name'],
                        "Ürün Kodu": ref_info['orig_code'],
                        "Liste Fiyatı": price_val,
                        "Net Fiyat": net_val,
                        "Sayfa": page_data['page']
                    })
                    found_codes_set.add(c_item['clean'])

        # 4. Sonuçları Göster
        if results:
            res_df = pd.DataFrame(results).drop_duplicates(subset=['Ürün Kodu'])
            st.success(f"✅ {len(res_df)} ürün başarıyla eşleşti!")
            st.dataframe(res_df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 Excel İndir", output.getvalue(), "hatasiz_fiyatlar.xlsx")
            
            # Bulunamayanlar
            not_found = [p for p in excel_products if p['clean'] not in found_codes_set]
            if not_found:
                with st.expander(f"❌ Bulunamayan Ürünler ({len(not_found)})"):
                    st.table(pd.DataFrame(not_found)[['name', 'code']])
        else:
            st.error("Hiçbir ürün eşleşmedi. Lütfen kodları kontrol edin.")
