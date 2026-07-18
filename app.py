import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import math

st.set_page_config(page_title="Hatasız Fiyat Eşleştirici PRO", layout="wide")

def clean_string(s):
    if not s or s == 'nan': return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    try:
        # Fiyatı temizle: ₺, nokta ve boşlukları at, virgülü noktaya çevir
        num_str = re.sub(r'[^\d,]', '', price_str).replace(',', '.')
        val = float(num_str)
        if not disc_str or disc_str == "0": return round(val, 2)
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

def get_best_match_score(code_word, price_item):
    """
    Fiyatın kodun altında ve aynı sütunda olup olmadığını kontrol eder.
    """
    dy = price_item['top'] - code_word['top']
    dx = abs(price_item['x0'] - code_word['x0'])
    
    # 1. Kural: Fiyat kodun yukarısındaysa (5 birimden fazla) eşleşme imkansız
    if dy < -5:
        return 999999
    
    # 2. Kural: Fiyat kodun çok uzağındaysa (yan sütun kontrolü)
    # Genelde katalog sütunları sayfa genişliğinin 1/3'ü kadardır
    if dx > 250:
        return 999999

    # 3. Kural: Dikey mesafe önceliklidir (dy)
    # Aynı hizada (satırda) olması en yüksek puanı alır
    return dy + (dx * 0.5)

st.title("🎯 Nokta Atışı Fiyat Eşleştirici")
st.write("Her kod için en yakın dikey fiyatı bulur. Yan yana veya alt alta fiyat kaymalarını önler.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="20")

if reference_excel and pdf_file:
    if st.button("🚀 İŞLEMİ BAŞLAT"):
        # Excel hazırlığı
        ref_df = pd.read_excel(reference_excel)
        excel_items = []
        excel_map = {}
        for _, row in ref_df.iterrows():
            name, code = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
            c_code = clean_string(code)
            if c_code:
                excel_items.append({"name": name, "code": code, "clean": c_code})
                excel_map[c_code] = {"name": name, "orig_code": code}

        st.info("PDF Analiz ediliyor...")
        final_results = []
        found_codes = set()
        price_regex = re.compile(r'₺?\s?\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress = st.progress(0)
            total_pages = len(pdf.pages)

            for p_idx, page in enumerate(pdf.pages):
                words = page.extract_words()
                text = page.extract_text() or ""
                
                # Sayfadaki tüm fiyat adaylarını koordinatıyla bul
                page_prices = []
                for match in price_regex.finditer(text):
                    p_str = match.group()
                    kuruş = p_str.split(',')[-1]
                    for w in words:
                        if kuruş in w['text']:
                            page_prices.append({"text": p_str, "top": w['top'], "x0": w['x0']})
                            break
                
                if not page_prices:
                    continue

                # Sayfadaki her kelimeyi tara
                for w in words:
                    w_clean = clean_string(w['text'])
                    
                    # Kelime Excel listemizde var mı?
                    matched_key = None
                    for k in excel_map.keys():
                        if k == w_clean or (len(w_clean) > 4 and w_clean in k):
                            matched_key = k
                            break
                    
                    if matched_key:
                        # Bu kod için o sayfadaki EN UYGUN fiyatı bul
                        best_price = None
                        min_score = 999998
                        
                        for p_item in page_prices:
                            score = get_best_match_score(w, p_item)
                            if score < min_score:
                                min_score = score
                                best_price = p_item['text']
                        
                        if best_price and min_score < 1000:
                            net_val = calculate_discount(best_price, discount_input)
                            final_results.append({
                                "Ürün İsmi": excel_map[matched_key]['name'],
                                "Ürün Kodu": excel_map[matched_key]['orig_code'],
                                "Liste Fiyatı": best_price,
                                "Net Fiyat": net_val,
                                "Sayfa": p_idx + 1
                            })
                            found_codes.add(matched_key)
                
                progress.progress((p_idx + 1) / total_pages)

        if final_results:
            res_df = pd.DataFrame(final_results).drop_duplicates(subset=['Ürün Kodu'])
            st.success(f"✅ {len(res_df)} Ürün Eşleşti!")
            st.dataframe(res_df, use_container_width=True)
            
            # İndirme
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 Excel İndir", out.getvalue(), "guncel_fiyat_listesi.xlsx")
            
            # Bulunamayanlar
            not_found = [v for k, v in excel_map.items() if k not in found_codes]
            if not_found:
                with st.expander(f"❌ Bulunamayan Ürünler ({len(not_found)})"):
                    st.table(pd.DataFrame(not_found))
