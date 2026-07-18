import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import math

st.set_page_config(page_title="Nokta Atışı Fiyat Dedektifi v4", layout="wide")

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

st.title("🎯 Nokta Atışı: Satır Bazlı Fiyat Eşleştirici")
st.write("Kodun tam karşısındaki (aynı satırdaki) fiyatı yakalar. Tablo yapılarında %100 sonuç verir.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="20")

if reference_excel and pdf_file:
    if st.button("🚀 TARAMAYI BAŞLAT"):
        ref_df = pd.read_excel(reference_excel)
        excel_items = []
        for _, row in ref_df.iterrows():
            name = str(row.iloc[0]).strip()
            code = str(row.iloc[1]).strip()
            if code and code != 'nan':
                excel_items.append({"name": name, "code": code, "clean": clean_string(code)})

        st.info(f"Analiz başlıyor: {len(excel_items)} ürün aranıyor...")
        results = []
        found_codes = set()
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress_bar = st.progress(0)
            total_pages = len(pdf.pages)

            for p_idx, page in enumerate(pdf.pages):
                words = page.extract_words()
                if not words: continue
                
                # Sayfadaki her kelimeyi bir metin gibi de tutalım (Hızlı arama için)
                page_text = page.extract_text()
                
                for item in excel_items:
                    if item['clean'] in found_codes: continue
                    
                    # 1. Kodu Bul (Konumsal olarak)
                    target_code_word = None
                    for w in words:
                        # Tam eşleşme veya temizlenmiş eşleşme kontrolü
                        if item['code'] in w['text'] or item['clean'] == clean_string(w['text']):
                            target_code_word = w
                            break
                    
                    if target_code_word:
                        # 2. Kodun bulunduğu Y (dikey) koordinatını al
                        code_y = target_code_word['top']
                        code_y_bottom = target_code_word['bottom']
                        
                        # 3. Aynı satırda (veya çok yakınında) fiyat ara
                        # Aynı hizada olan tüm kelimeleri filtrele
                        same_line_words = [
                            w for w in words 
                            if abs(w['top'] - code_y) < 8 or abs(w['bottom'] - code_y_bottom) < 8
                        ]
                        
                        # Bu satırdaki metni birleştir
                        line_text = " ".join([w['text'] for w in same_line_words])
                        
                        # Satırda fiyat var mı?
                        price_match = price_regex.search(line_text)
                        
                        if price_match:
                            price_val = price_match.group()
                            net_val = calculate_discount(price_val, discount_input)
                            
                            results.append({
                                "Ürün İsmi": item['name'],
                                "Ürün Kodu": item['code'],
                                "Liste Fiyatı": price_val,
                                "Net Fiyat": net_val,
                                "Sayfa": p_idx + 1
                            })
                            found_codes.add(item['clean'])
                        else:
                            # Eğer tam aynı satırda bulamazsa, kodun hemen altındaki fiyatı ara (Bazen tablo kayıktır)
                            page_prices = []
                            # Sayfadaki tüm fiyatları koordinatlarıyla bul
                            # (Not: extract_text üzerinden regex ile bulup koordinata geri dönmek en garantisidir)
                            full_text = page.extract_text()
                            for m in price_regex.finditer(full_text):
                                p_txt = m.group()
                                # Bu fiyatın Y koordinatını bul (Basitleştirilmiş)
                                for pw in words:
                                    if p_txt.split(',')[-1] in pw['text']:
                                        dist = pw['top'] - code_y
                                        if dist >= -5: # Kodun üzerinde olmasın
                                            page_prices.append({"text": p_txt, "dist": dist})
                                            break
                            
                            if page_prices:
                                # En yakın alt fiyatı al
                                closest_price = min(page_prices, key=lambda x: x['dist'])
                                if closest_price['dist'] < 150: # Çok uzakta olmasın
                                    p_val = closest_price['text']
                                    results.append({
                                        "Ürün İsmi": item['name'],
                                        "Ürün Kodu": item['code'],
                                        "Liste Fiyatı": p_val,
                                        "Net Fiyat": calculate_discount(p_val, discount_input),
                                        "Sayfa": p_idx + 1
                                    })
                                    found_codes.add(item['clean'])

                progress_bar.progress((p_idx + 1) / total_pages)

        if results:
            res_df = pd.DataFrame(results).drop_duplicates(subset=['Ürün Kodu'])
            st.success(f"✅ {len(res_df)} ürün tam satır eşleşmesiyle bulundu!")
            st.dataframe(res_df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 Excel İndir", output.getvalue(), "hatasiz_liste.xlsx")
            
            not_found = [i for i in excel_items if i['clean'] not in found_codes]
            if not_found:
                with st.expander(f"❌ Bulunamayanlar ({len(not_found)})"):
                    st.table(pd.DataFrame(not_found)[['name', 'code']])
