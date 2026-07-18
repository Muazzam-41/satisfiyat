import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import math

st.set_page_config(page_title="Nokta Atışı Fiyat Eşleştirici v3", layout="wide")

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

def get_distance_score(code_y, price_y, code_x, price_x):
    """Fiyatın kodun altında ve yakınında olmasını ödüllendirir."""
    dy = price_y - code_y
    dx = abs(price_x - code_x)
    
    # Fiyat yukarıdaysa veya çok uzaktaysa puanı kötüleştir
    if dy < -5: 
        return 1000 + abs(dy)
    # Fiyat hemen altındaysa (0-150 birim) en iyi puanı ver
    return math.sqrt(dx**2 + dy**2)

st.title("🎯 Nokta Atışı Fiyat Eşleştirici (Gelişmiş Arama)")
st.write("Excel kodları PDF metni içinde aratılır ve en yakın fiyatla eşleştirilir.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="50+15")

if reference_excel and pdf_file:
    if st.button("🚀 TARAMAYI BAŞLAT"):
        # 1. Excel Verilerini Hazırla
        ref_df = pd.read_excel(reference_excel)
        excel_items = []
        for _, row in ref_df.iterrows():
            name = str(row.iloc[0]).strip()
            code = str(row.iloc[1]).strip()
            if code and code != 'nan':
                excel_items.append({"name": name, "code": code, "clean": clean_string(code)})

        if not excel_items:
            st.error("Excel dosyasında kod bulunamadı!")
            st.stop()

        # 2. PDF Analizi
        st.info(f"PDF analiz ediliyor ({len(excel_items)} ürün aranacak)...")
        results = []
        found_codes = set()
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress_bar = st.progress(0)
            total_pages = len(pdf.pages)

            for p_idx, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text: continue
                
                clean_page_text = clean_string(text)
                words = page.extract_words()
                
                # Sayfadaki tüm fiyatları ve koordinatlarını bul
                page_prices = []
                for match in price_regex.finditer(text):
                    p_str = match.group()
                    kuruş = p_str.split(',')[-1]
                    # Fiyatın koordinatını bul
                    for w in words:
                        if kuruş in w['text']:
                            page_prices.append({"text": p_str, "x": w['x0'], "y": w['top']})
                            break
                
                if not page_prices: continue

                # Excel'deki her ürünü bu sayfada ara
                for item in excel_items:
                    if item['clean'] in found_codes: continue # Zaten bulunduysa atla
                    
                    # Kod sayfada var mı? (Normal veya temizlenmiş haliyle)
                    if item['code'] in text or item['clean'] in clean_page_text:
                        # Kodun koordinatını bul (Sayfadaki kelimeleri birleştirip bak)
                        code_y = None
                        code_x = None
                        
                        # Kodu oluşturan kelimeyi/parçayı bul
                        for w in words:
                            if clean_string(w['text']) in item['clean'] or item['clean'] in clean_string(w['text']):
                                code_y = w['top']
                                code_x = w['x0']
                                break
                        
                        if code_y is not None:
                            # Bu koda en yakın fiyatı seç
                            best_price = min(page_prices, key=lambda p: get_distance_score(code_y, p['y'], code_x, p['x']))
                            
                            price_val = best_price['text']
                            net_val = calculate_discount(price_val, discount_input)
                            
                            results.append({
                                "Ürün İsmi": item['name'],
                                "Ürün Kodu": item['code'],
                                "Liste Fiyatı": price_val,
                                "Net Fiyat": net_val,
                                "Sayfa": p_idx + 1
                            })
                            found_codes.add(item['clean'])
                
                progress_bar.progress((p_idx + 1) / total_pages)

        # 3. Sonuçları Göster
        if results:
            res_df = pd.DataFrame(results)
            st.success(f"✅ {len(res_df)} ürün başarıyla eşleşti!")
            st.dataframe(res_df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 Excel İndir", output.getvalue(), "fiyat_listesi.xlsx")
            
            # Bulunamayanlar
            not_found = [i for i in excel_items if i['clean'] not in found_codes]
            if not_found:
                with st.expander(f"❌ Bulunamayan Ürünler ({len(not_found)})"):
                    st.table(pd.DataFrame(not_found)[['name', 'code']])
        else:
            st.error("Hiçbir ürün eşleşmedi. PDF'in metinlerinin okunabilir olduğundan emin olun.")
