import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import math

st.set_page_config(page_title="Nokta Atışı: Kesin Çözüm", layout="wide")

# 1. TEMİZLEME VE HESAPLAMA
def clean_string(s):
    if not s or s == 'nan': return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    try:
        num_str = re.sub(r'[^\d,]', '', price_str).replace(',', '.')
        val = float(num_str)
        if not disc_str or disc_str == "0": return round(val, 2)
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

# 2. ARAYÜZ
st.title("🎯 Nokta Atışı: Kesin Çözüm Botu")
st.write("Kodlar parçalanmış olsa bile PDF'in derinliklerinden bulur ve fiyatıyla eşleştirir.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="20")

# 3. ANA İŞLEM
if reference_excel and pdf_file:
    if st.button("🚀 DERİN TARAMAYI BAŞLAT"):
        # Excel Hazırlığı
        ref_df = pd.read_excel(reference_excel)
        excel_items = []
        for _, row in ref_df.iterrows():
            name, code = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
            if code and code != 'nan' and len(code) > 2:
                excel_items.append({"name": name, "orig": code, "clean": clean_string(code)})

        if not excel_items:
            st.error("Excel'de taranacak kod bulunamadı!")
            st.stop()

        st.info(f"Analiz ediliyor: {len(excel_items)} ürün aranıyor...")
        
        results = []
        found_codes = set()
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress_bar = st.progress(0)
            
            for p_idx, page in enumerate(pdf.pages):
                # Sayfadaki tüm kelimeleri ve fiyatları çıkar
                words = page.extract_words()
                text = page.extract_text()
                
                if not text: continue
                
                # Sayfadaki tüm fiyatları ve konumlarını bul
                page_prices = []
                for match in price_regex.finditer(text):
                    p_str = match.group()
                    kuruş = p_str.split(',')[-1]
                    # Fiyatın koordinatını en yakın kelimeden bul
                    for w in words:
                        if kuruş in w['text']:
                            page_prices.append({"text": p_str, "top": w['top'], "x0": w['x0']})
                            break

                # Excel kodlarını sayfa metni içinde ara
                for target in excel_items:
                    if target['clean'] in found_codes: continue
                    
                    # Kodu metin içinde bul (Normal veya boşluksuz haliyle)
                    # re.escape kullanarak özel karakterleri (nokta, tire) güvenli hale getiriyoruz
                    pattern = re.compile(re.escape(target['orig']), re.IGNORECASE)
                    match_in_text = pattern.search(text)
                    
                    if not match_in_text and len(target['clean']) > 4:
                        # Eğer tam kod yoksa, boşluksuz halini ara
                        clean_page_text = clean_string(text)
                        if target['clean'] in clean_page_text:
                            # Koordinatı yaklaşık olarak kelime listesinden çek
                            for w in words:
                                if clean_string(w['text']) in target['clean']:
                                    code_top = w['top']
                                    code_x = w['x0']
                                    break
                            else: continue
                        else: continue
                    elif match_in_text:
                        # Koordinatı bul
                        first_word_of_code = target['orig'].split()[0]
                        for w in words:
                            if first_word_of_code in w['text']:
                                code_top = w['top']
                                code_x = w['x0']
                                break
                        else: continue
                    else: continue

                    # Kod bulundu, şimdi en yakın fiyatı (sağda veya altta) bul
                    best_price = None
                    min_score = 999999
                    
                    for p in page_prices:
                        dy = p['top'] - code_top
                        dx = p['x0'] - code_x
                        
                        # MIKNATIS MANTIĞI: Sağda (satır sonu) veya Altta (resim altı)
                        if abs(dy) < 12 and dx > 0: # Aynı satır
                            score = dx * 0.5
                        elif dy > 0 and abs(dx) < 200: # Alt satırlar
                            score = dy
                        else:
                            score = 999999
                        
                        if score < min_score:
                            min_score = score
                            best_price = p['text']
                    
                    if best_price and min_score < 1000:
                        net_val = calculate_discount(best_price, discount_input)
                        results.append({
                            "Ürün İsmi": target['name'],
                            "Ürün Kodu": target['orig'],
                            "Liste Fiyatı": best_price,
                            "Net Fiyat": net_val,
                            "Sayfa": p_idx + 1
                        })
                        found_codes.add(target['clean'])

                progress_bar.progress((p_idx + 1) / len(pdf.pages))

        # 4. SONUÇLAR
        if results:
            res_df = pd.DataFrame(results).drop_duplicates(subset=['Ürün Kodu'])
            st.success(f"✅ Başarı! {len(res_df)} / {len(excel_items)} ürün eşleşti.")
            st.dataframe(res_df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            
            st.download_button("📥 Excel Olarak İndir", output.getvalue(), "fiyat_listesi.xlsx")
        else:
            st.error("Hala eşleşme bulunamadı. Lütfen PDF'in 'taranmış bir resim' (OCR gerektiren) olup olmadığını kontrol edin.")
            st.info("Eğer PDF'de yazıları farenizle seçemiyorsanız, bu bir resimdir ve bu yazılımın okuyabilmesi için OCR destekli bir sürüm gerekir.")
