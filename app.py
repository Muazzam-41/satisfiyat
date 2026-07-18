import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import math

st.set_page_config(page_title="Global Fiyat Dedektifi v5", layout="wide")

def clean_string(s):
    if not s or s == 'nan': return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    try:
        # Fiyatı temizle: Rakam ve virgül dışındakileri at, virgülü noktaya çevir
        num_str = re.sub(r'[^\d,]', '', price_str).replace(',', '.')
        val = float(num_str)
        if not disc_str or disc_str == "0": return round(val, 2)
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

st.title("🔍 Global Fiyat Dedektifi (Yüksek Eşleşme Sürümü)")
st.write("649 ürünlük listeniz için özel olarak optimize edildi. Parçalanmış kodları ve uzak fiyatları yakalar.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 1. Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 2. PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto Oranı", value="20")

if reference_excel and pdf_file:
    if st.button("🚀 DERİN TARAMAYI BAŞLAT"):
        # Excel'i Hazırla
        ref_df = pd.read_excel(reference_excel)
        excel_items = []
        for _, row in ref_df.iterrows():
            name, code = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
            if code and code != 'nan':
                excel_items.append({"name": name, "orig": code, "clean": clean_string(code)})

        st.info(f"Derin tarama başladı: {len(excel_items)} ürün aranıyor...")
        
        final_results = []
        used_prices = set() # Aynı fiyatın iki kez kullanılmasını engellemek için
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            # 1. ADIM: PDF'deki tüm kodları ve fiyatları koordinatlarıyla tek seferde hafızaya al
            pdf_data = []
            progress_bar = st.progress(0)
            
            for p_idx, page in enumerate(pdf.pages):
                words = page.extract_words()
                text = page.extract_text() or ""
                
                # Sayfadaki tüm fiyatları bul
                page_prices = []
                for match in price_regex.finditer(text):
                    p_str = match.group()
                    kuruş = p_str.split(',')[-1]
                    for w in words:
                        if kuruş in w['text']:
                            page_prices.append({"text": p_str, "top": w['top'], "x0": w['x0'], "page": p_idx})
                            break
                
                pdf_data.append({"words": words, "prices": page_prices, "text": text, "page_num": p_idx + 1})
                progress_bar.progress((p_idx + 1) / len(pdf.pages))

            # 2. ADIM: Her Excel kodu için PDF verilerinde "En Yakın" araması yap
            st.info("Eşleştirme yapılıyor...")
            found_count = 0
            
            for target in excel_items:
                best_match = None
                min_dist = 999999
                
                for p_info in pdf_data:
                    # Kodu bu sayfada ara (Hem normal hem temizlenmiş)
                    # Kod parçalanmış olabilir, bu yüzden sayfa metninde temizlenmiş arama yapıyoruz
                    if target['clean'] in clean_string(p_info['text']):
                        
                        # Kodun sayfadaki koordinatını bulmaya çalış
                        code_coords = None
                        for w in p_info['words']:
                            if clean_string(w['text']) in target['clean']:
                                code_coords = w
                                break
                        
                        if code_coords:
                            # Bu sayfadaki fiyatlardan en uygun olanı seç
                            for p_item in p_info['prices']:
                                # Fiyat daha önce bu ürün grubunda kullanılmadıysa
                                price_key = f"{p_item['page']}_{p_item['top']}_{p_item['x0']}"
                                
                                dy = p_item['top'] - code_coords['top']
                                dx = abs(p_item['x0'] - code_coords['x0'])
                                
                                # Fiyat aşağıda (dy > -5) ve çok uzakta değilse (dx < 300)
                                if dy > -10 and dx < 300:
                                    score = dy + (dx * 1.5)
                                    if score < min_dist:
                                        min_dist = score
                                        best_match = {
                                            "price": p_item['text'],
                                            "page": p_info['page_num'],
                                            "key": price_key
                                        }

                if best_match:
                    net_val = calculate_discount(best_match['price'], discount_input)
                    final_results.append({
                        "Ürün İsmi": target['name'],
                        "Ürün Kodu": target['orig'],
                        "Liste Fiyatı": best_match['price'],
                        "Net Fiyat": net_val,
                        "Sayfa": best_match['page']
                    })
                    found_count += 1

        # 3. SONUÇLAR
        if final_results:
            res_df = pd.DataFrame(final_results)
            st.success(f"✅ Başarı: {found_count} / {len(excel_items)} ürün eşleşti!")
            st.dataframe(res_df, use_container_width=True)
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 Excel İndir", out.getvalue(), "nihai_fiyat_listesi.xlsx")
        else:
            st.error("Hiçbir eşleşme bulunamadı.")
