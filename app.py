import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Hatasız 1-1 Fiyat Eşleştirici", layout="wide")

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

st.title("🎯 Nokta Atışı: Hatasız 1'e 1 Eşleştirici")
st.write("Fiyatları yukarıdan aşağıya sırayla kodlara atar. Aynı fiyatın iki ürüne yazılmasını engeller.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="20")

if reference_excel and pdf_file:
    if st.button("🚀 ANALİZİ BAŞLAT"):
        ref_df = pd.read_excel(reference_excel)
        excel_map = {}
        for _, row in ref_df.iterrows():
            name, code = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
            c_code = clean_string(code)
            if c_code:
                excel_map[c_code] = {"name": name, "orig": code}

        st.info("PDF taranıyor, lütfen bekleyin...")
        final_results = []
        found_codes_global = set()
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress = st.progress(0)
            
            for p_idx, page in enumerate(pdf.pages):
                w, h = page.width, page.height
                
                # Sayfayı 3 sütuna böl (Sol, Orta, Sağ)
                cols = [
                    (0, 0, w*0.33, h),
                    (w*0.33, 0, w*0.66, h),
                    (w*0.66, 0, w, h)
                ]
                
                for bbox in cols:
                    crop = page.within_bbox(bbox)
                    words = crop.extract_words()
                    text = crop.extract_text() or ""
                    
                    if not words or not text: continue
                    
                    # 1. Bu sütundaki TÜM KODLARI bul ve Y koordinatına göre diz
                    col_codes = []
                    for wrd in words:
                        w_clean = clean_string(wrd['text'])
                        # Excel'deki bir kodun parçası mı?
                        for k in excel_map.keys():
                            if k == w_clean or (len(w_clean) > 4 and w_clean in k):
                                if k not in found_codes_global:
                                    col_codes.append({"key": k, "top": wrd['top']})
                                    break
                    
                    # 2. Bu sütundaki TÜM FİYATLARI bul ve Y koordinatına göre diz
                    col_prices = []
                    for match in price_regex.finditer(text):
                        p_str = match.group()
                        kuruş = p_str.split(',')[-1]
                        for wrd in words:
                            if kuruş in wrd['text']:
                                col_prices.append({"text": p_str, "top": wrd['top']})
                                break
                    
                    # Y koordinatına (yukarıdan aşağıya) göre sırala
                    col_codes = sorted(col_codes, key=lambda x: x['top'])
                    col_prices = sorted(col_prices, key=lambda x: x['top'])
                    
                    # 3. SIRALI EŞLEŞTİRME (1. kod 1. fiyata, 2. kod 2. fiyata)
                    # Sadece o sütundaki sayıları eşleştirir
                    match_count = min(len(col_codes), len(col_prices))
                    for i in range(match_count):
                        c_item = col_codes[i]
                        p_item = col_prices[i]
                        
                        net_f = calculate_discount(p_item['text'], discount_input)
                        final_results.append({
                            "Ürün İsmi": excel_map[c_item['key']]['name'],
                            "Ürün Kodu": excel_map[c_item['key']]['orig'],
                            "Liste Fiyatı": p_item['text'],
                            "Net Fiyat": net_f,
                            "Sayfa": p_idx + 1
                        })
                        found_codes_global.add(c_item['key'])
                
                progress.progress((p_idx + 1) / len(pdf.pages))

        if final_results:
            res_df = pd.DataFrame(final_results).drop_duplicates(subset=['Ürün Kodu'])
            st.success(f"✅ {len(res_df)} ürün benzersiz fiyatlarla eşleşti!")
            st.dataframe(res_df, use_container_width=True)
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 Excel İndir", out.getvalue(), "hatasiz_liste.xlsx")
        else:
            st.error("Hiçbir ürün eşleşmedi.")
