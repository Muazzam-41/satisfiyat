import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import math

# Sayfa Yapılandırması
st.set_page_config(page_title="Hatasız Benzersiz Fiyat Botu", layout="wide")

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

def get_match_score(code_w, price_w):
    """
    Fiyatın kodun altında ve aynı sütunda olmasını ölçer.
    Fiyat yukarıdaysa veya çok sağda/soldaysa elenir.
    """
    dy = price_w['top'] - code_w['top']
    dx = abs(price_w['x0'] - code_w['x0'])
    
    # Kural 1: Fiyat kodun yukarısındaysa imkansız (dy negatifse elenir)
    if dy < -5: return 999999
    
    # Kural 2: Fiyat kodun çok uzağındaysa (farklı sütunsa) elenir
    if dx > 180: return 999999
    
    # Kural 3: Aynı dikey hizada (sütunda) olmaya büyük öncelik ver
    return dy + (dx * 2)

st.title("🎯 Hatasız Benzersiz Fiyat Eşleştirici")
st.write("Her fiyat sadece bir ürüne atanır. Aynı fiyatın iki üründe çıkması engellenmiştir.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="20")

if reference_excel and pdf_file:
    if st.button("🚀 BENZERSİZ TARAMAYI BAŞLAT"):
        # Excel hazırlığı
        ref_df = pd.read_excel(reference_excel)
        excel_items = []
        for _, row in ref_df.iterrows():
            name, code = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
            c_code = clean_string(code)
            if c_code:
                excel_items.append({"name": name, "orig": code, "clean": c_code})

        st.info(f"{len(excel_items)} ürün analiz ediliyor...")
        
        found_results = []
        found_codes_global = set()
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress = st.progress(0)
            
            for p_idx, page in enumerate(pdf.pages):
                words = page.extract_words()
                text = page.extract_text() or ""
                
                # Sayfadaki tüm fiyatları koordinatlarıyla bul
                available_prices = []
                for match in price_regex.finditer(text):
                    p_str = match.group()
                    kuruş = p_str.split(',')[-1]
                    for w in words:
                        if kuruş in w['text']:
                            available_prices.append({"text": p_str, "top": w['top'], "x0": w['x0'], "id": f"{p_idx}_{w['top']}_{w['x0']}"})
                            break
                
                if not available_prices: continue

                # Sayfadaki kodları bul ve yukarıdan aşağıya sırala (Daha kararlı eşleşme için)
                page_codes = []
                for w in words:
                    w_clean = clean_string(w['text'])
                    for item in excel_items:
                        if item['clean'] == w_clean or (len(w_clean) > 5 and w_clean in item['clean']):
                            if item['clean'] not in found_codes_global:
                                page_codes.append({"item": item, "coords": w})
                
                # Sayfadaki kodları yukarıdan aşağıya tara
                page_codes = sorted(page_codes, key=lambda x: x['coords']['top'])

                for c_box in page_codes:
                    best_price_idx = -1
                    min_score = 999998
                    
                    for i, p_box in enumerate(available_prices):
                        score = get_match_score(c_box['coords'], p_box)
                        if score < min_score:
                            min_score = score
                            best_price_idx = i
                    
                    # Eğer uygun bir fiyat bulunduysa
                    if best_price_idx != -1 and min_score < 1000:
                        selected_price = available_prices.pop(best_price_idx) # Fiyatı listeden SİL (Bir daha kullanılamaz)
                        
                        net_f = calculate_discount(selected_price['text'], discount_input)
                        found_results.append({
                            "Ürün İsmi": c_box['item']['name'],
                            "Ürün Kodu": c_box['item']['orig'],
                            "Liste Fiyatı": selected_price['text'],
                            "Net Fiyat": net_f,
                            "Sayfa": p_idx + 1
                        })
                        found_codes_global.add(c_box['item']['clean'])

                progress.progress((p_idx + 1) / len(pdf.pages))

        if found_results:
            res_df = pd.DataFrame(found_results)
            st.success(f"✅ {len(res_df)} Ürün benzersiz fiyatlarla eşleşti!")
            st.dataframe(res_df, use_container_width=True)
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 Excel İndir", out.getvalue(), "hatasiz_liste.xlsx")
            
            not_found = [i for i in excel_items if i['clean'] not in found_codes_global]
            if not_found:
                with st.expander(f"❌ Bulunamayanlar ({len(not_found)})"):
                    st.table(pd.DataFrame(not_found)[['name', 'orig']])
