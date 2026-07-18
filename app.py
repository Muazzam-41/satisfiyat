import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Çift Sütun Katalog İşleyici", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #d9534f; color: white; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Profesyonel Katalog İşleyici (Yan Yana Ürün Desteği)")
st.info("Bu sürüm, sayfayı iki sütuna bölerek yan yana duran ürünleri birbirine karıştırmadan çeker.")

# Girdiler
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("PDF Katalog Dosyasını Yükleyin", type="pdf")
with col2:
    discount_input = st.text_input("İskonto Oranları (Örn: 50+15)", value="50+10")

def calculate_net_price(list_price_str, discount_str):
    try:
        clean_price = list_price_str.replace('.', '').replace(',', '.')
        price = float(clean_price)
        if not discount_str: return round(price, 2)
        discounts = [float(d.strip()) for d in discount_str.split('+') if d.strip()]
        for d in discounts:
            price = price * (1 - d / 100)
        return round(price, 2)
    except:
        return 0

def process_region(region_words, discount_str):
    """Belirli bir bölgedeki (sol veya sağ sütun) kelimeleri işler."""
    region_data = []
    current_title = "Bilinmeyen Ürün"
    
    price_regex = r'\d{1,3}(?:\.\d{3})*,\d{2}'
    code_regex = r'\b[A-Z]{1,4}\s?\d{4,}\b'
    ignored_words = ["FİYAT", "KOLİ", "ADET", "KUTU", "P.ADET", "YENİ ÜRÜN"]

    # Kelimeleri satırlara göre grupla
    lines = {}
    for w in region_words:
        y = round(w['top'])
        if y not in lines: lines[y] = []
        lines[y].append(w)
    
    for y in sorted(lines.keys()):
        line_words = lines[y]
        line_text = " ".join([w['text'] for w in line_words]).strip()
        
        # 1. Başlık Tespiti: Tamamı büyük harfse ve kod/fiyat içermiyorsa
        has_price = re.search(price_regex, line_text)
        has_code = re.search(code_regex, line_text)
        
        if line_text.isupper() and not has_price and not has_code:
            if not any(ig in line_text for ig in ignored_words) and len(line_text) > 3:
                current_title = line_text
                continue

        # 2. Kod ve Fiyat Tespiti
        if has_price:
            price_found = has_price.group()
            code_found = has_code.group() if has_code else "KOD BULUNAMADI"
            
            # Eğer kod o satırda yoksa, bu satırdaki diğer kelimelere bak
            if code_found == "KOD BULUNAMADI":
                # Bazen kod fiyatın hemen yanındaki kelimedir ama regex tam oturmamıştır
                for word in line_words:
                    if any(char.isdigit() for char in word['text']) and any(char.isalpha() for char in word['text']):
                        code_found = word['text']
                        break

            net_fiyat = calculate_net_price(price_found, discount_str)
            region_data.append({
                "Ürün İsmi": current_title,
                "Ürün Kodu": code_found,
                "Liste Fiyatı": price_found,
                "İskontolu Fiyat": net_fiyat
            })
            
    return region_data

if uploaded_file:
    with st.spinner('Sayfalar iki sütun olarak analiz ediliyor...'):
        all_data = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                width = page.width
                height = page.height
                
                # Sayfayı SOL ve SAĞ olarak ikiye bölüyoruz
                left_bbox = (0, 0, width / 2, height)
                right_bbox = (width / 2, 0, width, height)
                
                left_words = page.within_bbox(left_bbox).extract_words()
                right_words = page.within_bbox(right_bbox).extract_words()
                
                # Önce sol sütunu, sonra sağ sütunu işle
                all_data.extend(process_region(left_words, discount_input))
                all_data.extend(process_region(right_words, discount_input))
        
        if all_data:
            df = pd.DataFrame(all_data)
            st.success(f"✅ Toplam {len(df)} ürün başarıyla ayrıştırıldı!")
            st.dataframe(df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Excel Dosyasını İndir",
                data=output.getvalue(),
                file_name="cift_sutun_liste.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Veri çekilemedi. PDF formatını kontrol edin.")

st.divider()
st.caption("Bu sürüm, yan yana (çift sütun) yerleşimli kataloglar için özel olarak optimize edilmiştir.")
