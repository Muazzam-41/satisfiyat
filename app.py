import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Katalog Fiyat Botu Pro", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("📦 Katalogdan Akıllı Veri Çekme & İskonto")
st.write("PDF Kataloglarını analiz eder; Ürün İsmi, Kod ve Fiyatı eşleştirerek Excel verir.")

# Girdiler
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("PDF Katalog Dosyasını Buraya Yükleyin", type="pdf")
with col2:
    discount_input = st.text_input("İskonto Oranları (Örn: 50+15+10)", value="50+10")

def calculate_net_price(list_price_str, discount_str):
    """Metin halindeki fiyatı sayıya çevirir ve zincir iskontoyu uygular."""
    try:
        # 13.890,48 -> 13890.48
        clean_price = list_price_str.replace('.', '').replace(',', '.')
        price = float(clean_price)
        
        if not discount_str:
            return round(price, 2)
            
        discounts = [float(d.strip()) for d in discount_str.split('+') if d.strip()]
        for d in discounts:
            price = price * (1 - d / 100)
        return round(price, 2)
    except:
        return 0

def extract_data_from_catalog(pdf_file, discount_str):
    data = []
    current_product_name = "Bilinmeyen Ürün"
    
    # REGEX TANIMLARI (Senin verdiğin örneklere göre optimize edildi)
    # Fiyat: 1.425,49 veya 27,96 formatı
    price_regex = r'\d{1,3}(?:\.\d{3})*,\d{2}'
    # Kod: BHV 104528, T 33530222, ARM 717254, NRA 004866, VDS 753023 formatları
    code_regex = r'\b[A-Z]{1,4}\s?\d{4,}\b'

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            words = page.extract_words(extra_attrs=["fontname", "size"])
            
            # Kelimeleri satırlara grupla
            lines = {}
            for w in words:
                y = round(w['top'])
                if y not in lines: lines[y] = []
                lines[y].append(w)
            
            for y in sorted(lines.keys()):
                line_words = lines[y]
                line_text = " ".join([w['text'] for w in line_words]).strip()
                
                # 1. MANTIK: ÜRÜN İSMİNİ BULMA
                # Eğer satır tamamen büyük harfse ve kod/fiyat içermiyorsa Ürün İsmi'dir.
                # 'KOLİ ADET', 'FİYAT' gibi kelimeleri de eleyelim.
                is_mostly_uppercase = line_text.isupper()
                has_price = re.search(price_regex, line_text)
                has_code = re.search(code_regex, line_text)
                ignored_words = ["FİYAT", "KOLİ", "ADET", "KUTU", "P.ADET", "YENİ ÜRÜN"]

                if is_mostly_uppercase and not has_price and not has_code:
                    if not any(ignored in line_text for ignored in ignored_words) and len(line_text) > 3:
                        current_product_name = line_text
                        continue

                # 2. MANTIK: KOD VE FİYATI BULMA
                # Eğer aynı satırda hem kod hem fiyat varsa
                if has_price:
                    price_found = has_price.group()
                    
                    # Eğer satırda kod da varsa onu al, yoksa bir üstteki kodları ara
                    if has_code:
                        code_found = has_code.group()
                    else:
                        # Kod bazen fiyatın üstündeki satırda olabilir
                        code_found = "Kod Bulunamadı"
                    
                    net_fiyat = calculate_net_price(price_found, discount_str)
                    
                    data.append({
                        "Ürün İsmi": current_product_name,
                        "Ürün Kodu": code_found,
                        "Liste Fiyatı": price_found,
                        "İskontolu Fiyat": net_fiyat,
                        "İskonto Yapısı": discount_str
                    })

    return data

if uploaded_file:
    with st.spinner('Katalog analiz ediliyor, lütfen bekleyin...'):
        final_results = extract_data_from_catalog(uploaded_file, discount_input)
        
        if final_results:
            df = pd.DataFrame(final_results)
            
            # Görselleştirme
            st.success(f"✅ Toplam {len(df)} ürün ve fiyat eşleşmesi bulundu!")
            st.dataframe(df, use_container_width=True)
            
            # Excel Hazırlama
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Fiyat Listesi')
            
            st.download_button(
                label="📥 Hazırlanan Excel Dosyasını İndir",
                data=output.getvalue(),
                file_name="iskontolu_katalog_listesi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("PDF içerisinde ürün ismi, kod veya fiyat yapısı tespit edilemedi.")

st.divider()
st.caption("Not: Bu araç BHV, T, ARM, VDS gibi kod yapılarını ve büyük harfli ürün başlıklarını otomatik tanır.")
