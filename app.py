import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Pro Katalog Fiyat Botu", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f5f5f5; }
    .stButton>button { width: 100%; background-color: #007bff; color: white; }
    </style>
    """, unsafe_allow_stdio=True)

st.title("🚀 Gelişmiş Katalog Fiyat İşleyici")
st.write("PDF'deki büyük başlıkları, ürün kodlarını ve fiyatları akıllıca eşleştirir.")

# Kullanıcı Girişleri
col1, col2, col3 = st.columns([2, 1, 1])
with col1:
    uploaded_file = st.file_uploader("PDF Katalog Dosyasını Yükleyin", type="pdf")
with col2:
    discount_input = st.text_input("İskonto (Örn: 50+15+5)", value="50+10")
with col3:
    header_font_size = st.slider("Başlık Yakalama Hassasiyeti", 10.0, 25.0, 13.0)
    st.caption("Başlıkları bulamıyorsa bu değeri düşürün.")

def calculate_chain_discount(price, discount_str):
    """50+15 gibi zincir iskontoları hesaplar."""
    try:
        clean_p = str(price).replace('.', '').replace(',', '.')
        current_price = float(clean_p)
        discounts = [float(d.strip()) for d in discount_str.split('+') if d.strip()]
        for d in discounts:
            current_price = current_price * (1 - d / 100)
        return round(current_price, 2)
    except:
        return 0

def process_catalog_pdf(pdf_file, discount_str, threshold):
    extracted_data = []
    current_header = "Bilinmeyen Ürün Grubu"
    
    # Fiyat formatı regex (Örn: 14.223,53 veya 27,96)
    price_pattern = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')
    # Kod formatı regex (Örn: VDS 753023, ARM 717254, NRA 004866)
    code_pattern = re.compile(r'[A-Z]{1,4}\s?\d+')

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Sayfadaki her kelimeyi koordinat ve font bilgisiyle al
            words = page.extract_words(extra_attrs=["size", "fontname"])
            
            # Kelimeleri satırlara göre grupla
            lines = {}
            for w in words:
                y = round(w['top'])
                if y not in lines: lines[y] = []
                lines[y].append(w)
            
            for y in sorted(lines.keys()):
                line_words = lines[y]
                line_text = " ".join([w['text'] for w in line_words])
                max_font_in_line = max([w['size'] for w in line_words])
                
                # 1. ADIM: BAŞLIK YAKALAMA (Büyük ve kalın yazı)
                if max_font_in_line >= threshold:
                    # Eğer satır tamamen sayı veya fiyat değilse başlık kabul et
                    if not price_pattern.search(line_text):
                        current_header = line_text
                        continue

                # 2. ADIM: TABLO / SATIR YAKALAMA
                found_price = price_pattern.search(line_text)
                if found_price:
                    price_str = found_price.group()
                    # Aynı satırda kod ara
                    found_code = code_pattern.search(line_text)
                    
                    if found_code:
                        code_str = found_code.group()
                    else:
                        code_str = "Kod Bulunamadı"
                    
                    net_fiyat = calculate_chain_discount(price_str, discount_str)
                    
                    extracted_data.append({
                        "Ürün Grubu": current_header,
                        "Ürün Kodu": code_str,
                        "Liste Fiyatı": price_str,
                        "Net Fiyat": net_fiyat,
                        "Uygulanan İskonto": discount_str
                    })

    return extracted_data

if uploaded_file:
    with st.spinner('Katalog taranıyor, bu işlem PDF boyutuna göre biraz sürebilir...'):
        results = process_catalog_pdf(uploaded_file, discount_input, header_font_size)
        
        if results:
            df = pd.DataFrame(results)
            st.success(f"İşlem Tamam! {len(df)} adet ürün bulundu.")
            
            # Tabloyu göster
            st.dataframe(df, use_container_width=True)
            
            # Excel Çıktısı
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Fiyat Listesi')
            
            st.download_button(
                label="📥 Excel Dosyasını İndir",
                data=output.getvalue(),
                file_name="iskontolu_fiyat_listesi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Ürün veya fiyat bulunamadı. Lütfen 'Başlık Yakalama Hassasiyeti' ayarını değiştirin.")

st.info("""
**Nasıl Çalışır?**
1. Sayfanın üstündeki büyük yazıyı 'Ürün İsmi' olarak belirler.
2. Altındaki tablolarda 'ARM 123' gibi kodları ve '1.250,00' gibi fiyatları arar.
3. Bulduğu her fiyatı, en son gördüğü büyük başlığın altına yazar.
""")
