import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# Sayfa Ayarları
st.set_page_config(page_title="Pro Katalog Fiyat Botu", layout="wide")

# Hatalı olan kısım burasıydı, düzeltildi: unsafe_allow_html=True
st.markdown("""
    <style>
    .main { background-color: #f5f5f5; }
    .stButton>button { width: 100%; background-color: #28a745; color: white; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Profesyonel Katalog Fiyat İşleyici")
st.info("PDF'deki büyük başlıkları (Ürün İsmi) yakalar ve altındaki kod/fiyatlarla eşleştirir.")

# Kullanıcı Girişleri
col1, col2, col3 = st.columns([2, 1, 1])
with col1:
    uploaded_file = st.file_uploader("PDF Katalog Dosyasını Yükleyin", type="pdf")
with col2:
    discount_input = st.text_input("İskonto (Örn: 50+15+5)", value="50+10")
with col3:
    header_font_size = st.slider("Başlık Font Hassasiyeti", 8.0, 25.0, 12.0)
    st.caption("Başlıkları kaçırıyorsa bu değeri düşürün.")

def calculate_chain_discount(price_str, discount_str):
    """Zincir iskontoyu hesaplar (Örn: 1.250,50 için 50+10)"""
    try:
        # Fiyat metnini sayıya çevir (1.250,50 -> 1250.50)
        clean_p = price_str.replace('.', '').replace(',', '.')
        current_price = float(clean_p)
        
        if not discount_str:
            return round(current_price, 2)
            
        discounts = [float(d.strip()) for d in discount_str.split('+') if d.strip()]
        for d in discounts:
            current_price = current_price * (1 - d / 100)
        return round(current_price, 2)
    except:
        return 0

def process_pdf(pdf_file, discount_str, threshold):
    extracted_data = []
    current_header = "Bilinmeyen Ürün Grubu"
    
    # Fiyat ve Kod yakalamak için Regex (Resimlerdeki formatlara göre)
    price_pattern = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}') # Örn: 13.890,48
    
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
            
            # Satırları yukarıdan aşağıya tara
            for y in sorted(lines.keys()):
                line_words = lines[y]
                line_text = " ".join([w['text'] for w in line_words])
                max_font_in_line = max([w['size'] for w in line_words])
                
                # 1. BAŞLIK YAKALAMA (Büyük font ve içinde fiyat olmayan satır)
                if max_font_in_line >= threshold:
                    if not price_pattern.search(line_text):
                        current_header = line_text
                        continue

                # 2. SATIR İŞLEME (Fiyat içeren satırları yakala)
                found_price = price_pattern.search(line_text)
                if found_price:
                    price_val_str = found_price.group()
                    
                    # Satırdaki ilk kelimeyi 'Kod' olarak al (VDS, ARM, NRA ile başladığı için)
                    code_str = line_words[0]['text']
                    
                    # Eğer kod ve fiyat aynı şeyse (kod bulunamadıysa)
                    if code_str == price_val_str:
                        code_str = "KOD BELİRTİLMEMİŞ"

                    net_fiyat = calculate_chain_discount(price_val_str, discount_str)
                    
                    extracted_data.append({
                        "Ürün İsmi (Başlık)": current_header,
                        "Ürün Kodu": code_str,
                        "Liste Fiyatı": price_val_str,
                        "İskontolu Fiyat": net_fiyat,
                        "İskonto Oranı": discount_str
                    })

    return extracted_data

if uploaded_file:
    with st.spinner('Katalog analiz ediliyor...'):
        results = process_pdf(uploaded_file, discount_input, header_font_size)
        
        if results:
            df = pd.DataFrame(results)
            st.success(f"✅ {len(df)} adet ürün başarıyla listelendi!")
            
            # Tabloyu göster
            st.dataframe(df, use_container_width=True)
            
            # Excel Hazırlama
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Fiyatlar')
            
            st.download_button(
                label="📥 Excel Dosyasını İndir",
                data=output.getvalue(),
                file_name="iskontolu_fiyatlar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Veri bulunamadı. Lütfen 'Başlık Font Hassasiyeti' değerini düşürerek tekrar deneyin.")
