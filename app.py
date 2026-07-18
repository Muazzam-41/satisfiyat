import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Katalog Fiyat Botu", layout="wide")

st.title("📂 Katalogdan Fiyat ve Kod Çekici")
st.write("Büyük harfli başlıkları 'Ürün İsmi' olarak alır ve altındaki kod/fiyatlarla eşleştirir.")

# Kullanıcı Girişleri
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("PDF Katalog Dosyasını Yükleyin", type="pdf")
with col2:
    discount_input = st.text_input("İskonto (Örn: 50+15)", value="50+10")
    # Font büyüklüğü eşiği (Hangi yazıların başlık sayılacağını belirler)
    font_threshold = st.slider("Başlık Font Büyüklüğü Eşiği", 10, 20, 12)

def calculate_chain_discount(price, discount_str):
    try:
        discounts = [float(d.strip()) for d in discount_str.split('+')]
        for d in discounts:
            price = price * (1 - d / 100)
        return round(price, 2)
    except:
        return price

def is_price(text):
    """Metnin bir fiyat olup olmadığını kontrol eder."""
    clean = re.sub(r'[^\d,.]', '', str(text))
    return bool(re.search(r'\d', clean))

def process_catalog(pdf_file, discount_str, threshold):
    data = []
    current_product_name = "Tanımsız Ürün"
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Sayfadaki tüm kelimeleri ve özelliklerini (font büyüklüğü dahil) al
            words = page.extract_words(extra_attrs=["size", "fontname"])
            
            # Kelimeleri satır satır grupla (y koordinatına göre)
            lines = {}
            for word in words:
                y = round(word['top'])
                if y not in lines: lines[y] = []
                lines[y].append(word)
            
            # Satırları yukarıdan aşağıya işle
            for y in sorted(lines.keys()):
                line_words = lines[y]
                line_text = " ".join([w['text'] for w in line_words])
                avg_font_size = sum([w['size'] for w in line_words]) / len(line_words)
                
                # MANTIK 1: Eğer font büyükse ve harfler büyükse bu ÜRÜN İSMİDİR
                if avg_font_size >= threshold and line_text.isupper():
                    current_product_name = line_text
                    continue

                # MANTIK 2: Satırda hem kod hem fiyat araması yap
                # (Genelde kodlar harf+rakam, fiyatlar ise rakam+virgül içerir)
                # Bu kısım basit bir tablo ayırıcı gibi çalışır
                if len(line_words) >= 2:
                    potential_price = line_words[-1]['text']
                    potential_code = line_words[0]['text']
                    
                    # Fiyat temizleme ve kontrol
                    clean_p = re.sub(r'[^\d,.]', '', potential_price).replace(',', '.')
                    try:
                        price_val = float(clean_p)
                        if price_val > 0:
                            net_price = calculate_chain_discount(price_val, discount_str)
                            data.append({
                                "Ürün Grubu (Başlık)": current_product_name,
                                "Ürün Kodu": potential_code,
                                "Liste Fiyatı": price_val,
                                "Net Fiyat": net_price
                            })
                    except:
                        continue
    return data

if uploaded_file:
    with st.spinner('Katalog analiz ediliyor...'):
        extracted_data = process_catalog(uploaded_file, discount_input, font_threshold)
        
        if extracted_data:
            df = pd.DataFrame(extracted_data)
            st.success(f"{len(df)} kalem ürün/fiyat bulundu.")
            st.dataframe(df, use_container_width=True)
            
            # Excel İndirme
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 İskontolu Listeyi İndir",
                data=output.getvalue(),
                file_name="katalog_fiyat_listesi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Belirlediğiniz font eşiğine göre ürün ismi veya fiyat bulunamadı. Eşiği düşürmeyi deneyin.")
