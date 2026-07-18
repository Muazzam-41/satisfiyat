import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="PDF Fiyat Düzenleyici", layout="centered")

st.title("📊 PDF Fiyat & İskonto Hesaplayıcı")
st.write("PDF'i yükleyin, iskontoyu girin ve Excel olarak indirin.")

# Kullanıcı Girişleri
uploaded_file = st.file_uploader("PDF Dosyasını Buraya Yükleyin", type="pdf")
discount = st.number_input("Uygulanacak İskonto Oranı (%)", min_value=0.0, max_value=100.0, value=20.0)

def extract_data(pdf_file, discount_rate):
    extracted_rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # Boş olmayan satırları kontrol et
                    row_content = [str(cell) for cell in row if cell]
                    line_text = " ".join(row_content)
                    
                    # Satırda fiyat olabilecek sayıları ara (Örn: 1.250,50 veya 450.00)
                    # Bu regex Türkiye'deki fiyat formatlarına uygundur
                    prices = re.findall(r'\d+(?:\.\d+)*(?:,\d+)?', line_text)
                    
                    if prices:
                        # Genelde satırdaki son sayı fiyattır
                        raw_price = prices[-1]
                        try:
                            # Sayıyı float formatına çevir
                            clean_price = raw_price.replace('.', '').replace(',', '.')
                            price_val = float(clean_price)
                            
                            new_price = price_val * (1 - discount_rate / 100)
                            
                            extracted_rows.append({
                                "Ürün/Satır Bilgisi": line_text[:50] + "...",
                                "Orijinal Fiyat": price_val,
                                "İskontolu Fiyat": round(new_price, 2)
                            })
                        except:
                            continue
    return extracted_rows

if uploaded_file is not None:
    with st.spinner('PDF İşleniyor...'):
        results = extract_data(uploaded_file, discount)
        
        if results:
            df = pd.DataFrame(results)
            st.success(f"{len(df)} adet fiyat bulundu!")
            st.dataframe(df) # Ekranda göster
            
            # Excel'e Dönüştürme
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Fiyatlar')
            
            st.download_button(
                label="📥 İskontolu Listeyi Excel Olarak İndir",
                data=output.getvalue(),
                file_name="iskontolu_fiyat_listesi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("PDF içinde fiyat formatına uygun veri bulunamadı. Lütfen metin tabanlı bir PDF yükleyin.")

st.info("Not: Bu araç PDF içindeki tabloları ve sayısal verileri tarar. Eğer PDF bir 'resim/fotoğraf' ise (taranamıyorsa) çalışmayabilir.")