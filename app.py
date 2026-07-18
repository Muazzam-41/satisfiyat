import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Gelişmiş Fiyat Botu", layout="wide")

st.title("📊 Profesyonel PDF Fiyat İşleyici")
st.write("PDF'den Ürün Kodu, Adı ve Fiyatı ayırır; çoklu iskonto (Örn: 50+15) uygular.")

# Kullanıcı Girişleri
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("PDF Dosyasını Yükleyin", type="pdf")
with col2:
    discount_input = st.text_input("İskonto Oranları (Arada + kullanarak)", value="50+15")
    st.caption("Örnek: 50+15 yazarsanız önce %50, sonra kalan tutar üzerinden %15 düşer.")

def calculate_chain_discount(original_price, discount_str):
    """50+15+5 gibi verileri sırasıyla fiyata uygular."""
    try:
        discounts = [float(d.strip()) for d in discount_str.split('+')]
        current_price = original_price
        for d in discounts:
            current_price = current_price * (1 - d / 100)
        return round(current_price, 2)
    except:
        return original_price

def clean_price(text):
    """Metin içindeki fiyatı temizleyip sayıya çevirir."""
    if not text: return None
    # Nokta ve virgül karmaşasını çözer (Örn: 1.250,50 -> 1250.50)
    clean = re.sub(r'[^\d,.]', '', str(text))
    if ',' in clean and '.' in clean:
        clean = clean.replace('.', '').replace(',', '.')
    elif ',' in clean:
        clean = clean.replace(',', '.')
    try:
        return float(clean)
    except:
        return None

def process_pdf(pdf_file, discount_str):
    extracted_data = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row in table:
                    # Boş satırları ele
                    row = [cell for cell in row if cell is not None]
                    if len(row) < 2: continue # En az bir isim bir fiyat olmalı
                    
                    # Varsayımsal sütun ayırıcı:
                    # Genelde ilk sütun KOD, son sütun FİYAT, ortadakiler İSİM'dir.
                    product_code = str(row[0]).strip()
                    raw_price = str(row[-1]).strip()
                    
                    # Ürün adı: İlk ve son sütun haricinde kalan her şey
                    product_name = " ".join([str(item) for item in row[1:-1]]).strip()
                    if not product_name: product_name = "Belirtilmemiş"

                    price_val = clean_price(raw_price)
                    
                    if price_val:
                        final_price = calculate_chain_discount(price_val, discount_str)
                        extracted_data.append({
                            "Ürün Kodu": product_code,
                            "Ürün Adı": product_name,
                            "Liste Fiyatı": price_val,
                            "İskonto Yapısı": discount_str,
                            "Net Fiyat": final_price
                        })
    return extracted_data

if uploaded_file:
    data = process_pdf(uploaded_file, discount_input)
    
    if data:
        df = pd.DataFrame(data)
        st.success(f"{len(df)} Ürün Başarıyla İşlendi!")
        
        # Önizleme tablosu
        st.dataframe(df, use_container_width=True)
        
        # Excel hazırlama
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Fiyat_Listesi')
        
        st.download_button(
            label="✅ Excel Dosyasını İndir",
            data=output.getvalue(),
            file_name="iskontolu_liste.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Uygun tablo yapısı bulunamadı. PDF'in 'Metin' formatında ve tablolu olduğundan emin olun.")
