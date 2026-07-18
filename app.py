import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Referanslı Fiyat Botu", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stButton>button { width: 100%; background-color: #008CBA; color: white; font-weight: bold; border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🎯 Referans Excel Destekli Fiyat Botu")
st.write("Excel'deki kodları PDF içinde arar, fiyatı bulur ve senin belirlediğin ismi yazar.")

# Kullanıcı Girişleri
col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("1. Referans Excel'i Yükle (İsim ve Kod olan)", type="xlsx")
with col2:
    pdf_file = st.file_uploader("2. PDF Kataloğu Yükle", type="pdf")
with col3:
    discount_input = st.text_input("İskonto (Örn: 50+10)", value="50+10")

def calculate_discount(price_str, disc_str):
    try:
        # Fiyat temizleme: ₺, . ve boşlukları kaldır, virgülü noktaya çevir
        num_str = price_str.replace('₺', '').replace('.', '').replace(',', '.').strip()
        val = float(num_str)
        if not disc_str: return round(val, 2)
        discounts = [float(d.strip()) for d in disc_str.split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

if reference_excel and pdf_file:
    # 1. Excel'i Oku ve Sözlük Oluştur
    ref_df = pd.read_excel(reference_excel)
    # Varsayım: 1. Sütun Ürün İsmi, 2. Sütun Ürün Kodu
    # Kullanıcı sütun isimlerini bilmiyorsa index bazlı gidelim
    ref_mapping = {}
    for index, row in ref_df.iterrows():
        name = str(row.iloc[0]).strip()
        code = str(row.iloc[1]).strip()
        if code and code != 'nan':
            ref_mapping[code] = name

    st.info(f"Excel'den {len(ref_mapping)} adet kod sisteme yüklendi. PDF taranıyor...")

    final_results = []
    price_pattern = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

    with pdfplumber.open(pdf_file) as pdf:
        progress_bar = st.progress(0)
        total_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            text_lines = page.extract_text().split('\n')
            
            for line in text_lines:
                # Excel'deki her kodu bu satırda ara
                for code, excel_name in ref_mapping.items():
                    if code in line:
                        # Satırda fiyat ara
                        price_match = price_pattern.search(line)
                        if price_match:
                            price_found = price_match.group()
                            net_price = calculate_discount(price_found, discount_input)
                            
                            final_results.append({
                                "Ürün İsmi (Excel'den)": excel_name,
                                "Ürün Kodu": code,
                                "Liste Fiyatı (PDF'den)": price_found,
                                "Net Fiyat": net_price,
                                "İskonto": discount_input
                            })
            progress_bar.progress((i + 1) / total_pages)

    if final_results:
        result_df = pd.DataFrame(final_results)
        # Aynı kod birden fazla yerde varsa en sonuncuyu veya hepsini tutabiliriz. 
        # Genelde katalogda bir ürün bir kez olur. Duplicate'leri temizleyelim:
        result_df = result_df.drop_duplicates(subset=['Ürün Kodu'], keep='first')
        
        st.success(f"✅ Eşleşen {len(result_df)} ürün bulundu!")
        st.dataframe(result_df, use_container_width=True)

        # Excel Çıktısı
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Sonuç Excel Dosyasını İndir",
            data=output.getvalue(),
            file_name="pdf_fiyat_eslesme.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Üzgünüm, Excel'deki kodlar PDF içinde bulunamadı. Lütfen kodların PDF'te tam olarak nasıl yazıldığını kontrol edin.")
