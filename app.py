import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="3 Sütunlu Katalog İşleyici", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #0275d8; color: white; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Profesyonel Katalog İşleyici (Üçlü Yan Yana Ürün Desteği)")
st.info("Bu versiyon sayfayı 3 sütuna böler ve ürün kodları ile çok satırlı başlıkları hatasız eşleştirir.")

# Girdiler
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("PDF Katalog Dosyasını Yükleyin", type="pdf")
with col2:
    discount_input = st.text_input("İskonto Oranları (Örn: 50+15+10)", value="50+10")

def calculate_net_price(list_price_str, discount_str):
    try:
        # ₺ simgesini ve noktaları temizle
        clean_price = list_price_str.replace('₺', '').replace('.', '').replace(',', '.').strip()
        price = float(clean_price)
        if not discount_str: return round(price, 2)
        discounts = [float(d.strip()) for d in discount_str.split('+') if d.strip()]
        for d in discounts:
            price = price * (1 - d / 100)
        return round(price, 2)
    except:
        return 0

def process_column(column_words, discount_str):
    """Bir sütun içindeki kelimeleri analiz eder (Sol, Orta veya Sağ)."""
    items = []
    
    # Regex Tanımları
    price_regex = r'\d{1,3}(?:\.\d{3})*,\d{2}'
    code_regex = r'\b[A-Z]{2,4}\s?\d{5,}\b' # ADS 731649 gibi
    
    # Kelimeleri satırlara grupla
    lines = {}
    for w in column_words:
        y = round(w['top'])
        if y not in lines: lines[y] = []
        lines[y].append(w)
    
    sorted_y = sorted(lines.keys())
    
    # Başlıkları, Kodları ve Fiyatları ayıkla
    # Katalog yapısına göre: Başlık -> Kod -> Resim (boşluk) -> Fiyat
    temp_title_parts = []
    found_code = None
    
    for y in sorted_y:
        line_text = " ".join([w['text'] for w in lines[y]]).strip()
        
        # 1. Kod Bulma (ADS 731649 gibi)
        code_match = re.search(code_regex, line_text)
        if code_match:
            found_code = code_match.group()
            # Koddan önce gelen her şey muhtemelen başlıktır
            current_title = " ".join(temp_title_parts)
            continue
        
        # 2. Fiyat Bulma (₺2.000,69 gibi)
        price_match = re.search(price_regex, line_text)
        if price_match and found_code:
            price_str = price_match.group()
            net_f = calculate_net_price(price_str, discount_str)
            
            items.append({
                "Ürün İsmi": current_title,
                "Ürün Kodu": found_code,
                "Liste Fiyatı": price_str,
                "İskontolu Fiyat": net_f
            })
            # Bir ürünü bitirince sıfırla
            temp_title_parts = []
            found_code = None
            continue
            
        # 3. Başlık Biriktirme (Eğer henüz kod bulunmadıysa satırları biriktir)
        if not found_code and len(line_text) > 3 and "Koli" not in line_text:
            temp_title_parts.append(line_text)

    return items

if uploaded_file:
    with st.spinner('Sayfalar 3 sütunlu olarak taranıyor...'):
        all_results = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                w, h = page.width, page.height
                
                # Sayfayı 3 eşit sütuna böl
                cols = [
                    (0, 0, w * 0.33, h),       # Sol
                    (w * 0.33, 0, w * 0.66, h), # Orta
                    (w * 0.66, 0, w, h)         # Sağ
                ]
                
                for bbox in cols:
                    col_words = page.within_bbox(bbox).extract_words()
                    if col_words:
                        all_results.extend(process_column(col_words, discount_input))
        
        if all_results:
            df = pd.DataFrame(all_results)
            st.success(f"✅ {len(df)} adet ürün başarıyla eşleştirildi!")
            st.dataframe(df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Excel Olarak İndir",
                data=output.getvalue(),
                file_name="uclu_katalog_fiyatlar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Veri bulunamadı. Lütfen PDF'in taranabilir bir metin dosyası olduğundan emin olun.")

st.divider()
st.caption("Bu sürüm ADS, BHV, NRA gibi kodları ve alt alta yazılmış ürün isimlerini 3 sütunlu sayfada takip eder.")
