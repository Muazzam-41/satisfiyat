import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Gelişmiş Hibrit Eşleştirici", layout="wide")

st.title("🚀 Akıllı Hibrit Eşleştirme Sistemi")
st.write("Kodlar değişmiş olsa bile hem koddan hem de isimden PDF'i tarar.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("1. Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("2. PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("İskonto (Örn: 50+10)", value="50+10")

def clean_string(s):
    """Kodlardaki boşluk ve tireleri temizler."""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    try:
        num_str = price_str.replace('₺', '').replace('.', '').replace(',', '.').strip()
        val = float(num_str)
        if not disc_str: return round(val, 2)
        discounts = [float(d.strip()) for d in disc_str.split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

if reference_excel and pdf_file:
    # Excel Verilerini Hazırla
    ref_df = pd.read_excel(reference_excel)
    # İlk sütun İsim, ikinci sütun Kod varsayılıyor
    products_to_search = []
    for _, row in ref_df.iterrows():
        name = str(row.iloc[0]).strip()
        code = str(row.iloc[1]).strip()
        products_to_search.append({
            "name": name,
            "code": code,
            "clean_code": clean_string(code),
            "clean_name": name.upper()[:20] # İsmin ilk 20 karakteriyle arama yapacağız
        })

    st.info(f"Excel'den {len(products_to_search)} ürün yüklendi. Hibrit tarama başlıyor...")

    found_data = []
    not_found_codes = []
    price_pattern = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

    with pdfplumber.open(pdf_file) as pdf:
        progress = st.progress(0)
        
        # Hız için tüm PDF metnini sayfa sayfa bir listeye alalım
        pages_content = []
        for p in pdf.pages:
            pages_content.append(p.extract_text() or "")

        for idx, target in enumerate(products_to_search):
            found = False
            
            for page_txt in pages_content:
                lines = page_txt.split('\n')
                for i, line in enumerate(lines):
                    clean_line = clean_string(line)
                    
                    # 1. KONTROL: Kod birebir var mı veya temizlenmiş hali satırda geçiyor mu?
                    if (target['code'] != 'nan' and target['code'] in line) or \
                       (target['clean_code'] != '' and target['clean_code'] in clean_line):
                        
                        # Fiyatı ara (aynı satırda veya sonraki 2 satırda)
                        search_area = " ".join(lines[i:i+3])
                        price_match = price_pattern.search(search_area)
                        
                        if price_match:
                            p_found = price_match.group()
                            found_data.append({
                                "Excel Ürün İsmi": target['name'],
                                "Excel Ürün Kodu": target['code'],
                                "PDF'de Bulunan Fiyat": p_found,
                                "Net Fiyat": calculate_discount(p_found, discount_input),
                                "Eşleşme Türü": "Kod Eşleşti"
                            })
                            found = True
                            break

                if found: break # Ürün bulunduysa diğer sayfalara bakma

                # 2. KONTROL (Eğer kodla bulunamadıysa): İsimden ara
                if not found and len(target['name']) > 5:
                    for i, line in enumerate(lines):
                        if target['name'].upper()[:15] in line.upper(): # İsmin başından yakala
                            search_area = " ".join(lines[i:i+3])
                            price_match = price_pattern.search(search_area)
                            if price_match:
                                p_found = price_match.group()
                                found_data.append({
                                    "Excel Ürün İsmi": target['name'],
                                    "Excel Ürün Kodu": target['code'],
                                    "PDF'de Bulunan Fiyat": p_found,
                                    "Net Fiyat": calculate_discount(p_found, discount_input),
                                    "Eşleşme Türü": "İsimden Yakalandı"
                                })
                                found = True
                                break
                if found: break

            if not found:
                not_found_codes.append(target)
            
            progress.progress((idx + 1) / len(products_to_search))

    # Sonuçları Göster
    st.divider()
    c1, c2 = st.columns(2)
    c1.metric("Toplam Ürün", len(products_to_search))
    c2.metric("Bulunan Ürün", len(found_data))

    if found_data:
        res_df = pd.DataFrame(found_data).drop_duplicates(subset=['Excel Ürün Kodu'], keep='first')
        st.success("✅ Eşleşen Ürünler Tablosu")
        st.dataframe(res_df, use_container_width=True)
        
        # İndirme
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            res_df.to_excel(writer, index=False)
        st.download_button("📥 Başarılı Eşleşmeleri İndir", out.getvalue(), "basarili_listem.xlsx")

    if not_found_codes:
        with st.expander("❌ Bulunamayan Ürünler Listesi"):
            st.table(pd.DataFrame(not_found_codes)[['name', 'code']])
