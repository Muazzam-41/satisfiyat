import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# Sayfa Geniş Modu
st.set_page_config(page_title="Nihai Katalog İşleyici", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f4f7f6; }
    .stButton>button { width: 100%; border-radius: 8px; height: 3.5em; background-color: #107c10; color: white; font-weight: bold; border: none; }
    .stButton>button:hover { background-color: #0b5a0b; color: white; }
    </style>
    """, unsafe_allow_html=True)

st.title("🛡️ Profesyonel Katalog & Fiyat Listesi Botu")
st.write("Tüm katalog formatlarını (1-2-3 sütun, tablolu, çoklu kodlu) destekleyen nihai sürüm.")

# Kullanıcı Arayüzü
col_file, col_disc = st.columns([2, 1])
with col_file:
    uploaded_file = st.file_uploader("PDF Katalog Dosyasını Seçin", type="pdf")
with col_disc:
    discount_input = st.text_input("İskonto Yapısı (Örn: 50+15+5)", value="50+10")
    st.caption("Zincir iskonto uygular: Önce %50, sonra %15, sonra %5 düşer.")

def calculate_chain_discount(price_str, discount_str):
    """₺1.250,50 gibi metinleri sayıya çevirir ve iskontoları uygular."""
    try:
        # Fiyatı temizle
        clean_p = price_str.replace('₺', '').replace('.', '').replace(',', '.').strip()
        val = float(clean_p)
        if not discount_str: return round(val, 2)
        
        discounts = [float(d.strip()) for d in discount_str.split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except:
        return 0

def process_column_logic(column_words, discount_str):
    """Bir sütun içindeki hiyerarşiyi (Başlık -> Kod -> Fiyat) çözer."""
    data_list = []
    current_title = "Tanımsız Ürün Grubu"
    
    # Regex Tanımları
    price_regex = r'\d{1,3}(?:\.\d{3})*,\d{2}' # 22,75 veya 13.890,48
    code_regex = r'\b[A-Z]{1,4}\s?\d{3,}\b'    # ARM 001111, ADS 731649, VDS 753023
    
    # Kelimeleri satırlara grupla
    lines = {}
    for w in column_words:
        y = round(w['top'])
        if y not in lines: lines[y] = []
        lines[y].append(w)
    
    sorted_y = sorted(lines.keys())
    
    for i, y in enumerate(sorted_y):
        line_words = lines[y]
        line_text = " ".join([w['text'] for w in line_words]).strip()
        
        # Fiyat ve Kod araması
        has_price = re.search(price_regex, line_text)
        has_code = re.search(code_regex, line_text)
        
        # 1. MANTIK: BAŞLIK TESPİTİ
        # Eğer satırda ne fiyat ne kod varsa ve yazılar büyükse/başlıksa
        if not has_price and not has_code:
            # Gereksiz teknik kelimeleri ele
            ignore = ["FİYAT", "KOLİ", "ADET", "KUTU", "P.ADET", "YENİ"]
            if len(line_text) > 3 and not any(x in line_text.upper() for x in ignore):
                # Eğer bir önceki satır da başlığın devamıysa birleştir
                if i > 0 and len(current_title) < 50:
                    current_title = line_text if current_title == "Tanımsız Ürün Grubu" else f"{current_title} {line_text}"
                else:
                    current_title = line_text
            continue

        # 2. MANTIK: KOD VE FİYAT EŞLEŞTİRME
        if has_price:
            price_found = has_price.group()
            
            # Satırda kod varsa al, yoksa bir üstteki kelimelere bak
            if has_code:
                code_found = has_code.group()
            else:
                # Kod bazen fiyatın bir solundaki kelimedir
                code_found = line_words[0]['text'] if len(line_words) > 1 else "Kodsuz Ürün"
            
            # Detay bilgisini al (Kod ve fiyat arasında kalan her şey)
            detail = line_text.replace(price_found, "").replace(code_found, "").strip()
            
            net_f = calculate_chain_discount(price_found, discount_str)
            
            data_list.append({
                "Ürün İsmi (Ana Grup)": current_title,
                "Ürün Kodu": code_found,
                "Özellik/Detay": detail if detail else "-",
                "Liste Fiyatı": price_found,
                "İskontolu Fiyat": net_f
            })
            
    return data_list

if uploaded_file:
    with st.spinner('Tüm sayfalar ve sütunlar analiz ediliyor...'):
        final_data = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                w, h = page.width, page.height
                
                # Sayfayı 3 sütuna böl (Tüm senaryoları kapsar: 1, 2 veya 3 sütun)
                # Sütun sınırları: 0 -> 0.33 -> 0.66 -> 1.0
                col_boundaries = [
                    (0, 0, w * 0.335, h),       # Sol
                    (w * 0.335, 0, w * 0.665, h), # Orta
                    (w * 0.665, 0, w, h)         # Sağ
                ]
                
                for bbox in col_boundaries:
                    crop = page.within_bbox(bbox)
                    col_words = crop.extract_words()
                    if col_words:
                        column_results = process_column_logic(col_words, discount_input)
                        final_data.extend(column_results)
        
        if final_data:
            df = pd.DataFrame(final_data)
            
            st.success(f"✅ Analiz Tamamlandı: {len(df)} kalem ürün bulundu.")
            
            # Tabloyu göster
            st.dataframe(df, use_container_width=True)
            
            # Excel Çıktısı
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Fiyat Listesi')
            
            st.download_button(
                label="📥 İskontolu Excel Listesini İndir",
                data=output.getvalue(),
                file_name="hazir_fiyat_listesi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Üzgünüm, PDF içerisinde uygun formatta ürün veya fiyat bulunamadı.")

st.divider()
st.info("""
**Sistem Nasıl Çalışır?**
- **Hiyerarşi:** Sayfanın üstündeki büyük yazıyı 'Grup İsmi' yapar.
- **Eşleşme:** Bu grup isminin altındaki tüm kodları (ARM, VDS, ADS vb.) o gruba bağlar.
- **Sütun Desteği:** Sayfa yan yana 2 veya 3 ürün içerse bile bunları birbirinden ayırır.
- **Hesaplama:** Liste fiyatından girdiğiniz % iskontoları sırasıyla düşer.
""")
