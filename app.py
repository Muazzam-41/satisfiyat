import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# Sayfa Ayarları
st.set_page_config(page_title="Nihai Katalog İşleyici Pro", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 10px; height: 3.5em; background-color: #2c3e50; color: white; font-weight: bold; border: none; }
    .stButton>button:hover { background-color: #34495e; color: white; }
    </style>
    """, unsafe_allow_html=True)

st.title("🛡️ Profesyonel Katalog Veri Çekme Sistemi")
st.write("Tüm senaryolar (Detay sütunu, reklam ayıklama, kodsuz ürünler, 1-2-3 sütun yerleşim) tanımlandı.")

# Girişler
col_f, col_i = st.columns([2, 1])
with col_f:
    uploaded_file = st.file_uploader("PDF Katalog Dosyasını Seçin", type="pdf")
with col_i:
    discount_input = st.text_input("Uygulanacak İskonto (Örn: 50+15+5)", value="50+10")

def calculate_discount(price_str, disc_str):
    """Fiyatı temizler ve zincir iskonto uygular."""
    try:
        # ₺, . ve boşlukları temizle, virgülü noktaya çevir
        num_str = price_str.replace('₺', '').replace('.', '').replace(',', '.').strip()
        val = float(num_str)
        if not disc_str: return round(val, 2)
        discounts = [float(d.strip()) for d in disc_str.split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

def process_zone(words, discount_str):
    """Bir sütun içindeki dikey hiyerarşiyi (Başlık -> Kod -> Detay -> Fiyat) çözer."""
    extracted = []
    
    # Regex Kalıpları (Senin verdiğin tüm formatlar dahil)
    code_pattern = re.compile(r'\b[A-Z]{1,4}\s?\d{3,}\b|\b[A-Z]\s\d{4,}\b') 
    price_pattern = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}') 
    
    # Kelimeleri satırlara göre grupla
    lines = {}
    for w in words:
        y = round(w['top'])
        if y not in lines: lines[y] = []
        lines[y].append(w)
    
    sorted_y = sorted(lines.keys())
    
    current_name = ""
    current_code = ""
    current_details = []
    
    # Yasaklı kelimeler (Başlıkta veya detayda istemediklerimiz)
    ignore_list = ["FİYAT", "P.ADET", "NAKİT", "EROL TEKNİK", "ARMATÜR", "SAYFA", "VİTRİFİYE", "BANYO DOLAPLARI"]

    for i, y in enumerate(sorted_y):
        line_words = lines[y]
        line_text = " ".join([w['text'] for w in line_words]).strip()
        
        # 1. KOD VE FİYAT TESPİTİ
        code_match = code_pattern.search(line_text)
        price_match = price_pattern.search(line_text)
        
        if price_match:
            price_found = price_match.group()
            item_code = code_match.group() if code_match else current_code
            
            # Fiyat bulunduysa paketi hazırla
            net_f = calculate_discount(price_found, discount_str)
            
            # Kod ve Fiyat dışındaki her şeyi detaya at (Örn: 90 CM BEYAZ)
            line_detail = line_text.replace(price_found, "").replace(item_code if item_code != "KODSUZ" else "", "").strip()
            all_details = " | ".join(current_details)
            if line_detail and line_detail.upper() not in ["FİYAT", "NAKİT"]:
                all_details = f"{all_details} | {line_detail}" if all_details else line_detail

            extracted.append({
                "Ürün İsmi": current_name if current_name else "Bilinmeyen Ürün",
                "Ürün Kodu": item_code if item_code else "KODSUZ",
                "Detaylar": all_details,
                "Liste Fiyatı": price_found,
                "Net Fiyat": net_f
            })
            # Resetle (Sadece fiyat gelince ürün biter, başlık bir sonraki ürüne kadar korunmaz)
            current_code = ""
            current_details = []
            continue

        if code_match:
            current_code = code_match.group()
            # Satırda koddan başka yazı varsa detaya ekle
            rem = line_text.replace(current_code, "").strip()
            if rem and not any(ign in rem.upper() for ign in ignore_list):
                current_details.append(rem)
            continue

        # 2. İSİM VE DETAY AYIRIMI
        if len(line_text) > 2 and not any(ign in line_text.upper() for ign in ignore_list):
            # Teknik kelimeler varsa detaydır, yoksa isimdir
            tech_keys = ["mm", "adet", "gr", "koli", "bağ", "cm", "pvc", "kpk", "içt"]
            if any(tk in line_text.lower() for tk in tech_keys):
                current_details.append(line_text)
            else:
                # Eğer isim boşsa bu isimdir, doluysa üstüne ekle (Çok satırlı başlıklar için)
                if not current_name or i == 0:
                    current_name = line_text
                else:
                    # Yeni bir büyük blok mu yoksa ismin devamı mı?
                    # Eğer son 2 satırdır bir şey bulunmadıysa ismi günceller
                    current_name = f"{current_name} {line_text}" if len(current_details) == 0 else line_text
                        
    return extracted

if uploaded_file:
    with st.spinner('Katalog yapısı (1-2-3 sütun) analiz ediliyor...'):
        final_list = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                w, h = page.width, page.height
                
                # Sayfayı 3 dikey sütuna böl (Tüm yerleşimleri kapsar)
                zones = [
                    (0, 0, w * 0.335, h),
                    (w * 0.335, 0, w * 0.665, h),
                    (w * 0.665, 0, w, h)
                ]
                
                for bbox in zones:
                    zone_words = page.within_bbox(bbox).extract_words()
                    if zone_words:
                        results = process_zone(zone_words, discount_input)
                        final_list.append(results)
        
        # Verileri düzleştir ve temizle
        flat_list = [item for sublist in final_list for item in sublist]
        
        if flat_list:
            df = pd.DataFrame(flat_list)
            # İsim temizliği (Fiyat kelimesi ve gereksiz boşluklar)
            df['Ürün İsmi'] = df['Ürün İsmi'].apply(lambda x: re.sub(r'(?i)fiyat|nakit', '', str(x)).strip())
            
            st.success(f"✅ Analiz Başarılı! {len(df)} ürün bulundu.")
            st.dataframe(df, use_container_width=True)
            
            # Excel Çıktısı
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Fiyat Listesi')
            
            st.download_button(
                label="📥 Excel Dosyasını İndir",
                data=output.getvalue(),
                file_name="katalog_nihai_liste.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Veri bulunamadı. Lütfen PDF'in metin tabanlı olduğundan emin olun.")

st.divider()
st.caption("Bu yazılım gönderdiğiniz tüm örneklere (ARM, T, VDS, VLV, DLM kodları ve detay sütunu) göre optimize edilmiştir.")
