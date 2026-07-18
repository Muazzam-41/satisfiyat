import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import math

# Sayfa Ayarları
st.set_page_config(page_title="Hatasız Fiyat Eşleştirici", layout="wide")

# 1. FONKSİYONLAR
def clean_string(s):
    if not s or s == 'nan': return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    try:
        # Fiyatı temizle: Rakam ve virgül dışındakileri at
        clean_p = re.sub(r'[^\d,]', '', price_str).replace(',', '.')
        val = float(clean_p)
        if not disc_str or disc_str == "0": return round(val, 2)
        
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

def get_distance(w1, w2):
    """İki kelime (koordinat) arasındaki mesafeyi hesaplar (Öncelik: Aşağı ve Sağ)"""
    dx = w2['x0'] - w1['x0']
    dy = w2['top'] - w1['top']
    
    # Fiyat genelde kodun altındadır (dy > 0) veya sağındadır (dx > 0)
    # Eğer fiyat kodun çok üstündeyse mesafeyi yapay olarak büyüt (hatalı eşleşmeyi önle)
    if dy < -10: dy = dy * -10 
    
    return math.sqrt(dx**2 + dy**2)

# 2. ARAYÜZ
st.title("🎯 Nokta Atışı Fiyat Eşleştirici")
st.write("Kodun koordinatını bulur ve ona en yakın mesafedeki fiyatı çeker. Kaymaları önler.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+10)", value="20")

if reference_excel and pdf_file:
    if st.button("🚀 CTRL+F TARAMASINI BAŞLAT"):
        # Excel Oku
        ref_df = pd.read_excel(reference_excel)
        products = []
        for _, row in ref_df.iterrows():
            products.append({"name": str(row.iloc[0]).strip(), "code": str(row.iloc[1]).strip()})

        st.info(f"{len(products)} ürün için konumsal tarama yapılıyor...")
        
        found_data = []
        not_found = []
        # Fiyat deseni
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress = st.progress(0)
            
            for idx, target in enumerate(products):
                search_code = target['code']
                if not search_code or search_code == 'nan': continue
                
                clean_target = clean_string(search_code)
                is_matched = False
                
                for page in pdf.pages:
                    # Sayfadaki tüm kelimeleri koordinatlarıyla çek
                    words = page.extract_words()
                    
                    # 1. KODUN KOORDİNATINI BUL
                    code_coords = None
                    for w in words:
                        if search_code in w['text'] or clean_target == clean_string(w['text']):
                            code_coords = w
                            break
                    
                    if code_coords:
                        # 2. AYNI SAYFADAKİ TÜM FİYATLARI VE KOORDİNATLARINI BUL
                        # Bu işlem için sayfadaki metni 'text' olarak değil, 'words' olarak analiz etmeliyiz
                        # Fiyat genelde parçalı gelebilir, bu yüzden regex'i sayfada tekrar çalıştırıyoruz
                        page_text_with_coords = page.extract_words()
                        
                        # Sayfadaki fiyat olabilecek tüm kelimeleri/yapıları bul
                        potential_prices = []
                        # Sayfa metnini alıp içindeki fiyatların yerlerini bulalım
                        full_text = page.extract_text()
                        for match in price_regex.finditer(full_text):
                            p_str = match.group()
                            # Bu fiyatın sayfadaki koordinatını yaklaşık olarak bul
                            # (Hız ve doğruluk için fiyatın geçtiği satırdaki kelimeleri baz alıyoruz)
                            for w in words:
                                if p_str.split(',')[-1] in w['text']: # Kuruş kısmından yakala
                                    potential_prices.append({"text": p_str, "coords": w})
                                    break
                        
                        if potential_prices:
                            # 3. KODA EN YAKIN FİYATI SEÇ
                            # Mesafeye göre sırala
                            best_price = min(potential_prices, key=lambda p: get_distance(code_coords, p['coords']))
                            
                            p_val = best_price['text']
                            net_val = calculate_discount(p_val, discount_input)
                            
                            found_data.append({
                                "Ürün İsmi": target['name'],
                                "Ürün Kodu": search_code,
                                "Liste Fiyatı": p_val,
                                "Net Fiyat": net_val,
                                "Sayfa": page.page_number
                            })
                            is_matched = True
                            break
                
                if not is_matched:
                    not_found.append(target)
                
                progress.progress((idx + 1) / len(products))

        # 4. SONUÇLAR
        if found_data:
            res_df = pd.DataFrame(found_data).drop_duplicates(subset=['Ürün Kodu'])
            st.success(f"✅ {len(res_df)} ürün doğru fiyatla eşleşti!")
            st.dataframe(res_df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            
            st.download_button("📥 Excel İndir", output.getvalue(), "hatasiz_liste.xlsx")
        
        if not_found:
            with st.expander("❌ Bulunamayanlar"):
                st.table(pd.DataFrame(not_found))
