import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import math

st.set_page_config(page_title="Hatasız Fiyat Eşleştirici PRO", layout="wide")

# 1. TEMİZLEME VE HESAPLAMA FONKSİYONLARI
def clean_string(s):
    if not s or s == 'nan': return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    try:
        # Fiyatı temizle: Rakam ve virgül dışındakileri at, virgülü noktaya çevir
        num_str = re.sub(r'[^\d,]', '', price_str).replace(',', '.')
        val = float(num_str)
        if not disc_str or disc_str == "0": return round(val, 2)
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

# 2. ARAYÜZ
st.title("🎯 Nokta Atışı: Hatasız Mıknatıs Eşleştirici")
st.write("Kodun sağındaki (satır sonu) veya altındaki (resim altı) fiyatı hatasız yakalar.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="20")

# 3. ANA İŞLEM
if reference_excel and pdf_file:
    if st.button("🚀 TARAMAYI VE HESAPLAMAYI BAŞLAT"):
        # Excel Hazırlığı
        ref_df = pd.read_excel(reference_excel)
        excel_items = []
        for _, row in ref_df.iterrows():
            name, code = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
            if code and code != 'nan':
                excel_items.append({"name": name, "orig": code, "clean": clean_string(code)})

        st.info(f"Derin analiz başladı: {len(excel_items)} ürün taranıyor...")
        
        results = []
        found_codes = set()
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress_bar = st.progress(0)
            
            for p_idx, page in enumerate(pdf.pages):
                words = page.extract_words()
                text = page.extract_text() or ""
                
                # Sayfadaki tüm fiyatları ve konumlarını bul
                page_prices = []
                for match in price_regex.finditer(text):
                    p_str = match.group()
                    kuruş = p_str.split(',')[-1]
                    for w in words:
                        if kuruş in w['text']:
                            page_prices.append({"text": p_str, "top": w['top'], "x0": w['x0']})
                            break
                
                if not page_prices: continue

                # Her Excel kodu için bu sayfada "Mıknatıs" taraması yap
                for target in excel_items:
                    if target['clean'] in found_codes: continue
                    
                    # Sayfada kodun geçtiği yeri bul
                    code_word = None
                    for w in words:
                        if target['clean'] == clean_string(w['text']):
                            code_word = w
                            break
                    
                    if code_word:
                        best_price = None
                        min_score = 999999
                        
                        for p in page_prices:
                            dy = p['top'] - code_word['top']
                            dx = p['x0'] - code_word['x0']
                            
                            # --- MIKNATIS MANTIĞI ---
                            
                            # DURUM 1: AYNI SATIRIN SONUNDA (dx pozitif, dy yaklaşık 0)
                            if abs(dy) < 10 and dx > 0:
                                score = dx * 0.5 # Çok güçlü puan
                            
                            # DURUM 2: RESMİN ALTINDA (dy pozitif, dx yaklaşık 0)
                            elif dy > 0 and abs(dx) < 150:
                                score = dy # Güçlü puan
                                
                            else:
                                score = 999999 # Geçersiz konum
                            
                            if score < min_score:
                                min_score = score
                                best_price = p['text']
                        
                        if best_price and min_score < 1000:
                            net_val = calculate_discount(best_price, discount_input)
                            results.append({
                                "Ürün İsmi": target['name'],
                                "Ürün Kodu": target['orig'],
                                "Liste Fiyatı": best_price,
                                "Net Fiyat": net_val,
                                "Sayfa": p_idx + 1
                            })
                            found_codes.add(target['clean'])

                progress_bar.progress((p_idx + 1) / len(pdf.pages))

        # 4. SONUÇLAR
        if results:
            res_df = pd.DataFrame(results).drop_duplicates(subset=['Ürün Kodu'])
            st.success(f"✅ {len(res_df)} / {len(excel_items)} ürün hatasız eşleşti!")
            st.dataframe(res_df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 Excel Olarak İndir", output.getvalue(), "hatasiz_liste.xlsx")
            
            not_found = [i for i in excel_items if i['clean'] not in found_codes]
            if not_found:
                with st.expander(f"❌ Bulunamayan Ürünler ({len(not_found)})"):
                    st.table(pd.DataFrame(not_found)[['name', 'orig']])
        else:
            st.error("Hiçbir eşleşme sağlanamadı. Lütfen PDF formatını kontrol edin.")
