import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import math

# Sayfa Ayarları
st.set_page_config(page_title="Hızlı Fiyat Eşleştirici PRO", layout="wide")

def clean_string(s):
    if not s or s == 'nan': return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    try:
        clean_p = re.sub(r'[^\d,]', '', price_str).replace(',', '.')
        val = float(clean_p)
        if not disc_str or disc_str == "0": return round(val, 2)
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

st.title("🎯 Nokta Atışı: Süper Hızlı Eşleştirici")
st.write("649 ürünlük listeniz için optimize edildi. PDF bir kez taranır ve eşleştirme anında yapılır.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="20")

if reference_excel and pdf_file:
    if st.button("🚀 TARAMAYI ŞİMDİ BAŞLAT"):
        # 1. Excel Verilerini Hızlı Sözlüğe Al
        ref_df = pd.read_excel(reference_excel)
        excel_map = {}
        for _, row in ref_df.iterrows():
            name = str(row.iloc[0]).strip()
            code = str(row.iloc[1]).strip()
            clean_c = clean_string(code)
            if clean_c:
                excel_map[clean_c] = {"name": name, "orig_code": code}

        target_codes = list(excel_map.keys())
        st.info(f"Excel'deki {len(excel_map)} ürün belleğe alındı. PDF taranıyor...")

        found_results = []
        found_codes_set = set()
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress_bar = st.progress(0)
            total_pages = len(pdf.pages)
            
            # PDF'İ SADECE BİR KEZ DÖNÜYORUZ
            for p_idx, page in enumerate(pdf.pages):
                words = page.extract_words()
                if not words: continue
                
                full_page_text = page.extract_text()
                
                # Bu sayfadaki tüm fiyatları ve koordinatlarını bul
                page_prices = []
                for m in price_regex.finditer(full_page_text):
                    p_str = m.group()
                    kuruş = p_str.split(',')[-1]
                    for w in words:
                        if kuruş in w['text']:
                            page_prices.append({"text": p_str, "y": w['top'], "x": w['x0']})
                            break
                
                if not page_prices: continue

                # Sayfadaki her kelimeye bak, Excel'deki kodlardan biriyse en yakın fiyatla eşleştir
                for w in words:
                    w_clean = clean_string(w['text'])
                    
                    # Eğer bu kelime bir kodun parçasıysa veya kodun kendisiyse
                    matched_code = None
                    for t_code in target_codes:
                        if t_code == w_clean or (len(w_clean) > 4 and w_clean in t_code):
                            matched_code = t_code
                            break
                    
                    if matched_code and matched_code not in found_codes_set:
                        # Bu koda en yakın fiyatı bul (Aynı satır öncelikli)
                        code_y = w['top']
                        code_x = w['x0']
                        
                        best_price = None
                        min_dist = float('inf')
                        
                        for p in page_prices:
                            dy = p['y'] - code_y
                            dx = abs(p['x'] - code_x)
                            
                            # Mesafe puanı: Aynı satırda olmaya (dy=0) büyük öncelik ver
                            # Fiyat kodun çok yukarısındaysa (-10 birimden fazla) onu ele
                            if dy < -10:
                                dist = 1000 + abs(dy)
                            else:
                                dist = math.sqrt(dx**2 + (dy * 5)**2) # Dikey kaymaya karşı daha hassas
                            
                            if dist < min_dist:
                                min_dist = dist
                                best_price = p['text']
                        
                        if best_price:
                            found_results.append({
                                "Ürün İsmi": excel_map[matched_code]['name'],
                                "Ürün Kodu": excel_map[matched_code]['orig_code'],
                                "Liste Fiyatı": best_price,
                                "Net Fiyat": calculate_discount(best_price, discount_input),
                                "Sayfa": p_idx + 1
                            })
                            found_codes_set.add(matched_code)

                progress_bar.progress((p_idx + 1) / total_pages)

        # 4. SONUÇLAR
        if found_results:
            res_df = pd.DataFrame(found_results).drop_duplicates(subset=['Ürün Kodu'])
            st.success(f"✅ {len(res_df)} ürün başarıyla eşleşti!")
            st.dataframe(res_df, use_container_width=True)
            
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Excel Dosyasını İndir",
                data=output_excel.getvalue(),
                file_name="guncel_fiyatlar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Bulunamayanlar
            not_found = [excel_map[c] for c in target_codes if c not in found_codes_set]
            if not_found:
                with st.expander(f"❌ Bulunamayan Ürünler ({len(not_found)})"):
                    st.write("PDF içinde kodu veya fiyatı saptanamayan ürünler:")
                    st.table(pd.DataFrame(not_found)[['name', 'orig_code']])
        else:
            st.error("Hiçbir ürün eşleşmedi. Lütfen kodların doğruluğunu kontrol edin.")
