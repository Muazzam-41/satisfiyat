import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Hibrit Fiyat Botu (Yüksek Başarı)", layout="wide")

def clean_string(s):
    if not s or s == 'nan': return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    try:
        # Fiyat temizleme: ₺, nokta ve boşlukları at, virgülü noktaya çevir
        num_str = re.sub(r'[^\d,]', '', price_str).replace(',', '.')
        val = float(num_str)
        if not disc_str or disc_str == "0": return round(val, 2)
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except: return 0

st.title("🛡️ Yüksek Başarı Oranlı Hibrit Fiyat Botu")
st.write("660+ ürün için optimize edildi. Kod ve fiyat arasındaki tüm yerleşim farklarını tolore eder.")

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    reference_excel = st.file_uploader("📂 1. Referans Excel", type="xlsx")
with col2:
    pdf_file = st.file_uploader("📄 2. PDF Katalog", type="pdf")
with col3:
    discount_input = st.text_input("📉 İskonto (Örn: 50+15)", value="20")

if reference_excel and pdf_file:
    if st.button("🚀 GENİŞ KAPSAMLI TARAMAYI BAŞLAT"):
        # Excel hazırlığı
        ref_df = pd.read_excel(reference_excel)
        excel_items = []
        for _, row in ref_df.iterrows():
            name, code = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
            c_code = clean_string(code)
            if c_code:
                excel_items.append({"name": name, "orig_code": code, "clean": c_code})

        st.info(f"{len(excel_items)} ürün için tarama yapılıyor...")
        
        final_results = []
        found_codes = set()
        price_regex = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress = st.progress(0)
            
            for p_idx, page in enumerate(pdf.pages):
                words = page.extract_words()
                text = page.extract_text() or ""
                
                # Sayfadaki tüm fiyatları koordinatlarıyla bul
                page_prices = []
                for match in price_regex.finditer(text):
                    p_str = match.group()
                    kuruş = p_str.split(',')[-1]
                    for w in words:
                        if kuruş in w['text']:
                            page_prices.append({"text": p_str, "top": w['top'], "x0": w['x0'], "bottom": w['bottom']})
                            break
                
                if not page_prices: continue

                # Bu sayfadaki her kelimeyi tara
                for w in words:
                    w_clean = clean_string(w['text'])
                    
                    # Kelime Excel'deki bir kod mu veya kodun parçası mı?
                    matched_item = None
                    for item in excel_items:
                        if item['clean'] == w_clean or (len(w_clean) > 5 and w_clean in item['clean']):
                            if item['clean'] not in found_codes:
                                matched_item = item
                                break
                    
                    if matched_item:
                        # --- HİBRİT EŞLEŞTİRME MANTIĞI ---
                        best_price = None
                        min_dist = 999999
                        
                        for p_item in page_prices:
                            dy = p_item['top'] - w['top']
                            dx = abs(p_item['x0'] - w['x0'])
                            
                            # Strateji 1: Aynı satırda mı? (En yüksek öncelik)
                            if abs(dy) < 8 and dx < 400:
                                dist = dx * 0.1 # Yatay mesafe çok az etkili olsun
                            # Strateji 2: Kodun altında mı? (Katalog düzeni)
                            elif 8 <= dy < 250 and dx < 150:
                                dist = dy + dx
                            # Strateji 3: Çok uzak ama aynı sütunda mı?
                            elif 250 <= dy < 450 and dx < 50:
                                dist = dy * 1.5
                            else:
                                continue # Çok uzak veya kodun yukarısında
                            
                            if dist < min_dist:
                                min_dist = dist
                                best_price = p_item['text']
                        
                        if best_price:
                            net_val = calculate_discount(best_price, discount_input)
                            final_results.append({
                                "Ürün İsmi": matched_item['name'],
                                "Ürün Kodu": matched_item['orig_code'],
                                "Liste Fiyatı": best_price,
                                "Net Fiyat": net_val,
                                "Sayfa": p_idx + 1
                            })
                            found_codes.add(matched_item['clean'])

                progress.progress((p_idx + 1) / len(pdf.pages))

        # Sonuçlar
        if final_results:
            res_df = pd.DataFrame(final_results).drop_duplicates(subset=['Ürün Kodu'])
            st.success(f"✅ {len(res_df)} / {len(excel_items)} Ürün Başarıyla Bulundu!")
            
            st.dataframe(res_df, use_container_width=True)
            
            # Excel İndirme
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            st.download_button("📥 İskontolu Excel'i İndir", out.getvalue(), "hatasiz_fiyat_listesi.xlsx")
            
            # Bulunamayanlar Listesi
            not_found = [i for i in excel_items if i['clean'] not in found_codes]
            if not_found:
                with st.expander(f"❌ Bulunamayan Ürünler ({len(not_found)}) - Bunları Manuel Kontrol Edin"):
                    st.table(pd.DataFrame(not_found)[['name', 'orig_code']])
