import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Pro Fiyat Dedektifi", layout="wide")

# 1. YARDIMCI FONKSİYONLAR (En üstte)
def clean_string(s):
    """Kodlardaki boşluk, tire ve özel karakterleri temizler."""
    if not s or s == 'nan': return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    """Zincir iskonto hesaplar (Örn: 50+10)."""
    try:
        # Fiyatı sayıya çevir (₺25,09 -> 25.09)
        num_str = price_str.replace('₺', '').replace('.', '').replace(',', '.').strip()
        val = float(num_str)
        if not disc_str: return round(val, 2)
        
        # Artı işaretiyle ayrılmış iskontoları listele
        discounts = [float(d.strip()) for d in disc_str.split('+') if d.strip()]
        for d in discounts:
            val = val * (1 - d / 100)
        return round(val, 2)
    except:
        return 0

# 2. ARAYÜZ
st.title("🔍 Akıllı Fiyat Eşleştirici")
st.write("Dosyaları yükleyin, iskontoyu girin ve 'Başlat' butonuna tıklayın.")

# Ayarlar Bölümü
with st.container():
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        reference_excel = st.file_uploader("📂 1. Referans Excel'i Seçin", type="xlsx")
    with col2:
        pdf_file = st.file_uploader("📄 2. PDF Kataloğu Seçin", type="pdf")
    with col3:
        discount_input = st.text_input("📉 İskonto (Örn: 50+10)", value="20")

st.divider()

# 3. ANA İŞLEM BLOĞU
if reference_excel and pdf_file:
    if st.button("🚀 TARAMAYI VE HESAPLAMAYI BAŞLAT"):
        
        # Excel Verilerini Oku
        ref_df = pd.read_excel(reference_excel)
        products_to_search = []
        for _, row in ref_df.iterrows():
            products_to_search.append({
                "name": str(row.iloc[0]).strip(),
                "code": str(row.iloc[1]).strip()
            })

        st.info(f"Tarama başladı: {len(products_to_search)} ürün aranıyor...")

        found_data = []
        not_found = []
        # Fiyat yakalamak için geliştirilmiş Regex
        price_pattern = re.compile(r'₺?\s?\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress_bar = st.progress(0)
            
            for idx, target in enumerate(products_to_search):
                search_code = target['code']
                if not search_code or search_code == 'nan' or search_code == "":
                    continue
                
                is_found = False
                clean_target_code = clean_string(search_code)
                
                # Her bir kodu PDF sayfalarında ara (CTRL+F mantığı)
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if not page_text: continue
                    
                    # Sayfada kodun geçip geçmediğini kontrol et (Normal ve temizlenmiş haliyle)
                    if search_code in page_text or clean_target_code in clean_string(page_text):
                        all_prices = price_pattern.findall(page_text)
                        
                        if all_prices:
                            # O sayfadaki ilk fiyatı al (Genelde koda en yakın fiyattır)
                            p_found = all_prices[0] 

                            found_data.append({
                                "Ürün İsmi": target['name'],
                                "Ürün Kodu": search_code,
                                "Liste Fiyatı": p_found,
                                "İskontolu Fiyat": calculate_discount(p_found, discount_input),
                                "Sayfa No": page.page_number
                            })
                            is_found = True
                            break
                
                if not is_found:
                    not_found.append(target)
                
                # Progress bar güncellemesi
                progress_bar.progress((idx + 1) / len(products_to_search))

        # 4. SONUÇLAR VE İNDİRME
        if found_data:
            st.success(f"İşlem Tamamlandı! {len(found_data)} ürün eşleşti.")
            res_df = pd.DataFrame(found_data).drop_duplicates(subset=['Ürün Kodu'])
            st.dataframe(res_df, use_container_width=True)
            
            # Excel Hazırlama
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            
            # Hatalı olan satır düzeltildi:
            excel_data = output_excel.getvalue()
            
            st.download_button(
                label="📥 Güncellenmiş Excel'i İndir",
                data=excel_data,
                file_name="guncel_fiyatlar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Hiçbir ürün eşleşmedi. Lütfen kodları ve PDF içeriğini kontrol edin.")
        
        if not_found:
            with st.expander(f"❌ Bulunamayan Ürünler ({len(not_found)})"):
                st.table(pd.DataFrame(not_found))
else:
    st.warning("Lütfen işlem yapabilmek için önce dosyaları yükleyin.")
