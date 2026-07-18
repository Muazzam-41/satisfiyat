import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# Sayfa Yapılandırması
st.set_page_config(page_title="Hatasız Fiyat & İskonto Botu", layout="wide")

# 1. YARDIMCI FONKSİYONLAR
def clean_string(s):
    """Kodlardaki boşluk ve tireleri temizler."""
    if not s or s == 'nan': return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).upper()

def calculate_discount(price_str, disc_str):
    """Fiyatı temizler ve iskontoyu matematiksel olarak uygular."""
    try:
        # Fiyattaki rakam ve virgül dışındaki her şeyi sil (₺ ve noktalar gider)
        # Örn: "₺1.250,50" -> "1250,50"
        clean_p = re.sub(r'[^\d,]', '', price_str)
        # Virgülü noktaya çevir (Python float formatı için)
        clean_p = clean_p.replace(',', '.')
        
        val = float(clean_p)
        
        if not disc_str or disc_str == "0":
            return round(val, 2)
        
        # Zincir iskonto: "50+10" -> [50.0, 10.0]
        discounts = [float(d.strip()) for d in str(disc_str).split('+') if d.strip()]
        
        for d in discounts:
            val = val * (1 - d / 100)
            
        return round(val, 2)
    except Exception as e:
        return 0

# 2. ARAYÜZ
st.title("🛡️ Hatasız Fiyat & İskonto Hesaplayıcı")
st.write("PDF fiyatlarını çeker ve belirttiğiniz iskontoyu (Örn: 50+10) hatasız uygular.")

with st.container():
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        reference_excel = st.file_uploader("📂 1. Referans Excel", type="xlsx")
    with col2:
        pdf_file = st.file_uploader("📄 2. PDF Katalog", type="pdf")
    with col3:
        discount_input = st.text_input("📉 İskonto Oranı (Örn: 20 veya 50+10)", value="20")

st.divider()

# 3. İŞLEM BAŞLATMA
if reference_excel and pdf_file:
    if st.button("🚀 HESAPLAMAYI BAŞLAT"):
        # Excel'i Oku
        ref_df = pd.read_excel(reference_excel)
        products_to_search = []
        for _, row in ref_df.iterrows():
            products_to_search.append({
                "name": str(row.iloc[0]).strip(),
                "code": str(row.iloc[1]).strip()
            })

        st.info("İşlem başladı, lütfen bekleyin...")

        found_data = []
        not_found = []
        # Fiyat yakalama deseni (1.250,00 veya 25,00 gibi yapıları bulur)
        price_pattern = re.compile(r'\d{1,3}(?:\.\d{3})*,\d{2}')

        with pdfplumber.open(pdf_file) as pdf:
            progress_bar = st.progress(0)
            
            for idx, target in enumerate(products_to_search):
                search_code = target['code']
                if not search_code or search_code == 'nan': continue
                
                is_found = False
                clean_target_code = clean_string(search_code)
                
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if not page_text: continue
                    
                    # Sayfada kod araması
                    if search_code in page_text or clean_target_code in clean_string(page_text):
                        prices_on_page = price_pattern.findall(page_text)
                        
                        if prices_on_page:
                            # İlk bulunan fiyat liste fiyatıdır
                            raw_price = prices_on_page[0]
                            net_price = calculate_discount(raw_price, discount_input)

                            found_data.append({
                                "Ürün İsmi": target['name'],
                                "Ürün Kodu": search_code,
                                "Liste Fiyatı": raw_price,
                                "Uygulanan İskonto": discount_input,
                                "İskontolu Net Fiyat": net_price,
                                "Sayfa No": page.page_number
                            })
                            is_found = True
                            break
                
                if not is_found:
                    not_found.append(target)
                
                progress_bar.progress((idx + 1) / len(products_to_search))

        # 4. SONUÇLAR
        if found_data:
            st.success(f"Başarılı! {len(found_data)} ürün hesaplandı.")
            res_df = pd.DataFrame(found_data).drop_duplicates(subset=['Ürün Kodu'])
            
            # Tabloyu göster
            st.dataframe(res_df, use_container_width=True)
            
            # Excel Hazırlama
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 İskontolu Excel'i İndir",
                data=output.getvalue(),
                file_name="iskontolu_fiyat_listesi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        if not_found:
            with st.expander(f"❌ Bulunamayan Ürünler ({len(not_found)})"):
                st.write("Aşağıdaki ürün kodları PDF içinde metin olarak bulunamadı.")
                st.table(pd.DataFrame(not_found))
else:
    st.warning("Lütfen dosyaları yükleyip 'Hesaplamayı Başlat' butonuna basın.")
