import streamlit as st
import pandas as pd
import io
import re
from difflib import SequenceMatcher

# Sayfa Yapılandırması
st.set_page_config(page_title="Gelişmiş Excel İşlem Merkezi", layout="wide")

# --- YARDIMCI FONKSİYONLAR ---
def clean_code(text):
    if not text or pd.isna(text): return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).upper()

def clean_name(text):
    if not text or pd.isna(text): return ""
    text = str(text).lower()
    cleaned = re.sub(r'[^a-z0-9ğüşıöç ]', '', text).strip() # Sembolleri sil, boşlukları koru
    return cleaned

def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

# Manuel eşleşmeleri saklamak için Session State (Sayfa yenilense de veriler kaybolmaz)
if 'manual_matches' not in st.session_state:
    st.session_state.manual_matches = []

# --- ARAYÜZ ---
st.title("🛡️ Profesyonel Excel İşlem ve Karşılaştırma Merkezi")

# ==============================================================================
# BÖLÜM 1: FİYAT HESAPLAMA (Stabil Bölüm)
# ==============================================================================
st.header("1️⃣ Muhasebe Entegre Fiyat Hesaplama")
# ... (Önceki bölüme ait kodlar burada aynı şekilde kalabilir, yer kaplamaması için özet geçiyorum)
# Not: Eğer Bölüm 1'i tamamen korumak istiyorsan önceki koddaki haliyle buraya yapıştırabilirsin.

st.divider()

# ==============================================================================
# BÖLÜM 2: HİBRİT STOK VE MANUEL EŞLEŞTİRME SİSTEMİ
# ==============================================================================
st.header("2️⃣ Stok Karşılaştırma ve Manuel Seçim Asistanı")

col_st1, col_st2 = st.columns(2)
with col_st1:
    real_stock_file = st.file_uploader("Kendi Reel Stok Excel'iniz", type="xlsx", key="st_real")
with col_st2:
    compare_stock_file = st.file_uploader("Karşılaştırılacak İkinci Excel", type="xlsx", key="st_comp")

if real_stock_file and compare_stock_file:
    df_real = pd.read_excel(real_stock_file)
    df_comp = pd.read_excel(compare_stock_file)

    st.write("⚙️ **Sütun Ayarları**")
    st1, st2, st3, st4 = st.columns(4)
    with st1: r_code_col = st.selectbox("Reel Stok: Kod Sütunu", df_real.columns, key="st_s1")
    with st2: r_name_col = st.selectbox("Reel Stok: İsim Sütunu", df_real.columns, key="st_s2")
    with st3: c_code_col = st.selectbox("2. Excel: Kod Sütunu", df_comp.columns, key="st_s3")
    with st4: c_name_col = st.selectbox("2. Excel: İsim Sütunu", df_comp.columns, key="st_s4")

    if st.button("🚀 Otomatik Eşleştirmeyi Başlat"):
        # Verileri hazırla
        comp_by_code = {clean_code(r[c_code_col]): r for _, r in df_comp.iterrows() if pd.notna(r[c_code_col])}
        comp_by_name = {clean_name(r[c_name_col]): r for _, r in df_comp.iterrows() if pd.notna(r[c_name_col])}

        auto_matched = []
        not_found = []

        for _, row in df_real.iterrows():
            r_code_c = clean_code(row[r_code_col])
            r_name_c = clean_name(row[r_name_col])

            # 1. Kod ile
            if r_code_c and r_code_c in comp_by_code:
                match_row = comp_by_code[r_code_c]
                auto_matched.append({**row.to_dict(), "Eşleşme Türü": "KOD İLE", **match_row.add_prefix("Karşı_")})
            # 2. İsim ile
            elif r_name_c and r_name_c in comp_by_name:
                match_row = comp_by_name[r_name_c]
                auto_matched.append({**row.to_dict(), "Eşleşme Türü": "İSİM İLE", **match_row.add_prefix("Karşı_")})
            # 3. Bulunamadı
            else:
                not_found.append(row.to_dict())
        
        st.session_state.auto_matched = auto_matched
        st.session_state.not_found = not_found
        st.session_state.comp_list = df_comp.to_dict('records')

    # --- EŞLEŞTİRME ASİSTANI ARAYÜZÜ ---
    if 'not_found' in st.session_state and st.session_state.not_found:
        tab_ok, tab_manual = st.tabs(["✅ Başarılı Eşleşmeler", "🔍 Manuel Eşleştirme Asistanı"])

        with tab_ok:
            all_matches = st.session_state.auto_matched + st.session_state.manual_matches
            if all_matches:
                df_final = pd.DataFrame(all_matches)
                st.dataframe(df_final, use_container_width=True)
                
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                st.download_button("📥 Tüm Listeyi İndir (Otomatik + Manuel)", out.getvalue(), "stok_eslesme_final.xlsx")

        with tab_manual:
            st.warning(f"Sistem {len(st.session_state.not_found)} ürünü otomatik eşleştiremedi. Aşağıdan manuel seçebilirsiniz.")
            
            # Manuel seçim alanı
            unmatched_item = st.selectbox("Eşleştirilecek Ürünü Seçin:", 
                                         options=st.session_state.not_found, 
                                         format_func=lambda x: f"{x[r_name_col]} (Kod: {x[r_code_col]})")
            
            if unmatched_item:
                search_term = clean_name(unmatched_item[r_name_col])
                
                # Öneriler oluştur (Sadece en benzer 5 taneyi getir)
                suggestions = sorted(st.session_state.comp_list, 
                                   key=lambda x: similarity(search_term, clean_name(x[c_name_col])), 
                                   reverse=True)[:5]
                
                st.write("💡 **Sistem Önerileri (En Yakın İsimler):**")
                
                for sug in suggestions:
                    col_text, col_btn = st.columns([4, 1])
                    sug_label = f"{sug[c_name_col]} | Kod: {sug[c_code_col]}"
                    col_text.write(f"🔹 {sug_label}")
                    
                    if col_btn.button("Eşleştir", key=f"btn_{sug[c_code_col]}_{unmatched_item[r_code_col]}"):
                        # Manuel eşleşmeyi kaydet
                        new_match = {**unmatched_item, "Eşleşme Türü": "MANUEL", **pd.Series(sug).add_prefix("Karşı_")}
                        st.session_state.manual_matches.append(new_match)
                        
                        # Listeden çıkar
                        st.session_state.not_found.remove(unmatched_item)
                        st.success(f"Eşleşti: {unmatched_item[r_name_col]} ↔️ {sug[c_name_col]}")
                        st.rerun()

                st.write("---")
                st.write("🔎 **Aradığınızı yukarıda bulamadıysanız tüm listeden seçin:**")
                manual_selection = st.selectbox("İkinci Excel'den Karşılığını Seçin:", 
                                               options=st.session_state.comp_list,
                                               format_func=lambda x: f"{x[c_name_col]} (Kod: {x[c_code_col]})")
                
                if st.button("Seçilenle Eşleştir"):
                    new_match = {**unmatched_item, "Eşleşme Türü": "MANUEL SEÇİM", **pd.Series(manual_selection).add_prefix("Karşı_")}
                    st.session_state.manual_matches.append(new_match)
                    st.session_state.not_found.remove(unmatched_item)
                    st.rerun()
