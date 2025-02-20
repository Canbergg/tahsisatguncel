import streamlit as st
import pandas as pd

def process_excel(file):
    df = pd.read_excel(file)
    
    # Yeni sütunları ekleyelim
    df.insert(df.columns.get_loc("MaxNeedForSalesParam"), "Unique Code", None)
    df.insert(df.columns.get_loc("MaxNeedForSalesParam") + 1, "İlişki", None)
    
    # Unique Code hesaplama
    if 'Mağaza Adı' in df.columns and 'AE' in df.columns:
        unique_store_count = df['Mağaza Adı'].nunique()
        df['Unique Code'] = df['AE'].map(df['AE'].value_counts()) / unique_store_count
    else:
        st.error("Gerekli sütunlar eksik: 'Mağaza Adı' ve 'AE' sütunlarını kontrol edin.")
    
    # İlişki sütunu doldurma
    if 'AF' in df.columns:
        df['İlişki'] = df['AF'].apply(lambda x: "Muadil" if x == 11 else ("Muadil Stoksuz" if x == 10 else "İlişki yok"))
    else:
        st.error("'AF' sütunu eksik, lütfen dosyanızı kontrol edin.")
    
    return df

st.title("Excel Veri İşleme Uygulaması")
uploaded_file = st.file_uploader("Excel dosyanızı yükleyin", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        processed_df = process_excel(uploaded_file)
        st.success("Dosya başarıyla işlendi!")
        st.dataframe(processed_df)
        
        # İşlenmiş dosyayı indirme bağlantısı oluşturma
        output_file = "processed_data.xlsx"
        processed_df.to_excel(output_file, index=False)
        with open(output_file, "rb") as file:
            st.download_button("İşlenmiş dosyayı indir", file, file_name=output_file)
    except Exception as e:
        st.error(f"İşleme sırasında bir hata oluştu: {str(e)}")
