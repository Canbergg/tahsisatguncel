import streamlit as st
import pandas as pd

def process_excel(file):
    df = pd.read_excel(file)
    
    # Yeni sütunları ekleyelim
    df.insert(df.columns.get_loc("MaxNeedForSalesParam"), "Unique Code", None)
    df.insert(df.columns.get_loc("MaxNeedForSalesParam") + 1, "İlişki", None)
    
    # Unique Code hesaplama
    unique_store_count = df['StoreName'].nunique()
    df['Unique Code'] = df['AE'].map(df['AE'].value_counts()) / unique_store_count
    
    # İlişki sütunu doldurma
    df['İlişki'] = df['AF'].apply(lambda x: "Muadil" if x == 11 else ("Muadil Stoksuz" if x == 10 else "İlişki yok"))
    
    return df

st.title("Excel Veri İşleme Uygulaması")
uploaded_file = st.file_uploader("Excel dosyanızı yükleyin", type=["xlsx", "xls"])

if uploaded_file is not None:
    processed_df = process_excel(uploaded_file)
    st.success("Dosya başarıyla işlendi!")
    st.dataframe(processed_df)
    
    # İşlenmiş dosyayı indirme bağlantısı oluşturma
    output_file = "processed_data.xlsx"
    processed_df.to_excel(output_file, index=False)
    with open(output_file, "rb") as file:
        st.download_button("İşlenmiş dosyayı indir", file, file_name=output_file)
git
gi
