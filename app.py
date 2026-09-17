import streamlit as st
import pandas as pd
from collections import Counter

# Konfigurasi halaman antarmuka
st.set_page_config(page_title="Advanced Scanner - Pola As", layout="wide")

st.title("Aplikasi Pemindai Probabilitas Posisi 'As'")
st.markdown("Menggali persentase kemunculan angka H+1 berdasarkan data historis dari Excel.")

# 1. Komponen Unggah File
uploaded_file = st.file_uploader("Unggah file Excel referensi (misal: HK.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Membaca data menggunakan pandas
    df = pd.read_excel(uploaded_file)
    
    # Membersihkan baris pertama jika berupa baris pemisah/kosong
    if df.iloc[0].isna().all():
        df = df.iloc[1:].reset_index(drop=True)
        
    # 2. Algoritma perataan (flattening) matriks ke urutan kronologis harian
    daily_data = []
    
    # Pemetaan kolom berdasarkan struktur matriks pada HK.xlsx
    days_cols = {
        'SABTU': ['SABTU', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3'],
        'MINGGU': ['MINGGU', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8'],
        'SENIN': ['SENIN', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 13'],
        'SELASA': ['SELASA', 'Unnamed: 16', 'Unnamed: 17', 'Unnamed: 18'],
        'RABU': ['RABU', 'Unnamed: 21', 'Unnamed: 22', 'Unnamed: 23'],
        'KAMIS': ['KAMIS', 'Unnamed: 26', 'Unnamed: 27', 'Unnamed: 28'],
        'JUMAT': ['JUMAT', 'Unnamed: 31', 'Unnamed: 32', 'Unnamed: 33']
    }

    # Looping per baris mingguan, lalu dipecah per hari secara berurutan
    for index, row in df.iterrows():
        for day, cols in days_cols.items():
            try:
                # Memastikan nilai 'As' pada hari tersebut memiliki data
                if pd.notna(row[cols[0]]): 
                    # Menyimpan susunan: [As, Kop, Kepala, Ekor]
                    daily_data.append([row[cols[0]], row[cols[1]], row[cols[2]], row[cols[3]]])
            except KeyError:
                continue 

    st.success(f"Basis data siap. Total riwayat hari terekstrak: {len(daily_data)} hari.")
    st.divider()
    
    # 3. Parameter Pemindaian
    st.subheader("Parameter Analisis")
    target_as = st.number_input("Masukkan angka As acuan hari ini (0-9):", min_value=0, max_value=9, value=4, step=1)
    
    if st.button("Jalankan Scanner", type="primary"):
        next_as_list = []
        
        # 4. Mesin Pencarian Pola H+1
        # Loop dibatasi sampai len - 1 karena kita butuh memeriksa hari setelahnya
        for i in range(len(daily_data) - 1): 
            current_as = daily_data[i][0]
            if current_as == target_as:
                # Ambil nilai As pada indeks hari berikutnya (H+1)
                next_day = daily_data[i+1]
                next_as = next_day[0]
                
                # Memastikan nilai tersebut adalah angka (bukan NaN)
                if pd.notna(next_as):
                    next_as_list.append(next_as)
                
        # 5. Kalkulasi dan Representasi Visual
        if len(next_as_list) > 0:
            total_found = len(next_as_list)
            st.info(f"Angka As **{target_as}** ditemukan sebagai acuan sebanyak **{total_found} kali** pada riwayat data.")
            
            # Menghitung frekuensi kemunculan
            counts = Counter(next_as_list)
            
            # Menyusun DataFrame untuk hasil analisis
            result_df = pd.DataFrame(counts.items(), columns=["Angka As (H+1)", "Frekuensi"])
            result_df["Persentase (%)"] = (result_df["Frekuensi"] / total_found) * 100
            
            # Mengurutkan dari probabilitas tertinggi
            result_df = result_df.sort_values(by="Persentase (%)", ascending=False).reset_index(drop=True)
            
            # Format visualisasi tabel
            display_df = result_df.copy()
            display_df["Angka As (H+1)"] = display_df["Angka As (H+1)"].astype(int)
            display_df["Persentase (%)"] = display_df["Persentase (%)"].round(2).astype(str) + " %"
            
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.dataframe(display_df, use_container_width=True)
                
            with col2:
                # Membuat bar chart sederhana menggunakan elemen bawaan Streamlit
                chart_data = result_df.copy()
                chart_data["Angka As (H+1)"] = chart_data["Angka As (H+1)"].astype(str)
                chart_data = chart_data.set_index("Angka As (H+1)")
                st.bar_chart(chart_data["Persentase (%)"])
                
        else:
            st.warning("Data tidak mencukupi atau pola belum pernah terjadi pada riwayat.")
