import streamlit as st
import pandas as pd
from collections import Counter

# Konfigurasi halaman antarmuka
st.set_page_config(page_title="Advanced Scanner - 4D", layout="wide")

st.title("Aplikasi Pemindai Probabilitas 4D")
st.markdown("Menggali persentase kemunculan angka H+1 berdasarkan posisi (As, Kop, Kepala, Ekor).")

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
                # Memastikan nilai pada hari tersebut memiliki data
                if pd.notna(row[cols[0]]): 
                    # Menyimpan susunan: [As, Kop, Kepala, Ekor]
                    daily_data.append([row[cols[0]], row[cols[1]], row[cols[2]], row[cols[3]]])
            except KeyError:
                continue 

    st.success(f"Basis data siap. Total riwayat hari terekstrak: {len(daily_data)} hari.")
    st.divider()
    
    # 3. Parameter Pemindaian Dinamis
    st.subheader("Parameter Analisis")
    
    # Kamus pemetaan posisi ke indeks array (0: As, 1: Kop, 2: Kepala, 3: Ekor)
    posisi_dict = {"As": 0, "Kop": 1, "Kepala": 2, "Ekor": 3}
    
    col_param1, col_param2 = st.columns(2)
    with col_param1:
        pilihan_posisi = st.selectbox("Pilih Posisi yang ingin dipindai:", list(posisi_dict.keys()))
        indeks_posisi = posisi_dict[pilihan_posisi]
        
    with col_param2:
        target_angka = st.number_input(f"Masukkan angka {pilihan_posisi} acuan (0-9):", min_value=0, max_value=9, value=4, step=1)
    
    if st.button("Jalankan Scanner", type="primary"):
        next_angka_list = []
        
        # 4. Mesin Pencarian Pola H+1
        # Loop dibatasi sampai len - 1 karena kita butuh memeriksa hari setelahnya
        for i in range(len(daily_data) - 1): 
            # Menggunakan indeks_posisi yang dinamis berdasarkan pilihan dropdown
            current_angka = daily_data[i][indeks_posisi]
            if current_angka == target_angka:
                # Ambil array 4D pada indeks hari berikutnya (H+1)
                next_day = daily_data[i+1]
                # Ekstrak angka pada posisi yang sama untuk hari esoknya
                next_angka = next_day[indeks_posisi]
                
                # Memastikan nilai tersebut adalah angka (bukan NaN)
                if pd.notna(next_angka):
                    next_angka_list.append(next_angka)
                
        # 5. Kalkulasi dan Representasi Visual
        if len(next_angka_list) > 0:
            total_found = len(next_angka_list)
            st.info(f"Angka {pilihan_posisi} **{target_angka}** ditemukan sebagai acuan sebanyak **{total_found} kali** pada riwayat data.")
            
            # Menghitung frekuensi kemunculan
            counts = Counter(next_angka_list)
            
            # Menyusun DataFrame untuk hasil analisis
            result_df = pd.DataFrame(counts.items(), columns=[f"Angka {pilihan_posisi} (H+1)", "Frekuensi"])
            result_df["Persentase (%)"] = (result_df["Frekuensi"] / total_found) * 100
            
            # Mengurutkan dari probabilitas tertinggi
            result_df = result_df.sort_values(by="Persentase (%)", ascending=False).reset_index(drop=True)
            
            # Format visualisasi tabel
            display_df = result_df.copy()
            display_df[f"Angka {pilihan_posisi} (H+1)"] = display_df[f"Angka {pilihan_posisi} (H+1)"].astype(int)
            display_df["Persentase (%)"] = display_df["Persentase (%)"].round(2).astype(str) + " %"
            
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.dataframe(display_df, use_container_width=True)
                
            with col2:
                # Membuat bar chart menggunakan elemen bawaan Streamlit
                chart_data = result_df.copy()
                chart_data[f"Angka {pilihan_posisi} (H+1)"] = chart_data[f"Angka {pilihan_posisi} (H+1)"].astype(str)
                chart_data = chart_data.set_index(f"Angka {pilihan_posisi} (H+1)")
                st.bar_chart(chart_data["Persentase (%)"])
                
        else:
            st.warning("Data tidak mencukupi atau pola belum pernah terjadi pada riwayat.")
