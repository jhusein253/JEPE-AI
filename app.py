import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import io
from collections import Counter

# Konfigurasi halaman antarmuka Streamlit
st.set_page_config(page_title="Advanced Scanner 4D - Angka Kuat H+1", layout="wide")

st.title("Aplikasi Pemindai Probabilitas 4D & Visualizer Excel")
st.markdown("Menggali kemunculan angka di H+1 serta mengidentifikasi **🔥 Angka Kuat** (frekuensi muncul $\ge$ 5 kali).")

# 1. Komponen Unggah File
uploaded_file = st.file_uploader("Unggah file Excel referensi (misal: HK.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Membaca data menggunakan pandas
    file_bytes = uploaded_file.getvalue()
    df = pd.read_excel(io.BytesIO(file_bytes))
    
    # Membersihkan baris pertama jika berupa baris pemisah/kosong
    if df.iloc[0].isna().all():
        df = df.iloc[1:].reset_index(drop=True)
        
    # 2. Algoritma perataan (flattening) matriks ke urutan kronologis harian
    daily_data = []
    days_cols = {
        'SABTU': ['SABTU', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3'],
        'MINGGU': ['MINGGU', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8'],
        'SENIN': ['SENIN', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 13'],
        'SELASA': ['SELASA', 'Unnamed: 16', 'Unnamed: 17', 'Unnamed: 18'],
        'RABU': ['RABU', 'Unnamed: 21', 'Unnamed: 22', 'Unnamed: 23'],
        'KAMIS': ['KAMIS', 'Unnamed: 26', 'Unnamed: 27', 'Unnamed: 28'],
        'JUMAT': ['JUMAT', 'Unnamed: 31', 'Unnamed: 32', 'Unnamed: 33']
    }

    for index, row in df.iterrows():
        for day, cols in days_cols.items():
            try:
                if pd.notna(row[cols[0]]): 
                    daily_data.append([int(row[cols[0]]), int(row[cols[1]]), int(row[cols[2]]), int(row[cols[3]])])
            except KeyError:
                continue 

    st.success(f"Basis data siap. Total riwayat hari terekstrak: {len(daily_data)} hari.")
    st.divider()
    
    # 3. Parameter Pemindaian Dinamis
    st.subheader("Parameter Analisis")
    
    mode_analisis = st.radio("Pilih Mode Analisis:", ["1 Digit (Tunggal)", "2 Digit Kombinasi (As-Kop / Kepala-Ekor)"], horizontal=True)
    
    if mode_analisis == "1 Digit (Tunggal)":
        posisi_dict = {"As": 0, "Kop": 1, "Kepala": 2, "Ekor": 3}
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            pilihan_posisi = st.selectbox("Pilih Posisi:", list(posisi_dict.keys()))
            indeks_posisi = [posisi_dict[pilihan_posisi]]
        with col_p2:
            target_1 = st.number_input(f"Masukkan angka {pilihan_posisi} acuan (0-9):", min_value=0, max_value=9, value=4, step=1)
            target_str = str(target_1)
    else:
        kombinasi_dict = {"As - Kop": [0, 1], "Kepala - Ekor": [2, 3]}
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            pilihan_kombinasi = st.selectbox("Pilih Kombinasi Posisi Acuan:", list(kombinasi_dict.keys()))
            indeks_posisi = kombinasi_dict[pilihan_kombinasi]
            nama1, nama2 = pilihan_kombinasi.split(" - ")
        with col_p2:
            target_1 = st.number_input(f"Angka {nama1} (0-9):", min_value=0, max_value=9, value=9, step=1)
        with col_p3:
            target_2 = st.number_input(f"Angka {nama2} (0-9):", min_value=0, max_value=9, value=4, step=1)
            target_str = f"{target_1}{target_2}"
            pilihan_posisi = pilihan_kombinasi

    if st.button("Jalankan Scanner & Buat Excel Highlight", type="primary"):
        next_angka_list = []
        hari_ditemukan = 0
        
        # 4. Mesin Pencarian Pola H+1
        for i in range(len(daily_data) - 1):
            if mode_analisis == "1 Digit (Tunggal)":
                # Cari 1 posisi, ambil 1 posisi di H+1
                if daily_data[i][indeks_posisi[0]] == target_1:
                    hari_ditemukan += 1
                    next_val = daily_data[i+1][indeks_posisi[0]]
                    next_angka_list.append(str(next_val))
            else:
                # Cari Kombinasi, ambil SEMUA posisi (4 digit) di H+1
                if daily_data[i][indeks_posisi[0]] == target_1 and daily_data[i][indeks_posisi[1]] == target_2:
                    hari_ditemukan += 1
                    for digit in daily_data[i+1]: # Ekstrak As, Kop, Kepala, Ekor dari H+1
                        next_angka_list.append(str(digit))
                        
        # 5. Kalkulasi dan Tampilan Statistik
        if len(next_angka_list) > 0:
            total_found = len(next_angka_list)
            
            # Hitung Frekuensi
            counts = Counter(next_angka_list)
            
            # Tentukan judul kolom berdasarkan mode
            kolom_hasil = f"Hasil {pilihan_posisi} (H+1)" if mode_analisis == "1 Digit (Tunggal)" else "Angka H+1 (Semua Posisi)"
            
            result_df = pd.DataFrame(counts.items(), columns=[kolom_hasil, "Frekuensi"])
            result_df["Persentase (%)"] = (result_df["Frekuensi"] / total_found) * 100
            # Mengurutkan dari frekuensi terbanyak
            result_df = result_df.sort_values(by="Frekuensi", ascending=False).reset_index(drop=True)
            
            angka_kuat_list = []
            
            st.info(f"Pola acuan **{target_str}** ditemukan pada **{hari_ditemukan} hari** dalam riwayat.")
            
            # Tampilan Khusus Mode 2 Digit (Angka Kuat di semua posisi)
            if mode_analisis == "2 Digit Kombinasi (As-Kop / Kepala-Ekor)":
                # Filter khusus angka tunggal yang frekuensinya >= 5
                angka_kuat_df = result_df[result_df["Frekuensi"] >= 5].copy()
                angka_kuat_list = angka_kuat_df[kolom_hasil].astype(str).tolist()
                
                st.subheader("🔥 Identifikasi Angka Kuat (Minimal 5x Muncul di H+1)")
                if not angka_kuat_df.empty:
                    angka_kuat_str = ", ".join(angka_kuat_list)
                    st.success(f"Ditemukan Angka Kuat: **{angka_kuat_str}** (Angka-angka ini akan di-highlight hijau pada Excel).")
                    
                    display_kuat_df = angka_kuat_df.copy()
                    display_kuat_df["Persentase (%)"] = display_kuat_df["Persentase (%)"].round(2).astype(str) + " %"
                    st.dataframe(display_kuat_df, use_container_width=True)
                else:
                    st.warning(f"Tidak ada angka yang memenuhi syarat Angka Kuat (\u2265 5 kali muncul).")
                
                st.divider()
                st.subheader("📊 Statistik Seluruh Angka di H+1")
            
            # Tabel dan Grafik Keseluruhan
            display_df = result_df.copy()
            display_df["Persentase (%)"] = display_df["Persentase (%)"].round(2).astype(str) + " %"
            
            col1, col2 = st.columns([1, 2])
            with col1:
                st.dataframe(display_df, use_container_width=True)
            with col2:
                chart_data = result_df.copy().set_index(kolom_hasil)
                st.bar_chart(chart_data["Frekuensi"]) # Tampilkan grafik berdasarkan jumlah frekuensi
                
            st.divider()
            
            # 6. Pembuatan File Excel Berwarna (Highlighting via OpenPyXL)
            st.subheader("📥 Unduh Hasil Excel Terwarnai")
            st.caption("Keterangan Warna pada File Excel:")
            st.markdown("- 🟡 **Warna Kuning**: Kotak angka acuan (Hari Ini) yang terpilih.")
            if mode_analisis == "2 Digit Kombinasi (As-Kop / Kepala-Ekor)":
                st.markdown("- 🟢 **Warna Hijau**: Menandai **Angka Kuat** di baris H+1 (bisa di posisi As, Kop, Kepala, maupun Ekor).")
            else:
                st.markdown("- 🟢 **Warna Hijau**: Kotak angka keluaran H+1 pada posisi yang sama.")
            
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
            ws = wb.active
            
            fill_target = PatternFill(start_color="FFEB3B", end_color="FFEB3B", fill_type="solid")  # Kuning
            fill_outcome = PatternFill(start_color="A5D6A7", end_color="A5D6A7", fill_type="solid") # Hijau
            
            days_start_col = [1, 6, 11, 16, 21, 26, 31]
            
            daily_cells = []
            for r in range(1, ws.max_row + 1):
                for sc in days_start_col:
                    v_as = ws.cell(row=r, column=sc).value
                    v_kop = ws.cell(row=r, column=sc+1).value
                    v_kep = ws.cell(row=r, column=sc+2).value
                    v_eko = ws.cell(row=r, column=sc+3).value
                    if v_as is not None and isinstance(v_as, (int, float)):
                        daily_cells.append({
                            'cells': [(r, sc), (r, sc+1), (r, sc+2), (r, sc+3)],
                            'values': [int(v_as), int(v_kop), int(v_kep), int(v_eko)]
                        })
            
            # Mewarnai sel acuan dan sel H+1
            for i in range(len(daily_cells) - 1):
                match = False
                if mode_analisis == "1 Digit (Tunggal)":
                    if daily_cells[i]['values'][indeks_posisi[0]] == target_1:
                        match = True
                else:
                    if daily_cells[i]['values'][indeks_posisi[0]] == target_1 and daily_cells[i]['values'][indeks_posisi[1]] == target_2:
                        match = True
                        
                if match:
                    # 1. Warnai Kuning untuk Acuan (Hari Ini)
                    for idx in indeks_posisi:
                        tr, tc = daily_cells[i]['cells'][idx]
                        ws.cell(row=tr, column=tc).fill = fill_target
                    
                    # 2. Warnai Hijau untuk H+1
                    if mode_analisis == "1 Digit (Tunggal)":
                        # Warnai posisi yang sama saja
                        or_, oc = daily_cells[i+1]['cells'][indeks_posisi[0]]
                        ws.cell(row=or_, column=oc).fill = fill_outcome
                    else:
                        # Mode 2 Digit: Cek ke-4 posisi di H+1, warnai hijau jika itu adalah Angka Kuat
                        for pos_idx in range(4):
                            digit_h1 = str(daily_cells[i+1]['values'][pos_idx])
                            if digit_h1 in angka_kuat_list:
                                or_, oc = daily_cells[i+1]['cells'][pos_idx]
                                ws.cell(row=or_, column=oc).fill = fill_outcome
            
            # Simpan workbook ke memory buffer
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_bytes = excel_buffer.getvalue()
            
            # Tombol Download
            filename = f"Analisis_AngkaKuat_{pilihan_posisi}_{target_str}.xlsx".replace(" ", "_")
            st.download_button(
                label=f"📄 Unduh Excel ({filename})",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        else:
            st.warning("Data tidak mencukupi atau pola belum pernah terjadi pada riwayat.")
