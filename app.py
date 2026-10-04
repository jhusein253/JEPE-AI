import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import io
from collections import Counter

# Konfigurasi halaman antarmuka Streamlit
st.set_page_config(page_title="Advanced Scanner 4D - Angka Kuat", layout="wide")

st.title("Aplikasi Pemindai Probabilitas 4D & Visualizer Excel")
st.markdown("Menggali kemunculan angka H+1 serta mengidentifikasi **🔥 Angka Kuat** (frekuensi muncul $\ge$ 5 kali).")

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
            pilihan_kombinasi = st.selectbox("Pilih Kombinasi Posisi:", list(kombinasi_dict.keys()))
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
        
        # 4. Mesin Pencarian Pola H+1
        for i in range(len(daily_data) - 1):
            if mode_analisis == "1 Digit (Tunggal)":
                if daily_data[i][indeks_posisi[0]] == target_1:
                    next_val = daily_data[i+1][indeks_posisi[0]]
                    next_angka_list.append(str(next_val))
            else:
                if daily_data[i][indeks_posisi[0]] == target_1 and daily_data[i][indeks_posisi[1]] == target_2:
                    v1 = daily_data[i+1][indeks_posisi[0]]
                    v2 = daily_data[i+1][indeks_posisi[1]]
                    next_angka_list.append(f"{v1}{v2}")
                    
        # 5. Kalkulasi dan Tampilan Statistik
        if len(next_angka_list) > 0:
            total_found = len(next_angka_list)
            
            # Hitung Frekuensi
            counts = Counter(next_angka_list)
            result_df = pd.DataFrame(counts.items(), columns=[f"Hasil {pilihan_posisi} (H+1)", "Frekuensi"])
            result_df["Persentase (%)"] = (result_df["Frekuensi"] / total_found) * 100
            # Mengurutkan dari frekuensi terbanyak
            result_df = result_df.sort_values(by="Frekuensi", ascending=False).reset_index(drop=True)
            
            angka_kuat_list = []
            
            # Tampilan Khusus Mode 2 Digit (Angka Kuat)
            if mode_analisis == "2 Digit Kombinasi (As-Kop / Kepala-Ekor)":
                # Filter khusus angka yang frekuensinya >= 5
                angka_kuat_df = result_df[result_df["Frekuensi"] >= 5].copy()
                angka_kuat_list = angka_kuat_df[f"Hasil {pilihan_posisi} (H+1)"].astype(str).tolist()
                
                st.subheader("🔥 Identifikasi Angka Kuat (Minimal 5x Muncul)")
                if not angka_kuat_df.empty:
                    st.success(f"Ditemukan **{len(angka_kuat_df)} pasang Angka Kuat** dari acuan **{target_str}**!")
                    
                    display_kuat_df = angka_kuat_df.copy()
                    display_kuat_df["Persentase (%)"] = display_kuat_df["Persentase (%)"].round(2).astype(str) + " %"
                    st.dataframe(display_kuat_df, use_container_width=True)
                else:
                    st.warning(f"Riwayat acuan {target_str} belum memiliki Angka Kuat (tidak ada H+1 yang muncul \u2265 5 kali).")
                
                st.divider()
                st.subheader("📊 Seluruh Riwayat Keluaran H+1")
            else:
                st.info(f"Angka {pilihan_posisi} **{target_str}** ditemukan sebagai acuan sebanyak **{total_found} kali** pada riwayat data.")
            
            # Tabel dan Grafik Keseluruhan
            display_df = result_df.copy()
            display_df["Persentase (%)"] = display_df["Persentase (%)"].round(2).astype(str) + " %"
            
            col1, col2 = st.columns([1, 2])
            with col1:
                st.dataframe(display_df, use_container_width=True)
            with col2:
                chart_data = result_df.copy().set_index(f"Hasil {pilihan_posisi} (H+1)")
                st.bar_chart(chart_data["Persentase (%)"])
                
            st.divider()
            
            # 6. Pembuatan File Excel Berwarna (Highlighting via OpenPyXL)
            st.subheader("📥 Unduh Hasil Excel Terwarnai")
            st.caption("Keterangan Warna pada File Excel:")
            st.markdown("- 🟡 **Warna Kuning**: Kotak angka acuan (Hari Ini) yang terpilih.")
            if mode_analisis == "2 Digit Kombinasi (As-Kop / Kepala-Ekor)":
                st.markdown("- 🟢 **Warna Hijau**: Kotak angka keluaran H+1 **(HANYA untuk Angka Kuat yang muncul $\ge$ 5 kali)**.")
            else:
                st.markdown("- 🟢 **Warna Hijau**: Kotak angka keluaran H+1.")
            
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
                    # Validasi apakah keluaran H+1 layak diwarnai hijau
                    is_angka_kuat = False
                    if mode_analisis == "1 Digit (Tunggal)":
                        is_angka_kuat = True # Mode 1 digit semua diwarnai
                    else:
                        outcome_str = f"{daily_cells[i+1]['values'][indeks_posisi[0]]}{daily_cells[i+1]['values'][indeks_posisi[1]]}"
                        if outcome_str in angka_kuat_list:
                            is_angka_kuat = True # Mode 2 digit hanya hijau jika >= 5 kali
                            
                    # Warnai Kuning untuk Acuan
                    for idx in indeks_posisi:
                        tr, tc = daily_cells[i]['cells'][idx]
                        ws.cell(row=tr, column=tc).fill = fill_target
                    
                    # Warnai Hijau untuk H+1 (Jika Lolos Validasi)
                    if is_angka_kuat:
                        for idx in indeks_posisi:
                            or_, oc = daily_cells[i+1]['cells'][idx]
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
