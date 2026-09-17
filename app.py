import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import io
from collections import Counter

# Konfigurasi halaman antarmuka Streamlit
st.set_page_config(page_title="Advanced Scanner 4D - Excel Highlighter", layout="wide")

st.title("Aplikasi Pemindai Probabilitas 4D & Visualizer Excel")
st.markdown("Menggali persentase kemunculan angka H+1 serta mengunduh file Excel dengan penanda warna otomatis.")

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
                if pd.notna(row[cols[0]]): 
                    daily_data.append([row[cols[0]], row[cols[1]], row[cols[2]], row[cols[3]]])
            except KeyError:
                continue 

    st.success(f"Basis data siap. Total riwayat hari terekstrak: {len(daily_data)} hari.")
    st.divider()
    
    # 3. Parameter Pemindaian Dinamis
    st.subheader("Parameter Analisis")
    
    posisi_dict = {"As": 0, "Kop": 1, "Kepala": 2, "Ekor": 3}
    
    col_param1, col_param2 = st.columns(2)
    with col_param1:
        pilihan_posisi = st.selectbox("Pilih Posisi yang ingin dipindai:", list(posisi_dict.keys()))
        indeks_posisi = posisi_dict[pilihan_posisi]
        
    with col_param2:
        target_angka = st.number_input(f"Masukkan angka {pilihan_posisi} acuan (0-9):", min_value=0, max_value=9, value=4, step=1)
    
    if st.button("Jalankan Scanner & Buat Excel Highlight", type="primary"):
        next_angka_list = []
        
        # 4. Mesin Pencarian Pola H+1
        for i in range(len(daily_data) - 1): 
            current_angka = daily_data[i][indeks_posisi]
            if current_angka == target_angka:
                next_day = daily_data[i+1]
                next_angka = next_day[indeks_posisi]
                
                if pd.notna(next_angka):
                    next_angka_list.append(next_angka)
                
        # 5. Kalkulasi dan Tampilan Statistik
        if len(next_angka_list) > 0:
            total_found = len(next_angka_list)
            st.info(f"Angka {pilihan_posisi} **{target_angka}** ditemukan sebagai acuan sebanyak **{total_found} kali** pada riwayat data.")
            
            # Hitung Frekuensi
            counts = Counter(next_angka_list)
            result_df = pd.DataFrame(counts.items(), columns=[f"Angka {pilihan_posisi} (H+1)", "Frekuensi"])
            result_df["Persentase (%)"] = (result_df["Frekuensi"] / total_found) * 100
            result_df = result_df.sort_values(by="Persentase (%)", ascending=False).reset_index(drop=True)
            
            display_df = result_df.copy()
            display_df[f"Angka {pilihan_posisi} (H+1)"] = display_df[f"Angka {pilihan_posisi} (H+1)"].astype(int)
            display_df["Persentase (%)"] = display_df["Persentase (%)"].round(2).astype(str) + " %"
            
            col1, col2 = st.columns([1, 2])
            with col1:
                st.dataframe(display_df, use_container_width=True)
            with col2:
                chart_data = result_df.copy()
                chart_data[f"Angka {pilihan_posisi} (H+1)"] = chart_data[f"Angka {pilihan_posisi} (H+1)"].astype(str)
                chart_data = chart_data.set_index(f"Angka {pilihan_posisi} (H+1)")
                st.bar_chart(chart_data["Persentase (%)"])
                
            st.divider()
            
            # 6. Pembuatan File Excel Berwarna (Highlighting via OpenPyXL)
            st.subheader("📥 Unduh Hasil Excel Terwarnai")
            st.caption("Keterangan Warna pada File Excel:")
            st.markdown("- 🟡 **Warna Kuning**: Kotak angka acuan yang terpilih.")
            st.markdown("- 🟢 **Warna Hijau**: Kotak angka keluaran pada hari berikutnya (H+1).")
            
            # Memuat workbook asli
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
            ws = wb.active
            
            fill_target = PatternFill(start_color="FFEB3B", end_color="FFEB3B", fill_type="solid")  # Kuning
            fill_outcome = PatternFill(start_color="A5D6A7", end_color="A5D6A7", fill_type="solid") # Hijau
            
            days_start_col = [1, 6, 11, 16, 21, 26, 31]
            
            # Mengumpulkan koordint sel (row, col) harian
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
                if daily_cells[i]['values'][indeks_posisi] == target_angka:
                    # Sel Acuan Hari Ini -> Kuning
                    tr, tc = daily_cells[i]['cells'][indeks_posisi]
                    ws.cell(row=tr, column=tc).fill = fill_target
                    
                    # Sel Keluaran H+1 -> Hijau
                    or_, oc = daily_cells[i+1]['cells'][indeks_posisi]
                    ws.cell(row=or_, column=oc).fill = fill_outcome
            
            # Simpan workbook ke memory buffer
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_bytes = excel_buffer.getvalue()
            
            # Tombol Download
            filename = f"Analisis_{pilihan_posisi}_Angka_{target_angka}.xlsx"
            st.download_button(
                label=f"📄 Unduh Excel ({filename})",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        else:
            st.warning("Data tidak mencukupi atau pola belum pernah terjadi pada riwayat.")
