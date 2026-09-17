import streamlit as st
import pandas as pd
from collections import Counter

# Konfigurasi halaman antarmuka
st.set_page_config(page_title="Advanced Scanner - 4D", layout="wide")

st.title("Aplikasi Pemindai Probabilitas 4D (As, Kop, Kepala, Ekor)")
st.markdown("Menggali persentase kemunculan angka H+1 berdasarkan posisi spesifik dari data historis Excel.")

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

    # Looping per baris mingguan, laluUntuk memproses keempat posisi angka (As, Kop, Kepala, dan Ekor) secara komprehensif dalam pemindaian matriks angka, setiap nilai 4D perlu dipecah menjadi kolom atau variabel terpisah. Pendekatan ini memungkinkan algoritma mengevaluasi pola, frekuensi, atau rumus pada masing-masing digit secara independen.

Berikut adalah logika implementasinya menggunakan Python (Pandas) untuk memisahkan posisi tersebut dari sebuah dataset:

```python
import pandas as pd

# Contoh dataset berisi angka 4D (pastikan formatnya string agar mudah diindeks)
df = pd.DataFrame({"Angka_4D": ["1234", "5678", "9012", "3456"]})

# Memecah angka ke masing-masing posisi
df['As'] = df['Angka_4D'].str[0].astype(int)      # Digit ke-1 (Ribuan)
df['Kop'] = df['Angka_4D'].str[1].astype(int)     # Digit ke-2 (Ratusan)
df['Kepala'] = df['Angka_4D'].str[2].astype(int)  # Digit ke-3 (Puluhan)
df['Ekor'] = df['Angka_4D'].str[3].astype(int)    # Digit ke-4 (Satuan)

print(df)
