import streamlit as st
import openpyxl
import io
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# Setup halaman
st.set_page_config(page_title="JEPE AI Pro", layout="wide")
st.title("JEPE AI - Advanced Scanner")

# Fungsi pembersih data
def clean_int(v):
    try: return int(float(str(v).strip()))
    except (ValueError, TypeError): return None

# Fungsi generate_excel dengan lebar kolom 3
def generate_excel(original_ws, highlighted_data):
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    
    # Atur lebar kolom menjadi 3
    for col_num in range(1, original_ws.max_column + 1):
        col_letter = get_column_letter(col_num)
        new_ws.column_dimensions[col_letter].width = 3
    
    # Warna yang sama dengan UI
    colors = {0: "3399FF", 1: "D2B48C", 2: "22C55E", 3: "FFD700"}
    
    # Copy data dan terapkan warna
    for r in range(1, original_ws.max_row + 1):
        for c in range(1, original_ws.max_column + 1):
            cell_val = original_ws.cell(row=r, column=c).value
            new_ws.cell(row=r, column=c).value = cell_val
            
            if (r, c) in highlighted_data:
                pos = highlighted_data[(r, c)]["pos"]
                hex_color = colors.get(pos, "FFFF00")
                new_ws.cell(row=r, column=c).fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    
    buf = io.BytesIO()
    new_wb.save(buf)
    buf.seek(0)
    return buf

# 1. UPLOAD FILE
uploaded_file = st.file_uploader("Unggah Database Paito (.xlsx):", type=["xlsx"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        
        hari_tabel = ["Sabtu", "Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat"]
        start_cols = [1, 6, 11, 16, 21, 26, 31]

        # 2. INPUT REFERENSI (Auto-Fetch & Live Update)
        st.header("2. Input Referensi (Auto-Fetch & Live Update)")
        
        # [REVISI]: Tambahan kontrol baris target untuk mencegah penumpukan data ke minggu lalu
        c_opt1, c_opt2 = st.columns(2)
        with c_opt1:
            hari_terpilih = st.selectbox("Pilih Hari Utama (Hari Ini):", hari_tabel, index=4)
        with c_opt2:
            target_row_utama = st.number_input("Baris Target (Default: Baris Terakhir)", min_value=1, value=ws.max_row)

        idx_day0 = hari_tabel.index(hari_terpilih)

        inputs = []
        update_targets = []
        
        cols = st.columns(6)
        
        current_r = target_row_utama
        for i in range(6):
            d_idx = (idx_day0 - i) % 7
            
            # Jika mundur dari Sabtu (0) ke Jumat (6), berarti mundur 1 minggu (baris naik ke atas/berkurang 1)
            if i > 0:
                prev_d_idx = (idx_day0 - (i - 1)) % 7
                if prev_d_idx == 0 and d_idx == 6:
                    current_r -= 1
            
            c_start = start_cols[d_idx]
            
            # Ambil data HANYA dari baris target yang benar. 
            vals = [ws.cell(row=current_r, column=c_start+j).value for j in range(4)]
            if any(v is not None for v in vals):
                auto_val = "".join([str(clean_int(v)) if clean_int(v) is not None else "0" for v in vals])
            else:
                auto_val = "0000" # Kosongkan jika memang belum ada data di minggu/baris ini
                
            with cols[i]:
                user_val = st.text_input(f"{hari_tabel[d_idx]} (-{i}):", value=auto_val, max_chars=4)
                inputs.append(user_val)
                update_targets.append((current_r, c_start))

        # Live Update Worksheet Memori
        # Menimpa data di memori Excel dengan inputan user saat ini secara otomatis
        for i, user_val in enumerate(inputs):
            r_target, c_start = update_targets[i]
            
            # Pastikan teks berjumlah 4 karakter (jika kurang, ditambah '0' di belakang)
            user_val = user_val.ljust(4, '0')[:4] 
            
            for offset in range(4):
                try:
                    # Update nilai sel secara live di memori ke baris yang tepat
                    ws.cell(row=r_target, column=c_start + offset).value = int(user_val[offset])
                except ValueError:
                    pass # Abaikan jika input bukan angka

        # 3. PENGATURAN
        st.divider()
        c_lurus = st.checkbox("Garis Lurus", value=True)
        c_naik = st.checkbox("Diagonal Naik", value=True)
        c_turun = st.checkbox("Diagonal Turun", value=True)
        
        use_single_ref = st.checkbox("Mode Acuan Posisi Tunggal", value=False)
        ref_pos_name = st.selectbox("Posisi Acuan:", ["As", "Kop", "Kepala", "Ekor"], index=0, disabled=not use_single_ref)
        ref_pos_offset = ["As", "Kop", "Kepala", "Ekor"].index(ref_pos_name)

        # 4. LOGIKA SCANNING
        if st.button("JALANKAN ANALISA"):
            cell_patterns = {}
            total_stats = {6: 0, 5: 0, 4: 0, 3: 0}
            days_indices = [(idx_day0 - k) % 7 for k in range(6)]

            for pos_offset in range(4):
                current_allowed = []
                for k in range(6):
                    val_str = inputs[k]
                    digit = int(val_str[ref_pos_offset if use_single_ref else pos_offset])
                    current_allowed.append([digit, (digit + 5) % 10])

                for r_start in range(1, ws.max_row + 1):
                    for mode in ["Lurus", "Naik", "Turun"]:
                        if (mode == "Lurus" and not c_lurus) or (mode == "Naik" and not c_naik) or (mode == "Turun" and not c_turun): continue
                        
                        for length in [6, 5, 4, 3]:
                            path, valid = [], True
                            for k in range(length):
                                r_target = r_start if mode == "Lurus" else (r_start - k if mode == "Naik" else r_start + k)
                                if r_target < 1 or r_target > ws.max_row: valid = False; break
                                
                                cell_val = ws.cell(row=r_target, column=start_cols[days_indices[k]] + pos_offset).value
                                val = clean_int(cell_val)
                                if val not in current_allowed[k]: valid = False; break
                                path.append((r_target, start_cols[days_indices[k]] + pos_offset))
                            
                            if valid:
                                total_stats[length] += 1
                                for r_c, c_c in path: cell_patterns[(r_c, c_c)] = {"length": length, "pos": pos_offset}
                                break 

            st.session_state.highlighted = cell_patterns
            st.session_state.stats = total_stats
            st.session_state.scanned = True
            st.rerun()

        # 5. OUTPUT
        if st.session_state.get("scanned"):
            st.divider()
            
            # Tombol Download (Bebas dari emoji penyebab error encoding)
            excel_buffer = generate_excel(ws, st.session_state.get("highlighted", {}))
            st.download_button(
                label="Download Hasil Scan (.xlsx)",
                data=excel_buffer,
                file_name="hasil_scan_paito.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.subheader("Statistik")
            stats = st.session_state.stats
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Pola 6 Hari", f"{stats[6]} Jalur")
            c2.metric("Pola 5 Hari", f"{stats[5]} Jalur")
            c3.metric("Pola 4 Hari", f"{stats[4]} Jalur")
            c4.metric("Pola 3 Hari", f"{stats[3]} Jalur")

            st.subheader("Live Preview Grid")
            highlighted = st.session_state.get("highlighted", {})
            
            html = ["<div style='overflow-x: auto;'><table style='border-collapse: collapse; width: 100%; text-align: center; font-family: monospace; font-size: 12px;'>"]
            
            # Header
            html.append("<tr style='background-color: #0f172a; color: white;'><th>Line</th>")
            for h in hari_tabel:
                html.append(f"<th colspan='4'>{h}</th><th style='width: 15px;'></th>") 
            html.append("</tr>")
            
            # Data Rows
            for r in range(max(1, ws.max_row - 30), ws.max_row + 1):
                html.append(f"<tr><td style='border: 1px solid #ccc; background-color: #f0f0f0; width: 25px; height: 25px; text-align: center; font-weight: bold;'>{r}</td>")
                
                for i, start_col in enumerate(start_cols):
                    for offset in range(4):
                        c_idx = start_col + offset
                        val = ws.cell(row=r, column=c_idx).value
                        display_val = str(val) if val is not None else "-"
                        
                        bg = "#ffffff"
                        if (r, c_idx) in highlighted:
                            p = highlighted[(r, c_idx)]["pos"]
                            colors = {0: "#3399FF", 1: "#D2B48C", 2: "#22C55E", 3: "#FFD700"}
                            bg = colors.get(p, "#ffffff")
                        
                        html.append(f"<td style='border: 1px solid #ccc; background-color: {bg}; font-weight: bold; width: 25px; height: 25px;'>{display_val}</td>")
                    
                    html.append("<td style='width: 15px;'></td>")
                    
                html.append("</tr>")
            html.append("</table></div>")
            st.markdown("".join(html), unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")