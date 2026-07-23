import streamlit as st
import openpyxl
import io
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from collections import Counter

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
            
            if i > 0:
                prev_d_idx = (idx_day0 - (i - 1)) % 7
                if prev_d_idx == 0 and d_idx == 6:
                    current_r -= 1
            
            c_start = start_cols[d_idx]
            
            vals = [ws.cell(row=current_r, column=c_start+j).value for j in range(4)]
            if any(v is not None for v in vals):
                auto_val = "".join([str(clean_int(v)) if clean_int(v) is not None else "0" for v in vals])
            else:
                auto_val = "0000"
                
            with cols[i]:
                user_val = st.text_input(f"{hari_tabel[d_idx]} (-{i}):", value=auto_val, max_chars=4)
                inputs.append(user_val)
                update_targets.append((current_r, c_start))

        # Live Update Worksheet Memori
        for i, user_val in enumerate(inputs):
            r_target, c_start = update_targets[i]
            user_val = user_val.ljust(4, '0')[:4] 
            
            for offset in range(4):
                try:
                    ws.cell(row=r_target, column=c_start + offset).value = int(user_val[offset])
                except ValueError:
                    pass

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
            
            predictions_raw = {0: [], 1: [], 2: [], 3: []}
            prediction_cells = set()

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
                                
                                # Proyeksi ke hari esok
                                r_next = r_start if mode == "Lurus" else (r_start + 1 if mode == "Naik" else r_start - 1)
                                
                                if 1 <= r_next <= ws.max_row:
                                    c_next_day_idx = (idx_day0 + 1) % 7
                                    c_next = start_cols[c_next_day_idx] + pos_offset
                                    
                                    pred_val = clean_int(ws.cell(row=r_next, column=c_next).value)
                                    if pred_val is not None:
                                        predictions_raw[pos_offset].append({"val": pred_val, "length": length})
                                        prediction_cells.add((r_next, c_next))
                                        
                                break 

            # [REVISI LOGIKA] Angka Kuat Tunggal & Cadangan Tunggal
            prediction_results = {}
            for p in range(4):
                preds = predictions_raw[p]
                if not preds: continue
                
                # 1. Mengumpulkan statistik dukungan untuk setiap angka yang diprediksi
                angka_stats = {}
                for x in preds:
                    v = x['val']
                    l = x['length']
                    
                    if v not in angka_stats:
                        angka_stats[v] = {'long_count': 0, 'short_count': 0, 'total': 0, 'max_len': 0}
                    
                    angka_stats[v]['total'] += 1
                    
                    # Identifikasi apakah dari pola panjang (>=4) atau pendek (3)
                    if l >= 4:
                        angka_stats[v]['long_count'] += 1
                    else:
                        angka_stats[v]['short_count'] += 1
                        
                    # Simpan panjang baris maksimal untuk angka ini
                    if l > angka_stats[v]['max_len']:
                        angka_stats[v]['max_len'] = l

                kuat_candidates = []
                cadangan_candidates = []
                all_vals = [x['val'] for x in preds]
                
                # 2. Menyaring angka berdasarkan Syarat Kuat:
                # - Ditunjuk oleh pola panjang (4-6 hari), ATAU
                # - Ditunjuk oleh pola pendek yang ditunjang pola panjang, ATAU
                # - Ditunjuk oleh pola pendek yang menghasilkan arah/angka sama (frekuensi > 1)
                for v, stats in angka_stats.items():
                    if stats['long_count'] > 0 or stats['short_count'] > 1:
                        kuat_candidates.append(v)
                    else:
                        cadangan_candidates.append(v)
                
                # 3. Mengurutkan pemenang berdasarkan frekuensi total terbanyak, lalu panjang pola
                kuat_candidates.sort(key=lambda x: (angka_stats[x]['total'], angka_stats[x]['max_len']), reverse=True)
                cadangan_candidates.sort(key=lambda x: (angka_stats[x]['total'], angka_stats[x]['max_len']), reverse=True)
                
                angka_kuat = [kuat_candidates[0]] if kuat_candidates else []
                
                # 4. Penentuan Angka Cadangan:
                # Prioritas 1: Runner-up dari kandidat kuat (jika ada > 1 kandidat kuat)
                # Prioritas 2: Angka dari cadangan murni
                angka_cadangan = []
                if len(kuat_candidates) > 1:
                    angka_cadangan = [kuat_candidates[1]]
                elif cadangan_candidates:
                    angka_cadangan = [cadangan_candidates[0]]
                    
                kuat_max_len = angka_stats[angka_kuat[0]]['max_len'] if angka_kuat else 0
                
                prediction_results[p] = {
                    "kuat": angka_kuat,
                    "cadangan": angka_cadangan,
                    "max_len": kuat_max_len,
                    "all_counts": Counter(all_vals)
                }

            st.session_state.highlighted = cell_patterns
            st.session_state.stats = total_stats
            st.session_state.prediction_results = prediction_results
            st.session_state.prediction_cells = prediction_cells
            st.session_state.scanned = True
            st.rerun()

        # 5. OUTPUT
        if st.session_state.get("scanned"):
            st.divider()
            
            excel_buffer = generate_excel(ws, st.session_state.get("highlighted", {}))
            st.download_button(
                label="📥 Download Hasil Scan (.xlsx)",
                data=excel_buffer,
                file_name="hasil_scan_paito.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # UI Prediksi Angka Kuat & Cadangan Tunggal
            st.subheader("🎯 Prediksi Hari Berikutnya")
            pos_names = ["As", "Kop", "Kepala", "Ekor"]
            
            pred_cols = st.columns(4)
            for p in range(4):
                with pred_cols[p]:
                    st.markdown(f"**Posisi {pos_names[p]}**")
                    if p in st.session_state.get("prediction_results", {}):
                        res = st.session_state.prediction_results[p]
                        
                        kuat_str = str(res['kuat'][0]) if res['kuat'] else "-"
                        cadangan_str = str(res['cadangan'][0]) if res['cadangan'] else "-"
                        
                        pola_info = f"*(Pola {res['max_len']} Baris)*" if res['kuat'] else "*(Tidak Ada Kandidat Kuat)*"
                        st.success(f"🔥 **Kuat:** {kuat_str}\n\n{pola_info}")
                        st.info(f"🛡️ **Cadangan:** {cadangan_str}")
                        
                        with st.expander("Detail Frekuensi (Semua Pola)"):
                            for val, count in res['all_counts'].most_common():
                                status = " (Kuat)" if val in res['kuat'] else (" (Cadangan Utama)" if res['cadangan'] and val == res['cadangan'][0] else "")
                                st.write(f"Angka {val}: didukung {count} jalur{status}")
                    else:
                        st.write("Belum ada pola")
            
            st.divider()
            
            st.subheader("Statistik Jalur Pola")
            stats = st.session_state.stats
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Pola 6 Hari", f"{stats[6]} Jalur")
            c2.metric("Pola 5 Hari", f"{stats[5]} Jalur")
            c3.metric("Pola 4 Hari", f"{stats[4]} Jalur")
            c4.metric("Pola 3 Hari", f"{stats[3]} Jalur")

            st.subheader("Live Preview Grid")
            highlighted = st.session_state.get("highlighted", {})
            prediction_cells = st.session_state.get("prediction_cells", set())
            
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
                        
                        is_pred = (r, c_idx) in prediction_cells
                        bg = "#ffffff"
                        
                        if (r, c_idx) in highlighted:
                            p = highlighted[(r, c_idx)]["pos"]
                            colors = {0: "#3399FF", 1: "#D2B48C", 2: "#22C55E", 3: "#FFD700"}
                            bg = colors.get(p, "#ffffff")
                        elif is_pred:
                            bg = "#fee2e2" 
                        
                        border_style = "2px solid #dc2626" if is_pred else "1px solid #ccc"
                        text_color = "#dc2626" if is_pred else "inherit"
                        
                        html.append(f"<td style='border: {border_style}; background-color: {bg}; color: {text_color}; font-weight: bold; width: 25px; height: 25px;'>{display_val}</td>")
                    
                    html.append("<td style='width: 15px;'></td>")
                    
                html.append("</tr>")
            html.append("</table></div>")
            st.markdown("".join(html), unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")
