import streamlit as st
import openpyxl
import io
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from collections import Counter

# Setup halaman
st.set_page_config(page_title="JEPE AI Pro - 6 Hari", layout="wide")
st.title("JEPE AI - Advanced Scanner (6 Hari)")

# Fungsi pembersih data
def clean_int(v):
    try: return int(float(str(v).strip()))
    except (ValueError, TypeError): return None

# Fungsi generate_excel
def generate_excel(original_ws, highlighted_data):
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    
    for col_num in range(1, original_ws.max_column + 1):
        col_letter = get_column_letter(col_num)
        new_ws.column_dimensions[col_letter].width = 3
    
    colors = {0: "3399FF", 1: "D2B48C", 2: "22C55E", 3: "FFD700"}
    
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

        # 2. INPUT REFERENSI
        st.header("2. Input Referensi (6 Hari)")
        
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

        for i, user_val in enumerate(inputs):
            r_target, c_start = update_targets[i]
            user_val = user_val.ljust(4, '0')[:4] 
            for offset in range(4):
                try: ws.cell(row=r_target, column=c_start + offset).value = int(user_val[offset])
                except ValueError: pass
                
        min_query_row = current_r

        # 3. PENGATURAN
        st.divider()
        st.write("**Mode Pola Scan (Baku):**")
        c_col1, c_col2, c_col3 = st.columns(3)
        with c_col1: c_lurus = st.checkbox("Garis Lurus", value=True)
        with c_col2: c_naik = st.checkbox("Diagonal Naik", value=True)
        with c_col3: c_turun = st.checkbox("Diagonal Turun", value=True)

        # 4. LOGIKA SCANNING
        if st.button("JALANKAN ANALISA"):
            cell_patterns = {}
            total_stats = {6: 0, 5: 0, 4: 0, 3: 0}
            days_indices = [(idx_day0 - k) % 7 for k in range(6)]
            
            predictions_raw = {0: [], 1: [], 2: [], 3: []}
            prediction_cells = set()
            
            max_analisa_row = min_query_row - 1

            for target_pos in range(4):
                current_allowed = []
                for k in range(6):
                    val_str = inputs[k]
                    digit = int(val_str[target_pos])
                    current_allowed.append([digit, (digit + 5) % 10])

                for search_pos in range(4):
                    for r_start in range(1, max_analisa_row + 1):
                        for mode in ["Lurus", "Naik", "Turun"]:
                            if (mode == "Lurus" and not c_lurus) or \
                               (mode == "Naik" and not c_naik) or \
                               (mode == "Turun" and not c_turun):
                                continue
                            
                            for length in [6, 5, 4, 3]:
                                path, valid = [], True
                                for k in range(length):
                                    r_target = r_start if mode == "Lurus" else (r_start - k if mode == "Naik" else r_start + k)
                                    if r_target < 1 or r_target > max_analisa_row: valid = False; break
                                    
                                    cell_val = ws.cell(row=r_target, column=start_cols[days_indices[k]] + search_pos).value
                                    val = clean_int(cell_val)
                                    if val not in current_allowed[k]: valid = False; break
                                    path.append((r_target, start_cols[days_indices[k]] + search_pos))
                                
                                if valid:
                                    total_stats[length] += 1
                                    
                                    # [REVISI]: Pewarnaan di Live Preview dan Download Excel HANYA dilakukan jika 
                                    # posisi pencarian sejajar dengan aslinya (As=As, Kop=Kop)
                                    if search_pos == target_pos:
                                        for r_c, c_c in path: 
                                            if (r_c, c_c) not in cell_patterns or cell_patterns[(r_c, c_c)]["length"] <= length:
                                                cell_patterns[(r_c, c_c)] = {"length": length, "pos": target_pos}
                                    
                                    r_next = r_start if mode == "Lurus" else (r_start + 1 if mode == "Naik" else r_start - 1)
                                    
                                    if idx_day0 == 6:
                                        r_next += 1
                                        
                                    if 1 <= r_next <= ws.max_row:
                                        c_next_day_idx = (idx_day0 + 1) % 7
                                        c_next = start_cols[c_next_day_idx] + search_pos
                                        pred_val = clean_int(ws.cell(row=r_next, column=c_next).value)
                                        if pred_val is not None:
                                            predictions_raw[target_pos].append({
                                                "val": pred_val, "length": length, "is_own_pos": (search_pos == target_pos) 
                                            })
                                            
                                            # Kotak merah prediksi juga HANYA ditandai jika posisinya sejajar
                                            if search_pos == target_pos:
                                                prediction_cells.add((r_next, c_next))
                                    break 

            prediction_results = {}
            for target_pos in range(4):
                preds = predictions_raw[target_pos]
                if not preds: continue
                max_len = max(x['length'] for x in preds)
                all_vals = [x['val'] for x in preds]
                all_counts = Counter(all_vals)
                own_pos_vals = set([x['val'] for x in preds if x['is_own_pos']])
                
                if own_pos_vals:
                    kuat_candidates = sorted(list(own_pos_vals), key=lambda val: all_counts[val], reverse=True)
                else:
                    kuat_candidates = [val for val, count in all_counts.most_common()]
                    
                angka_kuat = [kuat_candidates[0]] if kuat_candidates else []
                cadangan_candidates = [val for val, count in all_counts.most_common() if val not in angka_kuat]
                angka_cadangan = [cadangan_candidates[0]] if cadangan_candidates else []
                
                prediction_results[target_pos] = {
                    "kuat": angka_kuat, "cadangan": angka_cadangan, "max_len": max_len, "all_counts": all_counts, "own_pos_vals": own_pos_vals
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
            st.download_button("📥 Download Hasil Scan (.xlsx)", data=excel_buffer, file_name="hasil_scan_paito.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            st.subheader("🔮 Prediksi Hari Berikutnya")
            pos_names = ["As", "Kop", "Kepala", "Ekor"]
            pred_cols = st.columns(4)
            for p in range(4):
                with pred_cols[p]:
                    st.markdown(f"**Posisi {pos_names[p]}**")
                    if p in st.session_state.get("prediction_results", {}):
                        res = st.session_state.prediction_results[p]
                        kuat_str = str(res['kuat'][0]) if res['kuat'] else "-"
                        cadangan_str = str(res['cadangan'][0]) if res['cadangan'] else "-"
                        st.success(f"🔥 **Kuat:** {kuat_str}\n\n*(Pola {res['max_len']} Baris)*")
                        st.info(f"💡 **Cadangan:** {cadangan_str}")
                        with st.expander("Detail Dukungan"):
                            for val, count in res['all_counts'].most_common():
                                own_tag = " ✓*(Asli)*" if val in res['own_pos_vals'] else ""
                                status = " **[Kuat]**" if val in res['kuat'] else (" [Cadangan]" if res['cadangan'] and val == res['cadangan'][0] else "")
                                st.write(f"Angka {val}: {count} jalur{own_tag}{status}")
            
            st.divider()
            st.subheader("Statistik Jalur Pola")
            stats = st.session_state.stats
            m_cols = st.columns(4)
            for i, length in enumerate([6, 5, 4, 3]):
                m_cols[i].metric(f"Pola {length}", f"{stats[length]} Jalur")

            st.subheader("Live Preview Grid")
            highlighted = st.session_state.get("highlighted", {})
            prediction_cells = st.session_state.get("prediction_cells", set())
            
            html = ["<div style='overflow-x: auto;'><table style='border-collapse: collapse; width: 100%; text-align: center; font-family: monospace; font-size: 12px;'>"]
            html.append("<tr style='background-color: #0f172a; color: white;'><th>Line</th>")
            for h in hari_tabel: html.append(f"<th colspan='4'>{h}</th><th style='width: 15px;'></th>") 
            html.append("</tr>")
            
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
                        elif is_pred: bg = "#fee2e2" 
                        border_style = "2px solid #dc2626" if is_pred else "1px solid #ccc"
                        text_color = "#dc2626" if is_pred else "inherit"
                        html.append(f"<td style='border: {border_style}; background-color: {bg}; color: {text_color}; font-weight: bold; width: 25px; height: 25px;'>{display_val}</td>")
                    html.append("<td style='width: 15px;'></td>")
                html.append("</tr>")
            html.append("</table></div>")
            st.markdown("".join(html), unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")
