import streamlit as st
import openpyxl
import io
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from collections import Counter

# Setup halaman
st.set_page_config(page_title="JEPE AI Pro - Cyberpunk Edition", layout="wide")

# ==========================================
# INJEKSI CSS TEMA CYBERPUNK
# ==========================================
cyberpunk_css = """
<style>
    /* Latar Belakang & Teks Global */
    .stApp {
        background-color: #0b0c10;
        color: #66fcf1;
        font-family: 'Courier New', Courier, monospace;
    }
    
    /* Header & Judul */
    h1, h2, h3 {
        color: #ff007f !important;
        text-shadow: 0 0 5px #ff007f, 0 0 10px #ff007f;
        text-transform: uppercase;
    }
    
    /* Tombol (Glow Effect) */
    .stButton>button {
        background-color: transparent;
        color: #00ffcc;
        border: 2px solid #00ffcc;
        box-shadow: 0 0 8px #00ffcc;
        border-radius: 0px;
        text-transform: uppercase;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #00ffcc;
        color: #0b0c10;
        box-shadow: 0 0 15px #00ffcc, 0 0 30px #00ffcc;
        border-color: #00ffcc;
    }

    /* Input Fields (Dark dengan border Neon) */
    .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div>select {
        background-color: #1f2833 !important;
        color: #fcee0a !important;
        border: 1px solid #ff007f !important;
        border-radius: 0px;
    }
    .stTextInput>div>div>input:focus {
        box-shadow: 0 0 10px #ff007f;
    }
    
    /* Alert / Notifikasi Box */
    [data-testid="stAlert"] {
        background-color: rgba(255, 0, 127, 0.1);
        border-left: 4px solid #ff007f;
        color: #ffffff;
        box-shadow: 0 0 10px rgba(255, 0, 127, 0.3);
    }
    
    /* Metrik Statistik */
    [data-testid="stMetricValue"] {
        color: #fcee0a;
        text-shadow: 0 0 8px #fcee0a;
    }
    [data-testid="stMetricLabel"] {
        color: #00ffcc;
    }
    
    /* Garis Pembatas (Divider) */
    hr {
        border-color: #ff007f;
        box-shadow: 0 0 5px #ff007f;
    }
    
    /* Kustomisasi Expander */
    .streamlit-expanderHeader {
        color: #66fcf1 !important;
        border-bottom: 1px solid #00ffcc;
    }
</style>
"""
st.markdown(cyberpunk_css, unsafe_allow_html=True)
# ==========================================

st.title("JEPE AI - CYBERPUNK SCANNER")

# Fungsi pembersih data
def clean_int(v):
    try: return int(float(str(v).strip()))
    except (ValueError, TypeError): return None

# Fungsi generate_excel dengan warna neon cyberpunk
def generate_excel(original_ws, highlighted_data):
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    
    for col_num in range(1, original_ws.max_column + 1):
        col_letter = get_column_letter(col_num)
        new_ws.column_dimensions[col_letter].width = 3
    
    # Palet Excel Cyberpunk
    colors = {0: "00FFFF", 1: "FF007F", 2: "39FF14", 3: "FCEE0A"}
    
    for r in range(1, original_ws.max_row + 1):
        for c in range(1, original_ws.max_column + 1):
            cell_val = original_ws.cell(row=r, column=c).value
            new_ws.cell(row=r, column=c).value = cell_val
            
            if (r, c) in highlighted_data:
                pos = highlighted_data[(r, c)]["pos"]
                hex_color = colors.get(pos, "454545")
                new_ws.cell(row=r, column=c).fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    
    buf = io.BytesIO()
    new_wb.save(buf)
    buf.seek(0)
    return buf

# 1. UPLOAD FILE
uploaded_file = st.file_uploader("UNGGAH DATABASE PAITO (.XLSX):", type=["xlsx"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        
        hari_tabel = ["Sabtu", "Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat"]
        start_cols = [1, 6, 11, 16, 21, 26, 31]

        st.header("2. INPUT REFERENSI (AUTO-FETCH & LIVE UPDATE)")
        
        c_opt1, c_opt2 = st.columns(2)
        with c_opt1:
            hari_terpilih = st.selectbox("HARI UTAMA (HARI INI):", hari_tabel, index=4)
        with c_opt2:
            target_row_utama = st.number_input("BARIS TARGET (DEFAULT: TERAKHIR)", min_value=1, value=ws.max_row)

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
                user_val = st.text_input(f"{hari_tabel[d_idx].upper()} (-{i}):", value=auto_val, max_chars=4)
                inputs.append(user_val)
                update_targets.append((current_r, c_start))

        for i, user_val in enumerate(inputs):
            r_target, c_start = update_targets[i]
            user_val = user_val.ljust(4, '0')[:4] 
            
            for offset in range(4):
                try:
                    ws.cell(row=r_target, column=c_start + offset).value = int(user_val[offset])
                except ValueError:
                    pass

        st.divider()
        c_lurus = st.checkbox("GARIS LURUS", value=True)
        c_naik = st.checkbox("DIAGONAL NAIK", value=True)
        c_turun = st.checkbox("DIAGONAL TURUN", value=True)
        
        use_single_ref = st.checkbox("MODE ACUAN POSISI TUNGGAL", value=False)
        ref_pos_name = st.selectbox("POSISI ACUAN:", ["As", "Kop", "Kepala", "Ekor"], index=0, disabled=not use_single_ref)
        ref_pos_offset = ["As", "Kop", "Kepala", "Ekor"].index(ref_pos_name)

        # 4. LOGIKA SCANNING
        if st.button("EXECUTE ANALYSIS // JALANKAN"):
            cell_patterns = {}
            total_stats = {6: 0, 5: 0, 4: 0, 3: 0}
            days_indices = [(idx_day0 - k) % 7 for k in range(6)]
            
            predictions_raw = {0: [], 1: [], 2: [], 3: []}
            prediction_cells = set()

            # [REVISI] Dinamisasi Batas Bawah Scan
            # Cari tahu baris berapa saja yang dipakai sebagai input referensi
            ref_rows = [r for r, c in update_targets]
            batas_bawah_scan = min(ref_rows) # Batas ini secara ketat mencegah referensi men-scan dirinya sendiri

            for pos_offset in range(4):
                current_allowed = []
                for k in range(6):
                    val_str = inputs[k]
                    digit = int(val_str[ref_pos_offset if use_single_ref else pos_offset])
                    current_allowed.append([digit, (digit + 5) % 10])

                # Scan hanya boleh mencari histori sampai sebelum baris referensi (batas_bawah_scan)
                for r_start in range(1, batas_bawah_scan):
                    for mode in ["Lurus", "Naik", "Turun"]:
                        if (mode == "Lurus" and not c_lurus) or (mode == "Naik" and not c_naik) or (mode == "Turun" and not c_turun): continue
                        
                        for length in [6, 5, 4, 3]:
                            path, valid = [], True
                            for k in range(length):
                                r_target = r_start if mode == "Lurus" else (r_start - k if mode == "Naik" else r_start + k)
                                
                                # Batalkan lintasan pola jika menyentuh baris yang dijadikan referensi
                                if r_target < 1 or r_target >= batas_bawah_scan: 
                                    valid = False; break
                                
                                cell_val = ws.cell(row=r_target, column=start_cols[days_indices[k]] + pos_offset).value
                                val = clean_int(cell_val)
                                if val not in current_allowed[k]: valid = False; break
                                path.append((r_target, start_cols[days_indices[k]] + pos_offset))
                            
                            if valid:
                                total_stats[length] += 1
                                for r_c, c_c in path: cell_patterns[(r_c, c_c)] = {"length": length, "pos": pos_offset}
                                
                                r_next = r_start if mode == "Lurus" else (r_start + 1 if mode == "Naik" else r_start - 1)
                                
                                # Proyeksi (hari prediksi) tetap diperbolehkan meskipun jatuh di area baris referensi
                                if 1 <= r_next <= ws.max_row:
                                    c_next_day_idx = (idx_day0 + 1) % 7
                                    c_next = start_cols[c_next_day_idx] + pos_offset
                                    
                                    pred_val = clean_int(ws.cell(row=r_next, column=c_next).value)
                                    if pred_val is not None:
                                        predictions_raw[pos_offset].append({"val": pred_val, "length": length})
                                        prediction_cells.add((r_next, c_next))
                                        
                                break 

            prediction_results = {}
            for p in range(4):
                preds = predictions_raw[p]
                if not preds: continue
                
                angka_stats = {}
                for x in preds:
                    v = x['val']
                    l = x['length']
                    
                    if v not in angka_stats:
                        angka_stats[v] = {'long_count': 0, 'short_count': 0, 'total': 0, 'max_len': 0}
                    angka_stats[v]['total'] += 1
                    
                    if l >= 4: angka_stats[v]['long_count'] += 1
                    else: angka_stats[v]['short_count'] += 1
                        
                    if l > angka_stats[v]['max_len']: angka_stats[v]['max_len'] = l

                kuat_candidates = []
                cadangan_candidates = []
                all_vals = [x['val'] for x in preds]
                
                for v, stats in angka_stats.items():
                    if stats['long_count'] > 0 or stats['short_count'] > 1:
                        kuat_candidates.append(v)
                    else:
                        cadangan_candidates.append(v)
                
                kuat_candidates.sort(key=lambda x: (angka_stats[x]['total'], angka_stats[x]['max_len']), reverse=True)
                cadangan_candidates.sort(key=lambda x: (angka_stats[x]['total'], angka_stats[x]['max_len']), reverse=True)
                
                angka_kuat = [kuat_candidates[0]] if kuat_candidates else []
                angka_cadangan = []
                
                if len(kuat_candidates) > 1: angka_cadangan = [kuat_candidates[1]]
                elif cadangan_candidates: angka_cadangan = [cadangan_candidates[0]]
                    
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
                label="[⬇️] DOWNLOAD REPORT (.XLSX)",
                data=excel_buffer,
                file_name="hasil_scan_cyberpunk.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.subheader("TARGET LOCK // PREDIKSI HARI BERIKUTNYA")
            pos_names = ["As", "Kop", "Kepala", "Ekor"]
            
            pred_cols = st.columns(4)
            for p in range(4):
                with pred_cols[p]:
                    st.markdown(f"**[{pos_names[p].upper()}]**")
                    if p in st.session_state.get("prediction_results", {}):
                        res = st.session_state.prediction_results[p]
                        
                        kuat_str = str(res['kuat'][0]) if res['kuat'] else "-"
                        cadangan_str = str(res['cadangan'][0]) if res['cadangan'] else "-"
                        
                        pola_info = f"(Pola {res['max_len']} Baris)" if res['kuat'] else "(No Data)"
                        st.success(f"🔥 **KUAT:** {kuat_str}\n\n{pola_info}")
                        st.info(f"🛡️ **CADANGAN:** {cadangan_str}")
                        
                        total_kemunculan = sum(res['all_counts'].values())
                        
                        with st.expander("DATA FREKUENSI"):
                            for val, count in res['all_counts'].most_common():
                                persentase = (count / total_kemunculan) * 100 if total_kemunculan > 0 else 0
                                status_label = "🔥 [BET]" if persentase >= 50 else "⏳ [TUNGGU]"
                                status_tipe = " (Kuat)" if val in res['kuat'] else (" (Cadangan)" if res['cadangan'] and val == res['cadangan'][0] else "")
                                
                                st.write(f"V:{val} | {count}x ({persentase:.1f}%) {status_label}{status_tipe}")
                    else:
                        st.write("NO SIGNAL")
            
            st.divider()
            
            results = st.session_state.get("prediction_results", {})
            if all(p in results and results[p]['kuat'] for p in range(4)):
                kuat_4d = [results[p]['kuat'][0] for p in range(4)]
                index_map = {0:5, 1:6, 2:7, 3:8, 4:9, 5:0, 6:1, 7:2, 8:3, 9:4}
                index_4d = [index_map[d] for d in kuat_4d]
                found_match = False
                
                for r in range(1, ws.max_row): 
                    for c_start in start_cols:
                        try:
                            row_val = [clean_int(ws.cell(row=r, column=c_start+offset).value) for offset in range(4)]
                            if None in row_val: continue
                            match_asli = (row_val[1:] == kuat_4d[1:])
                            match_idx = (row_val[1:] == index_4d[1:])
                            
                            if match_asli or match_idx:
                                found_match = True
                                break
                        except Exception:
                            pass
                    if found_match: break
                
                if found_match:
                    st.success(f"⚠️ **JACKPOT TERDETEKSI (BET MANTAP!)** ⚠️\n\nFormasi 4D Kuat: **{''.join(map(str, kuat_4d))}** (Index: **{''.join(map(str, index_4d))}**) telah divalidasi dengan sejarah keluaran historis (Min 3D Matches)!")

            st.divider()
            
            st.subheader("SYSTEM STATS // POLA JALUR")
            stats = st.session_state.stats
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("6 HARI", f"{stats[6]} PATHS")
            c2.metric("5 HARI", f"{stats[5]} PATHS")
            c3.metric("4 HARI", f"{stats[4]} PATHS")
            c4.metric("3 HARI", f"{stats[3]} PATHS")

            st.subheader("NEON GRID // LIVE PREVIEW")
            highlighted = st.session_state.get("highlighted", {})
            prediction_cells = st.session_state.get("prediction_cells", set())
            
            html = ["<div style='overflow-x: auto; box-shadow: 0 0 10px #00ffcc; padding: 10px; background-color: #0b0c10;'><table style='border-collapse: collapse; width: 100%; text-align: center; font-family: Courier New, monospace; font-size: 13px;'>"]
            
            # Header Cyberpunk
            html.append("<tr style='background-color: #1f2833; color: #ff007f; text-shadow: 0 0 3px #ff007f;'><th>LINE</th>")
            for h in hari_tabel:
                html.append(f"<th colspan='4' style='border: 1px solid #00ffcc;'>{h.upper()}</th><th style='width: 15px;'></th>") 
            html.append("</tr>")
            
            # Data Rows Cyberpunk
            for r in range(max(1, ws.max_row - 30), ws.max_row + 1):
                html.append(f"<tr><td style='border: 1px solid #333; background-color: #1a1a1a; color: #00ffcc; width: 25px; height: 25px; text-align: center;'>{r}</td>")
                
                for i, start_col in enumerate(start_cols):
                    for offset in range(4):
                        c_idx = start_col + offset
                        val = ws.cell(row=r, column=c_idx).value
                        display_val = str(val) if val is not None else "-"
                        
                        is_pred = (r, c_idx) in prediction_cells
                        bg = "#0b0c10"
                        text_color = "#66fcf1"
                        border_style = "1px solid #333"
                        shadow = ""
                        
                        if (r, c_idx) in highlighted:
                            p = highlighted[(r, c_idx)]["pos"]
                            # Warna Neon Cyberpunk
                            colors = {0: "#00ffff", 1: "#ff007f", 2: "#39ff14", 3: "#fcee0a"}
                            bg = colors.get(p, "#0b0c10")
                            text_color = "#000000"
                            border_style = f"1px solid {bg}"
                            shadow = f"box-shadow: 0 0 8px {bg};"
                        elif is_pred:
                            bg = "#2a0000"
                            text_color = "#ff0000"
                            border_style = "2px solid #ff0000"
                            shadow = "box-shadow: 0 0 10px #ff0000;"
                        
                        html.append(f"<td style='border: {border_style}; background-color: {bg}; color: {text_color}; {shadow} font-weight: bold; width: 25px; height: 25px;'>{display_val}</td>")
                    
                    html.append("<td style='width: 15px;'></td>")
                    
                html.append("</tr>")
            html.append("</table></div>")
            st.markdown("".join(html), unsafe_allow_html=True)

    except Exception as e:
        st.error(f"SYSTEM ERROR: {e}")
