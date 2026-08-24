import streamlit as st
import openpyxl
import io
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from collections import Counter
from PIL import Image, ImageDraw, ImageFont

# Setup halaman
st.set_page_config(page_title="JEPE AI Pro", layout="wide")
st.title("JEPE AI")

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

# Fungsi generate_pdf (Visual Ekspor dengan Garis Alur)
def generate_pdf(ws, highlighted_data, prediction_cells, all_path_lines, start_cols, hari_tabel):
    SCALE = 2 
    CELL_W = 28 * SCALE
    ROW_H = 30 * SCALE
    GAP_W = 12 * SCALE
    COL_LINE_W = 40 * SCALE
    HDR_H = 32 * SCALE
    
    r_min = max(1, ws.max_row - 40)
    r_max = ws.max_row
    num_rows = r_max - r_min + 1
    
    width = COL_LINE_W + 7 * (4 * CELL_W + GAP_W)
    height = HDR_H + num_rows * ROW_H
    
    img = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(img)
    
    font = ImageFont.load_default()
    font_bold = font
    try:
        font = ImageFont.truetype("arial.ttf", 14 * SCALE)
        font_bold = ImageFont.truetype("arialbd.ttf", 14 * SCALE)
    except: pass 
        
    def get_center(r, c):
        row_idx = r - r_min
        cy = HDR_H + row_idx * ROW_H + (ROW_H / 2)
        cx = COL_LINE_W
        for day_i, sc in enumerate(start_cols):
            if sc <= c <= sc + 3:
                offset = c - sc
                cx += day_i * (4 * CELL_W + GAP_W) + offset * CELL_W + (CELL_W / 2)
                return cx, cy
        return None, None

    draw.rectangle([0, 0, width, HDR_H], fill="#0f172a")
    draw.text((10*SCALE, 8*SCALE), "Line", fill="white", font=font_bold)
    
    for i, h in enumerate(hari_tabel):
        x = COL_LINE_W + i * (4 * CELL_W + GAP_W)
        draw.text((x + CELL_W, 8*SCALE), h, fill="white", font=font_bold)
        
    bg_colors = {0: "#3399FF", 1: "#D2B48C", 2: "#22C55E", 3: "#FFD700"}
    
    for r in range(r_min, r_max + 1):
        row_idx = r - r_min
        y = HDR_H + row_idx * ROW_H
        
        draw.rectangle([0, y, COL_LINE_W, y + ROW_H], fill="#f0f0f0", outline="#cccccc")
        draw.text((10*SCALE, y + 6*SCALE), str(r), fill="black", font=font_bold)
        
        for i, start_col in enumerate(start_cols):
            for offset in range(4):
                c_idx = start_col + offset
                x = COL_LINE_W + i * (4 * CELL_W + GAP_W) + offset * CELL_W
                
                val = ws.cell(row=r, column=c_idx).value
                display_val = str(val) if val is not None else "-"
                
                bg = "white"
                is_pred = (r, c_idx) in prediction_cells
                if (r, c_idx) in highlighted_data:
                    p = highlighted_data[(r, c_idx)]["pos"]
                    bg = bg_colors.get(p, "white")
                elif is_pred:
                    bg = "#fee2e2"
                    
                outline = "#dc2626" if is_pred else "#cccccc"
                text_color = "#dc2626" if is_pred else "black"
                
                draw.rectangle([x, y, x + CELL_W, y + ROW_H], fill=bg, outline=outline)
                draw.text((x + 8*SCALE, y + 6*SCALE), display_val, fill=text_color, font=font)
                
    stroke_colors = {0: (51, 153, 255), 1: (210, 180, 140), 2: (34, 197, 94), 3: (255, 215, 0)}
    for path in all_path_lines:
        pos = path["pos"]
        nodes = path["nodes"]
        color = stroke_colors.get(pos, (255, 0, 0))
        
        pts = []
        for r, c in nodes:
            if r_min <= r <= r_max:
                cx, cy = get_center(r, c)
                if cx is not None:
                    pts.append((cx, cy))
        
        if len(pts) > 1:
            draw.line(pts, fill=color, width=2*SCALE)
            
    buf = io.BytesIO()
    img.save(buf, format="PDF", resolution=100.0)
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
        
        c_terdekat = st.checkbox("Toleransi Angka Terdekat (+/- 1 dari Angka & Indeks)", value=True)
        
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
            all_path_lines = []
            
            # --- PENAMBAHAN VARIABEL UNTUK POLA TERPANJANG ---
            longest_pattern_per_pos = {0: 0, 1: 0, 2: 0, 3: 0}

            batas_bawah = ws.max_row 

            for pos_offset in range(4):
                current_allowed = []
                for k in range(6):
                    val_str = inputs[k]
                    digit = int(val_str[ref_pos_offset if use_single_ref else pos_offset])
                    indek = (digit + 5) % 10
                    
                    if c_terdekat:
                        allowed = {
                            digit, indek,
                            (digit - 1) % 10, (digit + 1) % 10,
                            (indek - 1) % 10, (indek + 1) % 10
                        }
                        current_allowed.append(list(allowed))
                    else:
                        current_allowed.append([digit, indek])

                for r_start in range(1, batas_bawah):
                    for mode in ["Lurus", "Naik", "Turun"]:
                        if (mode == "Lurus" and not c_lurus) or (mode == "Naik" and not c_naik) or (mode == "Turun" and not c_turun): continue
                        
                        for length in [6, 5, 4, 3]:
                            path, valid = [], True
                            for k in range(length):
                                r_target = r_start if mode == "Lurus" else (r_start - k if mode == "Naik" else r_start + k)
                                
                                if r_target < 1 or r_target >= batas_bawah: 
                                    valid = False; break
                                
                                cell_val = ws.cell(row=r_target, column=start_cols[days_indices[k]] + pos_offset).value
                                val = clean_int(cell_val)
                                if val not in current_allowed[k]: valid = False; break
                                path.append((r_target, start_cols[days_indices[k]] + pos_offset))
                            
                            if valid:
                                total_stats[length] += 1
                                for r_c, c_c in path: cell_patterns[(r_c, c_c)] = {"length": length, "pos": pos_offset}
                                
                                # --- MENYIMPAN DATA POLA TERPANJANG ---
                                longest_pattern_per_pos[pos_offset] = max(longest_pattern_per_pos[pos_offset], length)
                                
                                r_next = r_start if mode == "Lurus" else (r_start + 1 if mode == "Naik" else r_start - 1)
                                
                                if 1 <= r_next < batas_bawah:
                                    c_next_day_idx = (idx_day0 + 1) % 7
                                    c_next = start_cols[c_next_day_idx] + pos_offset
                                    
                                    pred_val = clean_int(ws.cell(row=r_next, column=c_next).value)
                                    if pred_val is not None:
                                        predictions_raw[pos_offset].append(pred_val)
                                        prediction_cells.add((r_next, c_next))
                                        
                                        full_path = list(reversed(path)) + [(r_next, c_next)]
                                        all_path_lines.append({"pos": pos_offset, "nodes": full_path})
                                        
                                break 

            prediction_results = {}
            for p in range(4):
                prediction_results[p] = Counter(predictions_raw[p])

            st.session_state.highlighted = cell_patterns
            st.session_state.stats = total_stats
            st.session_state.prediction_results = prediction_results
            st.session_state.prediction_cells = prediction_cells
            st.session_state.all_path_lines = all_path_lines
            st.session_state.longest_pattern_per_pos = longest_pattern_per_pos # Menyimpan ke memori
            st.session_state.scanned = True
            st.rerun()

        # 5. OUTPUT
        if st.session_state.get("scanned"):
            st.divider()
            
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                excel_buffer = generate_excel(ws, st.session_state.get("highlighted", {}))
                st.download_button(
                    label="📥 Download Excel (.xlsx)",
                    data=excel_buffer,
                    file_name="hasil_scan_paito.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with dl_col2:
                pdf_buffer = generate_pdf(
                    ws, 
                    st.session_state.get("highlighted", {}),
                    st.session_state.get("prediction_cells", set()),
                    st.session_state.get("all_path_lines", []),
                    start_cols,
                    hari_tabel
                )
                st.download_button(
                    label="🖨️ Download Laporan PDF (Visual Garis Alur)",
                    data=pdf_buffer,
                    file_name="hasil_visual_paito.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            
            st.subheader("🎯 Ringkasan Jalur Prediksi (Gabungan Angka & Indeks)")
            pos_names = ["As", "Kop", "Kepala", "Ekor"]
            
            # --- MENGAMBIL DATA POLA TERPANJANG ---
            longest_patterns = st.session_state.get("longest_pattern_per_pos", {0: 0, 1: 0, 2: 0, 3: 0})
            
            pred_cols = st.columns(4)
            for p in range(4):
                with pred_cols[p]:
                    st.markdown(f"### **Posisi {pos_names[p]}**")
                    
                    # --- MENAMPILKAN INDIKATOR POLA TERPANJANG ---
                    max_len = longest_patterns.get(p, 0)
                    if max_len > 0:
                        st.markdown(f"**🔥 Pola Terpanjang: {max_len} Hari**")
                    else:
                        st.markdown("**🔥 Pola Terpanjang: -**")
                        
                    counts = st.session_state.get("prediction_results", {}).get(p, Counter())
                    
                    if counts:
                        combined_list = []
                        for base in range(5):
                            idx_partner = base + 5
                            c_base = counts.get(base, 0)
                            c_idx = counts.get(idx_partner, 0)
                            total_combined = c_base + c_idx
                            if total_combined > 0:
                                combined_list.append((base, idx_partner, total_combined, c_base, c_idx))
                        
                        combined_list.sort(key=lambda x: x[2], reverse=True)
                        
                        for base, idx_partner, total_combined, c_base, c_idx in combined_list:
                            st.write(f"🔹 **Angka {base} / {idx_partner}:** {total_combined} Jalur *(Detail: {base}={c_base}, {idx_partner}={c_idx})*")
                    else:
                        st.info("Tidak ada jalur pola ditemukan")
            
            st.divider()
            
            st.subheader("Statistik Jalur Keseluruhan")
            stats = st.session_state.stats
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Pola 6 Hari", f"{stats[6]} Jalur")
            c2.metric("Pola 5 Hari", f"{stats[5]} Jalur")
            c3.metric("Pola 4 Hari", f"{stats[4]} Jalur")
            c4.metric("Pola 3 Hari", f"{stats[3]} Jalur")

            st.subheader("Live Preview Grid (dengan Garis Alur)")
            highlighted = st.session_state.get("highlighted", {})
            prediction_cells = st.session_state.get("prediction_cells", set())
            all_path_lines = st.session_state.get("all_path_lines", [])
            
            ROW_H = 30
            HDR_H = 32
            COL_LINE_W = 40
            CELL_W = 28
            GAP_W = 12
            
            r_min = max(1, ws.max_row - 30)
            r_max = ws.max_row
            num_displayed_rows = r_max - r_min + 1
            
            total_width = COL_LINE_W + 7 * (4 * CELL_W + GAP_W)
            total_height = HDR_H + num_displayed_rows * ROW_H
            
            def get_cell_center_html(r, c):
                if r < r_min or r > r_max: return None, None
                row_idx = r - r_min
                cy = HDR_H + row_idx * ROW_H + (ROW_H / 2)
                
                cx = None
                for day_i, sc in enumerate(start_cols):
                    if sc <= c <= sc + 3:
                        offset = c - sc
                        cx = COL_LINE_W + day_i * (4 * CELL_W + GAP_W) + offset * CELL_W + (CELL_W / 2)
                        break
                return cx, cy

            stroke_colors = {
                0: "rgba(51, 153, 255, 0.7)",  
                1: "rgba(210, 180, 140, 0.8)", 
                2: "rgba(34, 197, 94, 0.7)",   
                3: "rgba(255, 215, 0, 0.8)"    
            }
            
            svg_elements = []
            for path in all_path_lines:
                pos = path["pos"]
                nodes = path["nodes"]
                color = stroke_colors.get(pos, "rgba(255, 0, 0, 0.6)")
                
                points = []
                for r, c in nodes:
                    cx, cy = get_cell_center_html(r, c)
                    if cx is not None and cy is not None:
                        points.append(f"{cx},{cy}")
                
                if len(points) > 1:
                    pts_str = " ".join(points)
                    svg_elements.append(f'<polyline points="{pts_str}" fill="none" stroke="{color}" stroke-width="2" stroke-dasharray="3,3" />')

            svg_html = f'<svg style="position: absolute; top: 0; left: 0; width: {total_width}px; height: {total_height}px; pointer-events: none; z-index: 2;">{"".join(svg_elements)}</svg>'

            html = [f"<div style='position: relative; overflow-x: auto; width: 100%;'>{svg_html}<table style='table-layout: fixed; border-collapse: collapse; width: {total_width}px; text-align: center; font-family: monospace; font-size: 12px; z-index: 1;'>"]
            
            html.append(f"<tr style='background-color: #0f172a; color: white; height: {HDR_H}px;'><th style='width: {COL_LINE_W}px;'>Line</th>")
            for h in hari_tabel:
                html.append(f"<th colspan='4' style='width: {4*CELL_W}px;'>{h}</th><th style='width: {GAP_W}px;'></th>") 
            html.append("</tr>")
            
            for r in range(r_min, r_max + 1):
                html.append(f"<tr style='height: {ROW_H}px;'><td style='border: 1px solid #ccc; background-color: #f0f0f0; width: {COL_LINE_W}px; font-weight: bold;'>{r}</td>")
                
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
                        
                        html.append(f"<td style='border: {border_style}; background-color: {bg}; color: {text_color}; font-weight: bold; width: {CELL_W}px;'>{display_val}</td>")
                    
                    html.append(f"<td style='width: {GAP_W}px;'></td>")
                    
                html.append("</tr>")
            html.append("</table></div>")
            
            st.markdown("".join(html), unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")
