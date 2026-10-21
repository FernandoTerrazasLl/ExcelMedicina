import re

with open('generador_sifarmu_actualizado.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Fix 1: Visual Inconsistency (Borders)
old_med_loop = """
    ws_med = sheets["02_MEDICAMENTOS"]
    for row_num, m_data in enumerate(MEDICAMENTOS, start=1):
        ws_med.write(row_num, 0, f"ID-{row_num:04d}", cell_format)
        ws_med.write(row_num, 1, m_data[0], cell_format)
        ws_med.write(row_num, 4, m_data[1], cell_format)
        ws_med.write(row_num, 6, m_data[2], cell_format)
        ws_med.write(row_num, 8, m_data[3], cell_format)
        ws_med.write(row_num, 9, m_data[4], cell_format)
        ws_med.write(row_num, 15, "ACTIVO", cell_format)
"""
new_med_loop = """
    ws_med = sheets["02_MEDICAMENTOS"]
    for row_num, m_data in enumerate(MEDICAMENTOS, start=1):
        for col_idx in range(16):
            ws_med.write(row_num, col_idx, "", cell_format)
        ws_med.write(row_num, 0, f"ID-{row_num:04d}", cell_format)
        ws_med.write(row_num, 1, m_data[0], cell_format)
        ws_med.write(row_num, 4, m_data[1], cell_format)
        ws_med.write(row_num, 6, m_data[2], cell_format)
        ws_med.write(row_num, 8, m_data[3], cell_format)
        ws_med.write(row_num, 9, m_data[4], cell_format)
        ws_med.write(row_num, 15, "ACTIVO", cell_format)
"""
code = code.replace(old_med_loop.strip(), new_med_loop.strip())

# Fix 2: Dashboard COUNTIF -> COUNTIFS
old_countif = """('Lotes Vencidos', "=COUNTIF('03_LOTES'!H:H, \\"<\\"&TODAY())")"""
new_countif = """('Lotes Vencidos', "=COUNTIFS('03_LOTES'!H:H, \\"<\\"&TODAY(), '03_LOTES'!H:H, \\"<>\\")")"""
code = code.replace(old_countif, new_countif)

with open('generador_sifarmu_actualizado.py', 'w', encoding='utf-8') as f:
    f.write(code)
