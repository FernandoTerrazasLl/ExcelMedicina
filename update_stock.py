with open('generador_sifarmu_actualizado.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Update the header in HOJAS
old_header = '"Requiere Receta", "Cadena de Frío", "Estado"'
new_header = '"Requiere Receta", "Cadena de Frío", "Stock Actual", "Estado"'
code = code.replace(old_header, new_header)

# Update the writing loop in 02_MEDICAMENTOS
old_loop = """
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
new_loop = """
    ws_med = sheets["02_MEDICAMENTOS"]
    for row_num, m_data in enumerate(MEDICAMENTOS, start=1):
        for col_idx in range(17):
            ws_med.write(row_num, col_idx, "", cell_format)
        ws_med.write(row_num, 0, f"ID-{row_num:04d}", cell_format)
        ws_med.write(row_num, 1, m_data[0], cell_format)
        ws_med.write(row_num, 4, m_data[1], cell_format)
        ws_med.write(row_num, 6, m_data[2], cell_format)
        ws_med.write(row_num, 8, m_data[3], cell_format)
        ws_med.write(row_num, 9, m_data[4], cell_format)
        
        # Fórmula para el Stock Actual (Suma de los disponibles en la hoja 03_LOTES)
        # La columna B de esta hoja tiene el Código, así que usamos B2, B3, etc.
        formula = f'=SUMIF(\\'03_LOTES\\'!B:B, B{row_num+1}, \\'03_LOTES\\'!E:E)'
        ws_med.write(row_num, 15, formula, cell_format)
        
        ws_med.write(row_num, 16, "ACTIVO", cell_format)
"""

code = code.replace(old_loop.strip(), new_loop.strip())

# Also update the data validation for 'Estado' from P2:P1000 to Q2:Q1000 because it shifted
code = code.replace("ws_med.data_validation('P2:P1000'", "ws_med.data_validation('Q2:Q1000'")

with open('generador_sifarmu_actualizado.py', 'w', encoding='utf-8') as f:
    f.write(code)

