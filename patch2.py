import re

with open('generador_sifarmu_actualizado.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Update HOJAS columns for 02_MEDICAMENTOS
old_med_cols = """[
        "ID", "Código", "Código LINAME", "Código ATC", "Categoría", "Subcategoría",
        "Principio Activo", "Nombre Comercial", "Concentración", "Forma Farmacéutica",
        "Presentación", "Unidad", "Vía", "Requiere Receta", "Cadena de Frío", "Stock Actual", "Estado"
    ]"""
new_med_cols = """[
        "ID", "Código", "Código LINAME", "Categoría",
        "Principio Activo", "Nombre Comercial", "Concentración", "Forma Farmacéutica",
        "Vía", "Stock Actual", "Estado"
    ]"""
code = code.replace(old_med_cols, new_med_cols)

# 2. Update the writing loop for 02_MEDICAMENTOS
old_loop = """
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
new_loop = """
    ws_med = sheets["02_MEDICAMENTOS"]
    for row_num, m_data in enumerate(MEDICAMENTOS, start=1):
        for col_idx in range(11):
            ws_med.write(row_num, col_idx, "", cell_format)
        ws_med.write(row_num, 0, f"ID-{row_num:04d}", cell_format)
        ws_med.write(row_num, 1, m_data[0], cell_format) # Código (B)
        ws_med.write(row_num, 3, m_data[1], cell_format) # Categoría (D)
        ws_med.write(row_num, 4, m_data[2], cell_format) # Princ Activo (E)
        ws_med.write(row_num, 6, m_data[3], cell_format) # Concentración (G)
        ws_med.write(row_num, 7, m_data[4], cell_format) # Forma Farm (H)
        
        # Fórmula para el Stock Actual
        formula = f'=SUMIF(\\'03_LOTES\\'!B:B, B{row_num+1}, \\'03_LOTES\\'!E:E)'
        ws_med.write(row_num, 9, formula, cell_format) # Stock Actual (J)
        
        ws_med.write(row_num, 10, "ACTIVO", cell_format) # Estado (K)
"""
code = code.replace(old_loop.strip(), new_loop.strip())

# 3. Update data validations for 02_MEDICAMENTOS
code = code.replace("ws_med.data_validation('E2:E1000',", "ws_med.data_validation('D2:D1000',")
code = code.replace("ws_med.data_validation('Q2:Q1000',", "ws_med.data_validation('K2:K1000',")

# 4. Add Date Validations to 03_LOTES for User-Friendly Date Picker
import_dt = "import xlsxwriter\\nfrom datetime import datetime\\n"
code = code.replace("import xlsxwriter\\n", import_dt)

date_val = """
    ws_lotes.data_validation('G2:H1000', {
        'validate': 'date',
        'criteria': 'between',
        'minimum': datetime(2000, 1, 1),
        'maximum': datetime(2100, 12, 31)
    })
    
    ws_dash = sheets["13_DASHBOARD"]
"""
code = code.replace('    ws_dash = sheets["13_DASHBOARD"]', date_val.lstrip('\n'))

with open('generador_sifarmu_actualizado.py', 'w', encoding='utf-8') as f:
    f.write(code)

