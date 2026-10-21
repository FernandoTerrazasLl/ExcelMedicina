import re

with open('generador_sifarmu_actualizado.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Replace the MEDICAMENTOS writing loop
old_med_loop = """    ws_med = sheets["02_MEDICAMENTOS"]
    for row_num, m_data in enumerate(MEDICAMENTOS, start=1):
        for col_idx in range(9):
            ws_med.write(row_num, col_idx, "", cell_format)
        # 0: Princ Activo, 1: Conc, 2: Forma, 3: Cat, 4: Stock, 5: Nombre Comp, 6: Cod, 7: LINAME, 8: Estado
        ws_med.write(row_num, 0, m_data[2], cell_format) # Princ Activo
        ws_med.write(row_num, 1, m_data[3], cell_format) # Concentracion
        ws_med.write(row_num, 2, m_data[4], cell_format) # Forma Farm
        ws_med.write(row_num, 3, m_data[1], cell_format) # Categoria
        
        formula_stock = f'=SUMIF(\\'03_LOTES\\'!B:B, F{row_num+1}, \\'03_LOTES\\'!E:E)'
        ws_med.write(row_num, 4, formula_stock, cell_format) # Stock
        
        formula_nombre = f'=A{row_num+1} & " | " & B{row_num+1} & " | " & C{row_num+1}'
        ws_med.write(row_num, 5, formula_nombre, cell_format) # Nombre Completo
        
        ws_med.write(row_num, 6, m_data[0], cell_format) # Codigo
        ws_med.write(row_num, 8, "ACTIVO", cell_format)  # Estado"""

new_med_loop = """    ws_med = sheets["02_MEDICAMENTOS"]
    
    # Pre-format columns to ensure borders exist even if cells are populated by ARRAYFORMULA
    ws_med.set_column('A:I', 20, cell_format)
    
    for row_num, m_data in enumerate(MEDICAMENTOS, start=1):
        # We don't write anything to col 4 (Stock) and 5 (Nombre Completo) so ARRAYFORMULA can expand freely
        ws_med.write(row_num, 0, m_data[2], cell_format) # Princ Activo
        ws_med.write(row_num, 1, m_data[3], cell_format) # Concentracion
        ws_med.write(row_num, 2, m_data[4], cell_format) # Forma Farm
        ws_med.write(row_num, 3, m_data[1], cell_format) # Categoria
        ws_med.write(row_num, 6, m_data[0], cell_format) # Codigo
        ws_med.write(row_num, 8, "ACTIVO", cell_format)  # Estado
        
    # Inyectar las ARRAYFORMULAS Maestras en la fila 2 (row_num = 1 en xlsxwriter)
    # Importante: Como Google Sheets asume la sintaxis en inglés para fórmulas internas, usamos IF y SUMIF
    array_formula_stock = '=ARRAYFORMULA(IF(A2:A="", "", SUMIF(\\'03_LOTES\\'!B:B, F2:F, \\'03_LOTES\\'!E:E)))'
    array_formula_nombre = '=ARRAYFORMULA(IF(A2:A="", "", A2:A & " | " & B2:B & " | " & C2:C))'
    
    ws_med.write_formula(1, 4, array_formula_stock, cell_format)
    ws_med.write_formula(1, 5, array_formula_nombre, cell_format)"""

code = code.replace(old_med_loop, new_med_loop)

with open('generador_sifarmu_actualizado.py', 'w', encoding='utf-8') as f:
    f.write(code)

