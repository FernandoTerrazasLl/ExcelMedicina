import re

with open('generador_sifarmu_actualizado.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Update Date Validation for 00_DESPACHO (Column A)
old_val_desp = """    ws_desp.data_validation('A2:A1000', {
        'validate': 'date',
        'criteria': 'between',
        'minimum': datetime(2000, 1, 1),
        'maximum': datetime(2100, 12, 31)
    })"""
new_val_desp = """    ws_desp.data_validation('A2:A1000', {
        'validate': 'date',
        'criteria': 'between',
        'minimum': datetime(2000, 1, 1),
        'maximum': datetime(2100, 12, 31),
        'input_title': '📆 Ingresar Fecha',
        'input_message': 'Haz DOBLE CLIC aquí para abrir el calendario.\\nO escribe en formato: DD/MM/AAAA',
        'error_title': '❌ Fecha Inválida',
        'error_message': 'Por favor usa el formato DD/MM/AAAA o haz doble clic para usar el calendario.'
    })"""
code = code.replace(old_val_desp, new_val_desp)

# 2. Update Date Validation for 03_LOTES (Columns G and H)
old_val_lotes = """    ws_lotes.data_validation('G2:H1000', {
        'validate': 'date',
        'criteria': 'between',
        'minimum': datetime(2000, 1, 1),
        'maximum': datetime(2100, 12, 31)
    })"""
new_val_lotes = """    ws_lotes.data_validation('G2:H1000', {
        'validate': 'date',
        'criteria': 'between',
        'minimum': datetime(2000, 1, 1),
        'maximum': datetime(2100, 12, 31),
        'input_title': '📆 Ingresar Fecha',
        'input_message': 'Haz DOBLE CLIC aquí para abrir el calendario.\\nO escribe en formato: DD/MM/AAAA',
        'error_title': '❌ Fecha Inválida',
        'error_message': 'Por favor usa el formato DD/MM/AAAA o haz doble clic para usar el calendario.'
    })"""
code = code.replace(old_val_lotes, new_val_lotes)

with open('generador_sifarmu_actualizado.py', 'w', encoding='utf-8') as f:
    f.write(code)

