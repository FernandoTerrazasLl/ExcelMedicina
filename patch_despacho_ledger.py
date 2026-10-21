import re

with open('generador_sifarmu_actualizado.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Update HOJAS for 00_DESPACHO
old_hojas = """    ("00_DESPACHO", [
        "Caja de Salida Rápida"
    ]),"""
new_hojas = """    ("00_DESPACHO", [
        "Fecha", "Medicamento", "Cantidad", "Destinatario", "Ubicación", "Estado Acción"
    ]),"""
code = code.replace(old_hojas, new_hojas)

# 2. Update BUILD 00_DESPACHO block
# Actually, the BUILD 00_DESPACHO block has an if-condition in main():
# if name == "00_DESPACHO": ws.write(0, 0, headers[0], header_format); continue
# We need to remove that special case so it builds headers normally!

code = code.replace("""        if name == "00_DESPACHO":
            ws.write(0, 0, headers[0], header_format)
            continue""", "")

# Find the BUILD 00_DESPACHO section and replace it with validations
old_build_block = """    # BUILD 00_DESPACHO (Simple Sheet Trigger)
    ws_desp = sheets["00_DESPACHO"]
    ws_desp.set_column('A:Z', 25)
    title_fmt = workbook.add_format({'bold': True, 'font_size': 18, 'font_color': COLOR_PRIMARIO, 'align': 'center', 'valign': 'vcenter'})
    label_fmt = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'right', 'valign': 'vcenter'})
    input_fmt = workbook.add_format({'border': 1, 'bg_color': '#FFFFE0', 'font_size': 12, 'valign': 'vcenter'})
    trigger_fmt = workbook.add_format({'border': 2, 'bold': True, 'bg_color': '#E8F5E9', 'font_color': '#2E7D32', 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
    
    ws_desp.merge_range('B2:C2', 'REGISTRO DE SALIDAS', title_fmt)
    
    ws_desp.write('B4', 'Buscar Medicamento:', label_fmt)
    ws_desp.write('C4', '', input_fmt)
    ws_desp.data_validation('C4', {'validate': 'list', 'source': "='02_MEDICAMENTOS'!$F$2:$F$1000"})
    
    ws_desp.write('B5', 'Cantidad a Retirar:', label_fmt)
    ws_desp.write('C5', '', input_fmt)
    
    ws_desp.write('B6', 'Nombre Persona (Opcional):', label_fmt)
    ws_desp.write('C6', '', input_fmt)
    
    ws_desp.write('B7', 'Ubicación (Opcional):', label_fmt)
    ws_desp.write('C7', '', input_fmt)
    
    ws_desp.write('B9', 'Ejecutar Salida:', label_fmt)
    ws_desp.write('C9', '-- PENDIENTE --', trigger_fmt)
    ws_desp.data_validation('C9', {'validate': 'list', 'source': ['-- PENDIENTE --', '🚀 PROCESAR AHORA']})
    
    ws_desp.write('E9', '<-- Selecciona "PROCESAR AHORA" para descontar el stock mágicamente')"""

new_build_block = """    # BUILD 00_DESPACHO (LEDGER TABLE)
    ws_desp = sheets["00_DESPACHO"]
    # We already have headers from the standard loop
    ws_desp.set_column('A:A', 15, date_format)
    ws_desp.set_column('B:B', 40, cell_format)
    ws_desp.set_column('C:E', 20, cell_format)
    ws_desp.set_column('F:F', 25, cell_format)
    
    # Validaciones para el Ledger
    ws_desp.data_validation('A2:A1000', {
        'validate': 'date',
        'criteria': 'between',
        'minimum': datetime(2000, 1, 1),
        'maximum': datetime(2100, 12, 31)
    })
    ws_desp.data_validation('B2:B1000', {'validate': 'list', 'source': "='02_MEDICAMENTOS'!$F$2:$F$1000"})
    ws_desp.data_validation('F2:F1000', {'validate': 'list', 'source': ['-- PENDIENTE --', 'APLICAR SALIDA']})
    
    # Formato condicional para la columna Estado (F)
    ws_desp.conditional_format('F2:F1000', {
        'type': 'cell',
        'criteria': 'equal to',
        'value': '"✅ PROCESADO"',
        'format': workbook.add_format({'bg_color': '#C8E6C9', 'font_color': '#1B5E20'})
    })
    ws_desp.conditional_format('F2:F1000', {
        'type': 'cell',
        'criteria': 'equal to',
        'value': '"❌ STOCK INSUFICIENTE"',
        'format': workbook.add_format({'bg_color': '#FFCDD2', 'font_color': '#B71C1C'})
    })
    ws_desp.conditional_format('F2:F1000', {
        'type': 'cell',
        'criteria': 'equal to',
        'value': '"APLICAR SALIDA"',
        'format': workbook.add_format({'bg_color': '#FFF9C4', 'font_color': '#F57F17', 'bold': True})
    })"""

code = code.replace(old_build_block, new_build_block)

with open('generador_sifarmu_actualizado.py', 'w', encoding='utf-8') as f:
    f.write(code)

