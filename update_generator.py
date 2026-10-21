import re

with open('generador_sifarmu_actualizado.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Replace 00_DESPACHO generation
old_despacho = """    # BUILD 00_DESPACHO
    ws_desp = sheets["00_DESPACHO"]
    ws_desp.set_column('A:Z', 25)
    title_fmt = workbook.add_format({'bold': True, 'font_size': 24, 'font_color': COLOR_PRIMARIO, 'align': 'center', 'valign': 'vcenter'})
    label_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'right', 'valign': 'vcenter'})
    input_fmt = workbook.add_format({'border': 2, 'bg_color': '#FFFFE0', 'font_size': 14, 'valign': 'vcenter'})
    btn_fmt = workbook.add_format({'bold': True, 'font_size': 16, 'bg_color': '#FF9800', 'font_color': '#FFFFFF', 'border': 2, 'align': 'center', 'valign': 'vcenter'})
    
    ws_desp.merge_range('B2:D2', 'PUNTO DE SALIDA RÁPIDA (FEFO)', title_fmt)
    ws_desp.write('B4', 'Buscar Medicamento:', label_fmt)
    ws_desp.write('C4', '', input_fmt)
    ws_desp.data_validation('C4', {'validate': 'list', 'source': "='02_MEDICAMENTOS'!$F$2:$F$1000"})
    
    ws_desp.write('B6', 'Cantidad a Retirar:', label_fmt)
    ws_desp.write('C6', '', input_fmt)
    
    ws_desp.write('B8', 'Destino/Brigada:', label_fmt)
    ws_desp.write('C8', '', input_fmt)
    
    # Draw a simulated button
    ws_desp.write('E6', 'Hacer clic en el Menú Superior:\\n🏥 SIFARMU > 🚀 REGISTRAR SALIDA', btn_fmt)
    ws_desp.set_row(5, 60)"""

new_despacho = """    # BUILD 00_DESPACHO (Simple Sheet Trigger)
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

code = code.replace(old_despacho, new_despacho)

with open('generador_sifarmu_actualizado.py', 'w', encoding='utf-8') as f:
    f.write(code)

