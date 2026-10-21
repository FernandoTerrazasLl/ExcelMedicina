import xlsxwriter
from datetime import datetime

with open('medicamentos_extraidos.txt', 'r', encoding='utf-8') as f:
    lines = f.readlines()

meds = []
for line in lines[2:]:
    parts = line.strip().split('\t')
    if len(parts) >= 3:
        nombre = parts[0].strip()
        forma = parts[1].strip()
        conc = parts[2].strip()
        if nombre:
            cat = "Otros"
            nl = nombre.lower()
            if any(k in nl for k in ["amoxi", "azitro", "cefa", "cefo", "ceft", "cipro", "clinda", "cloran", "cloxa", "doxi", "eritro", "genta", "levo", "penicilina", "tetra"]):
                cat = "Antibióticos"
            elif any(k in nl for k in ["atenolol", "losartan", "enalapril", "amlodipina", "nifedipino", "hidroclorotiazida", "bisoprolol", "carvedilol"]):
                cat = "Antihipertensivos"
            elif any(k in nl for k in ["glibenclamida", "metformina", "insulina"]):
                cat = "Hipoglucemiantes"
            elif any(k in nl for k in ["paracetamol", "ibuprofeno", "diclofenaco", "ketorolaco", "tramadol", "aspirina", "acetil salicílico"]):
                cat = "Analgésicos y Antiinflamatorios"
            elif any(k in nl for k in ["clotrimazol", "fluconazol", "nistatina", "terbinafina"]):
                cat = "Antifúngicos"
            elif any(k in nl for k in ["vitamina", "hierro", "sulfato ferroso", "ácido fólico", "complejo", "calcio", "zinc", "micronutrientes"]):
                cat = "Vitaminas y Suplementos"
            elif any(k in nl for k in ["omeprazol", "ranitidina", "bicarbonato", "hidróxido", "sucralfato"]):
                cat = "Gastrointestinales"
                
            meds.append((nombre, forma, conc, cat))

meds.sort(key=lambda x: x[0])

code = """import xlsxwriter
from datetime import datetime

COLOR_PRIMARIO = "#1565C0"
COLOR_BLANCO = "#FFFFFF"
COLOR_GRIS_FONDO = "#F5F5F5"

CATEGORIAS = [
    ["CAT001", "Analgésicos y Antiinflamatorios", "Medicamentos para el dolor e inflamación", "ACTIVO"],
    ["CAT002", "Antibióticos", "Tratamiento de infecciones bacterianas", "ACTIVO"],
    ["CAT003", "Antihipertensivos", "Control de la presión arterial", "ACTIVO"],
    ["CAT004", "Hipoglucemiantes", "Control de la diabetes", "ACTIVO"],
    ["CAT005", "Antifúngicos", "Tratamiento de hongos", "ACTIVO"],
    ["CAT006", "Vitaminas y Suplementos", "Nutrición y anemia", "ACTIVO"],
    ["CAT007", "Gastrointestinales", "Protección gástrica", "ACTIVO"],
    ["CAT008", "Otros", "Otras especialidades", "ACTIVO"]
]

MEDICAMENTOS = [
"""
for idx, (n, f, c, cat) in enumerate(meds, 1):
    cod = f"MED-{idx:05d}"
    code += f'    ["{cod}", "{cat}", "{n}", "{c}", "{f}"],\n'
code += "]\n\n"

code += """
HOJAS = [
    ("00_DESPACHO", [
        "Caja de Salida Rápida"
    ]),
    ("01_CONFIG", ["Parámetro", "Valor"]),
    ("02_MEDICAMENTOS", [
        "Principio Activo", "Concentración", "Forma Farmacéutica", "Categoría",
        "Stock Actual", "Nombre Completo", "Código", "Código LINAME", "Estado"
    ]),
    ("03_LOTES", [
        "ID", "Medicamento", "Lote", "Cantidad Inicial", "Disponible", "Proveedor",
        "Fecha Fabricación", "Fecha Vencimiento", "Costo Unitario", "Ubicación", "Estado"
    ]),
    ("04_MOVIMIENTOS", [
        "ID", "Fecha", "Hora", "Tipo", "Medicamento", "Cantidad", "Lote", "Destino", "Responsable", "Observaciones"
    ]),
    ("05_KARDEX", [
        "ID Movimiento", "Fecha", "Hora", "Medicamento", "Tipo", "Cantidad",
        "Stock Anterior", "Stock Nuevo", "Lote", "Responsable", "Observaciones"
    ]),
    ("06_BRIGADAS", [
        "ID", "Fecha", "Nombre Brigada", "Lugar", "Responsable", "Medicamento",
        "Cantidad", "Lote", "Paciente", "Observaciones", "Estado"
    ]),
    ("07_KITS", [
        "ID", "Kit", "Medicamento", "Cantidad", "Categoría", "Estado"
    ]),
    ("08_PROVEEDORES", [
        "ID", "Proveedor", "Contacto", "Teléfono", "Correo", "Dirección", "Observaciones"
    ]),
    ("09_DONANTES", [
        "ID", "Donante", "Contacto", "Teléfono", "Correo", "Observaciones"
    ]),
    ("10_USUARIOS", [
        "ID", "Nombre", "Correo", "Cargo", "Rol", "Estado"
    ]),
    ("11_CATEGORIAS", [
        "ID", "Categoría", "Descripción", "Estado"
    ]),
    ("12_UBICACIONES", [
        "ID", "Ubicación", "Descripción"
    ]),
    ("13_DASHBOARD", [
        "Indicador", "Valor"
    ]),
    ("14_REPORTES", [
        "Fecha", "Tipo", "Usuario", "Observaciones"
    ]),
    ("15_LOG", [
        "Fecha", "Hora", "Usuario", "Acción", "Detalle"
    ])
]

def main():
    workbook = xlsxwriter.Workbook('SIFARMU_Profesional.xlsx')
    
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': COLOR_PRIMARIO,
        'font_color': COLOR_BLANCO,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })
    
    cell_format = workbook.add_format({
        'border': 1,
        'valign': 'vcenter'
    })

    date_format = workbook.add_format({
        'border': 1,
        'valign': 'vcenter',
        'num_format': 'dd/mm/yyyy'
    })

    money_format = workbook.add_format({
        'border': 1,
        'valign': 'vcenter',
        'num_format': '#,##0.00'
    })
    
    sheets = {}
    
    for name, headers in HOJAS:
        ws = workbook.add_worksheet(name)
        sheets[name] = ws
        if name == "00_DESPACHO":
            ws.write(0, 0, headers[0], header_format)
            continue
        for col_num, header in enumerate(headers):
            ws.write(0, col_num, header, header_format)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, 1000, len(headers)-1)
        ws.set_column(0, len(headers)-1, 18, cell_format)
    
    ws_cat = sheets["11_CATEGORIAS"]
    for row_num, row_data in enumerate(CATEGORIAS, start=1):
        for col_num, val in enumerate(row_data):
            ws_cat.write(row_num, col_num, val, cell_format)
            
    ws_med = sheets["02_MEDICAMENTOS"]
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
        ws_med.write(row_num, 8, "ACTIVO", cell_format)  # Estado

    ws_med.data_validation('D2:D1000', {'validate': 'list', 'source': "='11_CATEGORIAS'!$B$2:$B$50"})
    ws_med.data_validation('I2:I1000', {'validate': 'list', 'source': ['ACTIVO', 'INACTIVO']})
    
    ws_lotes = sheets["03_LOTES"]
    ws_lotes.set_column('G:H', 15, date_format)
    ws_lotes.set_column('I:I', 15, money_format)
    
    # Validation for Medicamento uses the Nombre Completo column (F)
    ws_lotes.data_validation('B2:B1000', {'validate': 'list', 'source': "='02_MEDICAMENTOS'!$F$2:$F$1000"})
    ws_lotes.data_validation('F2:F1000', {'validate': 'list', 'source': "='08_PROVEEDORES'!$B$2:$B$100"})
    ws_lotes.data_validation('K2:K1000', {'validate': 'list', 'source': ['ACTIVO', 'AGOTADO', 'VENCIDO']})
    
    ws_lotes.data_validation('G2:H1000', {
        'validate': 'date',
        'criteria': 'between',
        'minimum': datetime(2000, 1, 1),
        'maximum': datetime(2100, 12, 31)
    })
    
    ws_lotes.conditional_format('H2:H1000', {'type': 'formula', 'criteria': '=AND(H2<>"", H2<TODAY())', 'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})})
    ws_lotes.conditional_format('H2:H1000', {'type': 'formula', 'criteria': '=AND(H2<>"", H2<TODAY()+90, H2>=TODAY())', 'format': workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})})
    
    # BUILD 00_DESPACHO
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
    ws_desp.set_row(5, 60)
    
    ws_dash = sheets["13_DASHBOARD"]
    ws_dash.set_column('A:A', 30)
    ws_dash.set_column('B:B', 20)
    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'font_color': COLOR_PRIMARIO})
    ws_dash.write('A1', '🏥 SIFARMU v2.0 - Dashboard', title_format)
    
    dash_items = [
        ('Medicamentos Registrados', "=COUNTA('02_MEDICAMENTOS'!A:A)-1"),
        ('Lotes Activos', "=COUNTIF('03_LOTES'!K:K, \\"ACTIVO\\")"),
        ('Stock Total Unidades', "=SUM('03_LOTES'!E:E)"),
        ('Lotes Vencidos', "=COUNTIFS('03_LOTES'!H:H, \\"<\\"&TODAY(), '03_LOTES'!H:H, \\"<>\\")")
    ]
    
    for idx, (label, formula) in enumerate(dash_items, start=3):
        ws_dash.write(idx, 0, label, workbook.add_format({'bold': True, 'border': 1, 'bg_color': COLOR_GRIS_FONDO}))
        ws_dash.write(idx, 1, formula, cell_format)

    workbook.close()
    print("¡Archivo generado exitosamente con la nueva UX/UI!")

if __name__ == "__main__":
    main()
"""

with open('generador_sifarmu_actualizado.py', 'w', encoding='utf-8') as f:
    f.write(code)

