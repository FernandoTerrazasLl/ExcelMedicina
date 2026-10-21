import re

with open('medicamentos_extraidos.txt', 'r', encoding='utf-8') as f:
    lines = f.readlines()

meds = []
for line in lines[2:]:  # skip header
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
    ("01_CONFIG", ["Parámetro", "Valor"]),
    ("02_MEDICAMENTOS", [
        "ID", "Código", "Código LINAME", "Código ATC", "Categoría", "Subcategoría",
        "Principio Activo", "Nombre Comercial", "Concentración", "Forma Farmacéutica",
        "Presentación", "Unidad", "Vía", "Requiere Receta", "Cadena de Frío", "Estado"
    ]),
    ("03_LOTES", [
        "ID", "Código Medicamento", "Lote", "Cantidad Inicial", "Disponible", "Proveedor",
        "Fecha Fabricación", "Fecha Vencimiento", "Costo Unitario", "Ubicación", "Estado"
    ]),
    ("04_MOVIMIENTOS", [
        "ID", "Fecha", "Hora", "Tipo", "Código", "Cantidad", "Lote", "Destino", "Responsable", "Observaciones"
    ]),
    ("05_KARDEX", [
        "ID Movimiento", "Fecha", "Hora", "Código", "Medicamento", "Tipo", "Cantidad",
        "Stock Anterior", "Stock Nuevo", "Lote", "Responsable", "Observaciones"
    ]),
    ("06_BRIGADAS", [
        "ID", "Fecha", "Nombre Brigada", "Lugar", "Responsable", "Código Medicamento",
        "Cantidad", "Lote", "Paciente", "Observaciones", "Estado"
    ]),
    ("07_KITS", [
        "ID", "Kit", "Código Medicamento", "Nombre Medicamento", "Cantidad", "Categoría", "Estado"
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
        ws_med.write(row_num, 0, f"ID-{row_num:04d}", cell_format)
        ws_med.write(row_num, 1, m_data[0], cell_format)
        ws_med.write(row_num, 4, m_data[1], cell_format)
        ws_med.write(row_num, 6, m_data[2], cell_format)
        ws_med.write(row_num, 8, m_data[3], cell_format)
        ws_med.write(row_num, 9, m_data[4], cell_format)
        ws_med.write(row_num, 15, "ACTIVO", cell_format)

    ws_med.data_validation('E2:E1000', {'validate': 'list', 'source': "='11_CATEGORIAS'!$B$2:$B$50"})
    ws_med.data_validation('P2:P1000', {'validate': 'list', 'source': ['ACTIVO', 'INACTIVO']})
    
    ws_lotes = sheets["03_LOTES"]
    ws_lotes.set_column('G:H', 15, date_format)
    ws_lotes.set_column('I:I', 15, money_format)
    ws_lotes.data_validation('F2:F1000', {'validate': 'list', 'source': "='08_PROVEEDORES'!$B$2:$B$100"})
    ws_lotes.data_validation('K2:K1000', {'validate': 'list', 'source': ['ACTIVO', 'AGOTADO', 'VENCIDO']})
    ws_lotes.conditional_format('H2:H1000', {'type': 'formula', 'criteria': '=AND(H2<>"", H2<TODAY())', 'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})})
    ws_lotes.conditional_format('H2:H1000', {'type': 'formula', 'criteria': '=AND(H2<>"", H2<TODAY()+90, H2>=TODAY())', 'format': workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})})
    
    ws_dash = sheets["13_DASHBOARD"]
    ws_dash.set_column('A:A', 30)
    ws_dash.set_column('B:B', 20)
    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'font_color': COLOR_PRIMARIO})
    ws_dash.write('A1', '🏥 SIFARMU v1.0 - Dashboard', title_format)
    
    dash_items = [
        ('Medicamentos Registrados', "=COUNTA('02_MEDICAMENTOS'!A:A)-1"),
        ('Lotes Activos', "=COUNTIF('03_LOTES'!K:K, \\"ACTIVO\\")"),
        ('Stock Total Unidades', "=SUM('03_LOTES'!E:E)"),
        ('Lotes Vencidos', "=COUNTIF('03_LOTES'!H:H, \\"<\\"&TODAY())")
    ]
    
    for idx, (label, formula) in enumerate(dash_items, start=3):
        ws_dash.write(idx, 0, label, workbook.add_format({'bold': True, 'border': 1, 'bg_color': COLOR_GRIS_FONDO}))
        ws_dash.write(idx, 1, formula, cell_format)

    workbook.close()
    print("¡Archivo generado exitosamente con la base inyectada!")

if __name__ == "__main__":
    main()
"""

with open('generador_sifarmu_actualizado.py', 'w', encoding='utf-8') as f:
    f.write(code)
