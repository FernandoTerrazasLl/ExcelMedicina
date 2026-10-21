import xlsxwriter
from datetime import datetime

# Definiciones de colores
COLOR_PRIMARIO = "#1565C0"
COLOR_SECUNDARIO = "#1976D2"
COLOR_EXITO = "#2E7D32"
COLOR_ALERTA = "#F9A825"
COLOR_ERROR = "#C62828"
COLOR_GRIS_FONDO = "#F5F5F5"
COLOR_BLANCO = "#FFFFFF"

# Estructura de Hojas
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

# Categorías iniciales
CATEGORIAS_DATA = [
    ["CAT001","Analgésicos","Medicamentos para el dolor","ACTIVO"],
    ["CAT002","Antibióticos","Tratamiento de infecciones","ACTIVO"],
    ["CAT003","Antiinflamatorios","AINEs y corticoides","ACTIVO"],
    ["CAT004","Antihipertensivos","Control de la presión arterial","ACTIVO"],
    ["CAT005","Hipoglucemiantes","Control de diabetes","ACTIVO"],
    ["CAT006","Antihistamínicos","Alergias","ACTIVO"],
    ["CAT007","Antiparasitarios","Tratamiento antiparasitario","ACTIVO"],
    ["CAT008","Antifúngicos","Tratamiento antimicótico","ACTIVO"],
    ["CAT009","Vitaminas","Vitaminas y suplementos","ACTIVO"],
    ["CAT010","Material Médico","Material descartable","ACTIVO"],
    ["CAT011","Soluciones","Sueros y soluciones","ACTIVO"],
    ["CAT012","Otros","Otros medicamentos","ACTIVO"]
]

def main():
    print("Iniciando generación de SIFARMU_Profesional.xlsx...")
    workbook = xlsxwriter.Workbook('SIFARMU_Profesional.xlsx')
    
    # Formatos
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
    
    # 1. Crear hojas y encabezados
    for name, headers in HOJAS:
        ws = workbook.add_worksheet(name)
        sheets[name] = ws
        
        # Escribir encabezados
        for col_num, header in enumerate(headers):
            ws.write(0, col_num, header, header_format)
            
        # Congelar primera fila
        ws.freeze_panes(1, 0)
        
        # Auto-filtro
        ws.autofilter(0, 0, 1000, len(headers)-1)
        
        # Ancho de columnas estándar
        ws.set_column(0, len(headers)-1, 18, cell_format)
    
    # Ajustes específicos por hoja
    
    # 11_CATEGORIAS (Data Semilla)
    ws_cat = sheets["11_CATEGORIAS"]
    for row_num, row_data in enumerate(CATEGORIAS_DATA, start=1):
        for col_num, val in enumerate(row_data):
            ws_cat.write(row_num, col_num, val, cell_format)
            
    # 02_MEDICAMENTOS
    ws_med = sheets["02_MEDICAMENTOS"]
    # Dropdown para Categoría (Columna E)
    ws_med.data_validation('E2:E1000', {
        'validate': 'list',
        'source': "='11_CATEGORIAS'!$B$2:$B$50"
    })
    # Dropdown para Estado (Columna P)
    ws_med.data_validation('P2:P1000', {
        'validate': 'list',
        'source': ['ACTIVO', 'INACTIVO']
    })
    
    # 03_LOTES
    ws_lotes = sheets["03_LOTES"]
    ws_lotes.set_column('G:H', 15, date_format) # Fechas
    ws_lotes.set_column('I:I', 15, money_format) # Costo
    
    # Validación Proveedor (Col F)
    ws_lotes.data_validation('F2:F1000', {
        'validate': 'list',
        'source': "='08_PROVEEDORES'!$B$2:$B$100"
    })
    
    # Validación Estado (Col K)
    ws_lotes.data_validation('K2:K1000', {
        'validate': 'list',
        'source': ['ACTIVO', 'AGOTADO', 'VENCIDO']
    })
    
    # Formato condicional: Fecha Vencimiento (Col H)
    # Rojo si Vencido (Fecha < HOY)
    ws_lotes.conditional_format('H2:H1000', {
        'type': 'formula',
        'criteria': '=AND(H2<>"", H2<TODAY())',
        'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    })
    
    # Amarillo si próximo a vencer (Fecha < HOY + 90)
    ws_lotes.conditional_format('H2:H1000', {
        'type': 'formula',
        'criteria': '=AND(H2<>"", H2<TODAY()+90, H2>=TODAY())',
        'format': workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
    })
    
    # 13_DASHBOARD
    ws_dash = sheets["13_DASHBOARD"]
    ws_dash.set_column('A:A', 30)
    ws_dash.set_column('B:B', 20)
    
    # Título Dashboard
    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'font_color': COLOR_PRIMARIO})
    ws_dash.write('A1', '🏥 SIFARMU v1.0 - Dashboard', title_format)
    
    dash_items = [
        ('Medicamentos Registrados', "=COUNTA('02_MEDICAMENTOS'!A:A)-1"),
        ('Lotes Activos', "=COUNTIF('03_LOTES'!K:K, \"ACTIVO\")"),
        ('Stock Total Unidades', "=SUM('03_LOTES'!E:E)"),
        ('Lotes Vencidos', "=COUNTIF('03_LOTES'!H:H, \"<\"&TODAY())")
    ]
    
    for idx, (label, formula) in enumerate(dash_items, start=3):
        ws_dash.write(idx, 0, label, workbook.add_format({'bold': True, 'border': 1, 'bg_color': COLOR_GRIS_FONDO}))
        ws_dash.write(idx, 1, formula, cell_format)

    workbook.close()
    print("¡Archivo generado exitosamente: SIFARMU_Profesional.xlsx!")

if __name__ == "__main__":
    main()
