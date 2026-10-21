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
            # Simple keyword categorization
            cat = "Otros"
            nl = nombre.lower()
            if any(k in nl for k in ["amoxi", "azitro", "cefa", "cefo", "ceft", "cipro", "clinda", "cloran", "cloxa", "doxi", "eritro", "genta", "levo", "penicilina", "tetra"]):
                cat = "Antibióticos"
            elif any(k in nl for k in ["atenolol", "losartan", "enalapril", "amlodipina", "nifedipino", "hidroclorotiazida"]):
                cat = "Antihipertensivos"
            elif any(k in nl for k in ["glibenclamida", "metformina", "insulina"]):
                cat = "Hipoglucemiantes"
            elif any(k in nl for k in ["paracetamol", "ibuprofeno", "diclofenaco", "ketorolaco", "tramadol", "aspirina", "acetil salicílico"]):
                cat = "Analgésicos y Antiinflamatorios"
            elif any(k in nl for k in ["clotrimazol", "fluconazol", "nistatina", "terbinafina"]):
                cat = "Antifúngicos"
            elif any(k in nl for k in ["vitamina", "hierro", "sulfato ferroso", "ácido fólico"]):
                cat = "Vitaminas y Suplementos"
            elif any(k in nl for k in ["omeprazol", "ranitidina", "bicarbonato", "hidróxido"]):
                cat = "Gastrointestinales"
                
            meds.append((nombre, forma, conc, cat))

# Sort alphabetically by name
meds.sort(key=lambda x: x[0])

# Generate python code for the generator
code = """import xlsxwriter

COLOR_PRIMARIO = "#1565C0"
COLOR_BLANCO = "#FFFFFF"

# Categorías
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
    code += f'    ["{cod}", "", "", "{cat}", "", "{n}", "", "{c}", "{f}", "", "", "", "", "", "ACT"],\n'
code += "]\n\n"

# Append the rest of the generator script
with open('generador_sifarmu.py', 'r', encoding='utf-8') as f:
    old_code = f.read()

# Replace CATEGORIAS_DATA
old_code = re.sub(r'CATEGORIAS_DATA = \[.*?\]\n', '', old_code, flags=re.DOTALL)

# Find where to insert MEDICAMENTOS write logic
injection = """
    # 02_MEDICAMENTOS (Data Inyectada del TSV)
    ws_med = sheets["02_MEDICAMENTOS"]
    for row_num, m_data in enumerate(MEDICAMENTOS, start=1):
        # Columnas: ID(0), Código(1), LINAME(2), ATC(3), Cat(4), Subcat(5), Princ.Activo(6), Nombre(7), Conc(8), Forma(9), Pres(10), Unid(11), Via(12), Receta(13), Frio(14), Estado(15)
        # Llenamos las necesarias
        ws_med.write(row_num, 0, f"ID-{row_num:04d}", cell_format)
        ws_med.write(row_num, 1, m_data[0], cell_format) # Codigo
        ws_med.write(row_num, 4, m_data[3], cell_format) # Categoria
        ws_med.write(row_num, 6, m_data[5], cell_format) # Princ. Activo
        ws_med.write(row_num, 8, m_data[7], cell_format) # Concentracion
        ws_med.write(row_num, 9, m_data[8], cell_format) # Forma
        ws_med.write(row_num, 15, "ACTIVO", cell_format) # Estado
        
    # Dropdowns para Medicamentos...
"""
old_code = old_code.replace('ws_med = sheets["02_MEDICAMENTOS"]', injection)
old_code = old_code.replace('CATEGORIAS_DATA', 'CATEGORIAS')

with open('generador_sifarmu_actualizado.py', 'w', encoding='utf-8') as f:
    f.write(code + old_code)
