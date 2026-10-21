with open('gas_scripts/sifarmu_optimizado.js', 'r', encoding='utf-8') as f:
    code = f.read()

# Fix 1: Principio Activo is now at index 4 instead of 6
code = code.replace("nombreMedicamento = medsData[i][6];", "nombreMedicamento = medsData[i][4];")

# Fix 2: Add trim() to codigo search logic
code = code.replace("if(lotesData[i][1] == codigo &&", "if(String(lotesData[i][1]).trim() === codigo.trim() &&")

with open('gas_scripts/sifarmu_optimizado.js', 'w', encoding='utf-8') as f:
    f.write(code)

