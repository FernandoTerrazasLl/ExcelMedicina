import json
import re
import os

FORMAS = [
    "Cápsula blanda", "Cartucho dental", "Comprimido ranurado", "Comprimido sublingual",
    "Comprimido vaginal", "Crema vaginal", "Emulsión inyectable", "Emulsión oral",
    "Gotas óticas", "Implante subdérmico", "Polvo oral", "Polvo para enema",
    "Polvo para inyectable", "Polvo para solución oral", "Polvo para suspensión oral",
    "Pomada oftálmica", "Solución acuosa", "Solución hidroalcohólica", "Solución nasal",
    "Solución oftálmica", "Solución oral gotas", "Solución oral", "Solución para atomización",
    "Solución para nebulización", "Solución parenteral de gran volúmen", "Solución parenteral de gran volumen",
    "Solución Tópica", "Suspensión oral", "Ungüento oftálmico", "Solución parenteral",
    "Cápsula o Comprimido", "Cápsula", "Comprimido", "Crema", "Gas", "Gel", "Gotas", "Granulado",
    "Inyectable", "Jarabe", "Jalea", "Loción", "Óvulo", "Pasta o Pomada", "Pasta", "Polvo", "Solución", "Supositorio",
    "Suspensión o Jarabe", "Suspensiòn", "Suspensión", "Ungüento o Pomada", "Ungüento o crema", "Ungüento",
    "Crema o Pomada", "Crema dérmica", "Crema, Pomada o Gel", "Plegable Frasco Vial", "Pleglabe frasco vial"
]

def main():
    with open('liname_raw.txt', 'r', encoding='utf-8') as f:
        ocr_text = f.read()

    medicamentos = []
    lines = ocr_text.split('\n')
    in_list = False
    
    # Allow some leniency in the regex spacing
    regex = re.compile(r'^([A-Z])\s+(\d{2})\s+(\d{2})\s+(.+?)\s+([A-Z0-9*]{4,7})(?:\s+R)?$')
    
    for line in lines:
        if '==Start of OCR for page 23==' in line:
            in_list = True
        if '==End of OCR for page 62==' in line:
            in_list = False
            break
            
        if not in_list:
            continue
            
        line = line.strip()
        match = regex.match(line)
        if match:
            letra = match.group(1)
            num1 = match.group(2)
            num2 = match.group(3)
            middle = match.group(4)
            atq = match.group(5)
            
            codigo = f"{letra}{num1}{num2}"
            
            forma_encontrada = ""
            for forma in FORMAS:
                if forma in middle or forma.lower() in middle.lower():
                    forma_encontrada = forma
                    break
                    
            if forma_encontrada:
                idx = middle.rfind(forma_encontrada)
                if idx == -1:
                    idx = middle.lower().rfind(forma_encontrada.lower())
                    forma_encontrada = middle[idx:idx+len(forma_encontrada)]
                
                nombre = middle[:idx].strip()
                concentracion = middle[idx+len(forma_encontrada):].strip()
            else:
                nombre = middle
                concentracion = ""
                forma_encontrada = "No definida"
                
            medicamentos.append({
                "codigo": codigo,
                "letra": letra,
                "nombre": nombre,
                "forma": forma_encontrada,
                "concentracion": concentracion,
                "atq": atq
            })

    with open('liname_data.json', 'w', encoding='utf-8') as f:
        json.dump(medicamentos, f, ensure_ascii=False, indent=2)
        
    print(f"Extraídos {len(medicamentos)} medicamentos exitosamente.")

if __name__ == "__main__":
    main()
