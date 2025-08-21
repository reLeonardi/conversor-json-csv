import pandas as pd
import os

# Nome do arquivo Excel
file_path = "PR - NAI (Finalizado).xlsx"

# Pasta de saída
output_dir = "csv_export"
os.makedirs(output_dir, exist_ok=True)

# Carrega todas as abas
xls = pd.ExcelFile(file_path)

# Converte cada aba em CSV
for sheet in xls.sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet)
    safe_name = sheet.strip().replace(" ", "_").replace("\xa0", "")
    csv_path = os.path.join(output_dir, f"{safe_name}.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

print("Conversão concluída! Arquivos salvos em:", output_dir)
