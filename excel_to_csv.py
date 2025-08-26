import pandas as pd
import os

# Nome do arquivo Excel
file_path = "PR - NAI (Finalizado).xlsx"

# Pastas de saída
excel_dir = "excel_export"
csv_dir = "csv_export"

os.makedirs(excel_dir, exist_ok=True)
os.makedirs(csv_dir, exist_ok=True)

# Carrega todas as abas
xls = pd.ExcelFile(file_path)

for sheet in xls.sheet_names:
    # Lê cada aba
    df = pd.read_excel(file_path, sheet_name=sheet)

    # Nome seguro do arquivo
    safe_name = sheet.strip().replace(" ", "_").replace("\xa0", "")

    # 1) Salva a aba como Excel individual
    excel_path = os.path.join(excel_dir, f"{safe_name}.xlsx")
    df.to_excel(excel_path, index=False)

    # 2) Salva também como CSV
    csv_path = os.path.join(csv_dir, f"{safe_name}.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

print("Conversão concluída!")
print(f"Arquivos Excel salvos em: {excel_dir}")
print(f"Arquivos CSV salvos em: {csv_dir}")
