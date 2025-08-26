# Conversor de Excel para Excel + CSV

Este script em Python tem como objetivo **desvincular um arquivo Excel com várias abas** e salvar cada aba como arquivos separados em dois formatos:

- **Excel (.xlsx)** – cada aba vira um novo arquivo Excel individual.  
- **CSV (.csv)** – cada aba também é convertida para CSV.

---

## Funcionalidades

1. **Leitura de um arquivo Excel único**  
   - O script abre o arquivo principal especificado no caminho `file_path`.

2. **Separação por abas**  
   - Cada aba do Excel é lida individualmente.  
   - É criado um nome de arquivo "seguro", removendo espaços e caracteres especiais.

3. **Exportação em múltiplos formatos**  
   - Para cada aba:  
     - Salva como arquivo **Excel individual** na pasta `excel_export`.  
     - Salva também como arquivo **CSV** na pasta `csv_export`.

4. **Criação automática de diretórios de saída**  
   - Se as pastas `excel_export` e `csv_export` não existirem, elas são criadas automaticamente.

---

## Como usar

1. Instale as dependências (caso não tenha):  
   ```bash
   pip install pandas openpyxl
2. Ajuste o nome do arquivo Excel principal no script:
   file_path = "PR - NAI (Finalizado).xlsx"
3. Execute o script:
   ```bash
   python excel_to_csv.py
4. Verifique os resultados nas pastas:
   - excel_export/ - arquivos Excel individuais.
   - csv_export/ - arquivos CSV individuais.

---
   
