import pandas as pd
import json
import re
import unicodedata

# --- CONFIGURAÇÃO ---
file_path = "Pronta Referência Operações - TOTVS.xlsx"
output_path = "dados_consolidados.json"
# --------------------

def to_snake_case(texto):
    """
    Função utilitária para limpar e padronizar textos para snake_case.
    Ex: "Nome do Cliente" -> "nome_do_cliente"
    """
    if not isinstance(texto, str):
        return ""
    texto = texto.strip().lower()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r"[^a-z0-9\s_]", "", texto)
    texto = re.sub(r"\s+", "_", texto)
    return texto

def parse_formato_tabela(df):
    """
    PARSER ESPECIALISTA #1: Para abas com formato de tabela (cabeçalho + linhas de dados).
    Ex: 'PR ATENDIMENTO', 'VIP'S'.
    Retorna uma LISTA de dicionários.
    """
    # Encontra a primeira linha que não está completamente vazia para ser o cabeçalho
    header_row_index = -1
    for i, row in df.iterrows():
        if row.notna().any():
            header_row_index = i
            break
            
    if header_row_index == -1:
        return [] # Retorna lista vazia se a aba for vazia

    # Define a linha encontrada como o cabeçalho e remove as linhas acima dela
    df.columns = df.iloc[header_row_index]
    df = df.iloc[header_row_index + 1:].reset_index(drop=True)
    
    # Limpa e padroniza os nomes das colunas
    df.columns = [to_snake_case(col) for col in df.columns]

    # Remove linhas que são completamente vazias
    df.dropna(how='all', inplace=True)

    # Converte o dataframe limpo para uma lista de dicionários
    return df.to_dict(orient='records')

def parse_formato_chave_valor(df):
    """
    PARSER ESPECIALISTA #2: Para abas com formato de ficha (Chave: Valor).
    Ex: 'PR OPERACIONAL', 'MOVIDA'.
    Retorna um ÚNICO dicionário.
    """
    dados = {}
    # Itera sobre cada linha da planilha
    for i in range(len(df)):
        # Itera sobre as colunas em pares (A/B, C/D, E/F, etc.)
        for col_idx in range(0, df.shape[1] - 1, 2):
            chave = df.iloc[i, col_idx]
            valor = df.iloc[i, col_idx + 1]
            
            if pd.notna(chave) and isinstance(chave, str) and chave.strip() != "":
                dados[to_snake_case(chave)] = valor if pd.notna(valor) else None
    return dados

def converter_excel_para_json_final(file_path, output_path):
    """
    Função principal que atua como um ROTEADOR.
    Ela lê o nome de cada aba e decide qual parser usar.
    """
    try:
        excel_file = pd.ExcelFile(file_path)
    except FileNotFoundError:
        print(f"ERRO: Arquivo não encontrado em '{file_path}'")
        return

    sheet_names = excel_file.sheet_names
    resultado_final = {}

    print(f"Iniciando conversão do arquivo '{file_path}'...")
    
    for aba in sheet_names:
        df = pd.read_excel(excel_file, sheet_name=aba, header=None)
        
        # --- O ROTEADOR INTELIGENTE ---
        # Aqui decidimos qual função especialista chamar baseado no nome da aba
        
        nome_aba_padronizado = aba.strip().upper() # Padroniza nome da aba para comparação
        parser_usado = ""
        dados_extraidos = None

        if nome_aba_padronizado in ['PR ATENDIMENTO', "VIP'S", 'VIPS']:
            parser_usado = "Tabela"
            dados_extraidos = parse_formato_tabela(df)
        
        elif nome_aba_padronizado in ['PR OPERACIONAL', 'MOVIDA']:
            parser_usado = "Chave-Valor"
            dados_extraidos = parse_formato_chave_valor(df)
        
        else:
            print(f"⚠️  Aviso: Aba '{aba}' não possui um parser definido. Será ignorada.")
            continue # Pula para a próxima aba
            
        print(f"✅ Aba '{aba}' processada com sucesso usando o parser de '{parser_usado}'.")
        resultado_final[to_snake_case(aba)] = dados_extraidos

    # Salva o resultado final consolidado em um único arquivo JSON
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(resultado_final, f, ensure_ascii=False, indent=2, default=str)

    print(f"\n🚀 Conversão finalizada! JSON consolidado salvo em: {output_path}")

# --- Executa a conversão ---
if __name__ == "__main__":
    converter_excel_para_json_final(file_path, output_path)