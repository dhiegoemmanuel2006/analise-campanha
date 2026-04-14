import pandas as pd
import sys
from pathlib import Path

"""
    Planilha chega em .xlsx com o Header.
    ID, Número, Respondida, Data do Envio, Data da Resposta, Tempo de Resposta, Estado.

    O que quero Análisar?

    Conseguir automatizar a análise de dados, para que eu possa ter uma visão geral da performance da campanha.

    Este script é uma cópia do outro. Mas utiliza do terminal para receber o caminho e tentar ir lá buscar o campinho.
"""

REQUIRED_COLUMNS = {"Respondida", "Estado"}


def load_data(path: str) -> pd.DataFrame:
    file_path = Path(path.strip().strip('"').strip("'")).expanduser()

    if not file_path.exists():
        print(f"Erro: arquivo não encontrado em '{file_path}'.")
        sys.exit(1)

    if file_path.suffix.lower() not in {".xlsx", ".xls"}:
        print(f"Erro: formato inválido ('{file_path.suffix}'). Use .xlsx ou .xls.")
        sys.exit(1)

    try:
        return pd.read_excel(file_path)
    except Exception as e:
        print(f"Erro ao carregar o arquivo: {e}")
        sys.exit(1)


def analyze_data(df: pd.DataFrame) -> str:
    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        return f"Erro: colunas obrigatórias ausentes na planilha: {', '.join(sorted(missing))}."

    total_mensagens = len(df)
    if total_mensagens == 0:
        return "Planilha vazia: não há mensagens para analisar."

    respondidas_series = df["Respondida"].astype(str).str.strip().str.casefold()
    mensagens_respondidas = int((respondidas_series == "sim").sum())
    taxa_resposta = (mensagens_respondidas / total_mensagens) * 100

    count_estados = df["Estado"].fillna("Desconhecido").value_counts()
    linhas_estados = [
        f"• {estado}: *{qtd}*" for estado, qtd in count_estados.items()
    ] or ["• (sem dados)"]

    linhas = [
        "📊 *Relatório de Análise de Campanha*",
        "",
        "*Dados principais*",
        f"• Mensagens enviadas: *{total_mensagens}*",
        f"• Mensagens respondidas: *{mensagens_respondidas}*",
        f"• Taxa de resposta: *{taxa_resposta:.2f}%*",
        "",
        "*Distribuição por estado*",
        *linhas_estados,
    ]
    return "\n".join(linhas)


if __name__ == "__main__":
    input_path = input("Digite o caminho para o arquivo .xlsx: ")
    df = load_data(input_path)
    report = analyze_data(df)
    print(report)
