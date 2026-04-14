import pandas as pd
from pathlib import Path
import sys

"""
    Planilha chega em .xlsx com o Header.
    ID, Número, Respondida, Data do Envio, Data da Resposta, Tempo de Resposta, Estado.

    O que quero Análisar?

    Conseguir automatizar a análise de dados, para que eu possa ter uma visão geral da performance da campanha.

"""

BASE_PATH = Path(__file__).parent.parent

INPUT_PATH = BASE_PATH / "data" / "Campanha-Novo-Livinho + Grupo.xlsx"

REQUIRED_COLUMNS = {"Respondida", "Estado"}


def load_data(path: Path) -> pd.DataFrame:
    if not path.exists():
        print(f"Erro: arquivo não encontrado em '{path}'.")
        sys.exit(1)

    if path.suffix.lower() not in {".xlsx", ".xls"}:
        print(f"Erro: formato inválido ('{path.suffix}'). Use .xlsx ou .xls.")
        sys.exit(1)

    try:
        return pd.read_excel(path)
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
    largura_estado = max((len(str(e)) for e in count_estados.index), default=0)
    largura_valor = max((len(str(v)) for v in count_estados.values), default=1)
    linhas_estados = "\n".join(
        f"{str(estado).ljust(largura_estado)}  {str(qtd).rjust(largura_valor)}"
        for estado, qtd in count_estados.items()
    ) or "(sem dados)"

    linhas = [
        "📊 *Relatório de Análise de Campanha*",
        "",
        "*Dados principais*",
        f"• Mensagens enviadas: *{total_mensagens}*",
        f"• Mensagens respondidas: *{mensagens_respondidas}*",
        f"• Taxa de resposta: *{taxa_resposta:.2f}%*",
        "",
        "*Distribuição por estado*",
        "```",
        linhas_estados,
        "```",
    ]
    return "\n".join(linhas)

if __name__ == "__main__":
    df = load_data(INPUT_PATH)
    report = analyze_data(df)
    print(report)