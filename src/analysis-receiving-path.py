import pandas as pd
import sys

"""
    Planilha chega em .xlsx com o Header.
    ID, Número, Respondida, Data do Envio, Data da Resposta, Tempo de Resposta, Estado.

    O que quero Análisar?

    Conseguir automatizar a análise de dados, para que eu possa ter uma visão geral da performance da campanha.

    Este script é uma cópia do outro. Mas utiliza do terminal para receber o caminho e tentar ir lá buscar o campinho.
"""



# Carrega os dados em um DataFrame do Pandas
def load_data(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        return df
    except Exception as e:
        print(f"Erro ao carregar o arquivo: {e}")
        sys.exit(1) # Sai do programa pois não da para continuar sem os dados


# Analisa os dados e gera um relatório
def analyze_data(df: pd.DataFrame) -> str:
    total_mensagens = len(df)
    mensagens_respondidas = df[df['Respondida'] == 'Sim']

    taxa_resposta = (len(mensagens_respondidas) / total_mensagens) * 100

    count_estados = df['Estado'].value_counts()
    return f"--- *Relatório de Análise de Campanha* ---\n*Dados principais*\nTotal de Mensagens Enviadas: {total_mensagens}\nTotal de Mensagens Respondidas: {len(mensagens_respondidas)}\nTaxa de Resposta: {taxa_resposta:.2f}%\n\n*Distribuição do envio*\n{count_estados.to_string()}"

if __name__ == "__main__":
    input_path = input("Digite o caminho para o arquivo .xlsx: ")
    df = load_data(input_path)
    report = analyze_data(df)
    print(report)