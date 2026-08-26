from datetime import datetime
import os
import numpy as np
import pandas as pd


def criar_dados_exemplo(pasta_origem):
    """Cria planilhas de exemplo (Filial A, Filial B, Filial C) para simular o cenário do cliente."""
    if not os.path.exists(pasta_origem):
        os.makedirs(pasta_origem)

    dados_filial_a = {
        "Data": ["2026-08-10", "2026-08-10", "2026-08-11"],
        "Produto": ["Teclado", "Mouse", "Monitor"],
        "Quantidade": [5, 10, 2],
        "Valor_Unitario": [150.00, 80.00, 900.00],
    }

    dados_filial_b = {
        "Data": ["2026-08-10", "2026-08-11"],
        "Produto": ["Headset", "Teclado"],
        "Quantidade": [3, 4],
        "Valor_Unitario": [250.00, 150.00],
    }

    pd.DataFrame(dados_filial_a).to_excel(
        os.path.join(pasta_origem, "Vendas_Filial_A.xlsx"), index=False
    )
    pd.DataFrame(dados_filial_b).to_excel(
        os.path.join(pasta_origem, "Vendas_Filial_B.xlsx"), index=False
    )
    print("📁 Planilhas de exemplo criadas na pasta 'vendas_diarias'.")


def consolidar_planilhas_vendas():
    pasta_vendas = "vendas_diarias"
    arquivo_saida = "Relatorio_Consolidado_Vendas.xlsx"

    # Gera arquivos de teste caso a pasta não exista
    if not os.path.exists(pasta_vendas):
        criar_dados_exemplo(pasta_vendas)

    print("\n🤖 Iniciando consolidação de planilhas...")

    lista_dfs = []

    # 1. Percorre todos os arquivos dentro da pasta
    for arquivo in os.listdir(pasta_vendas):
        if arquivo.endswith(".xlsx") or arquivo.endswith(".xls"):
            caminho_completo = os.path.join(pasta_vendas, arquivo)

            print(f"📄 Processando: {arquivo}...")

            # 2. Lê a planilha atual
            df = pd.read_excel(caminho_completo)

            # Adiciona uma coluna identificando de qual arquivo/filial vieram os dados
            nome_filial = arquivo.replace("Vendas_", "").replace(".xlsx", "")
            df["Origem_Arquivo"] = nome_filial

            lista_dfs.append(df)

    if not lista_dfs:
        print("⚠️ Nenhuma planilha encontrada na pasta!")
        return

    # 3. Unifica todas as planilhas em um único DataFrame
    df_consolidado = pd.concat(lista_dfs, ignore_index=False)

    # 4. Processamento e cálculo automatizado
    df_consolidado["Faturamento_Total"] = (
        df_consolidado["Quantidade"] * df_consolidado["Valor_Unitario"]
    )

    # 5. Salva o resultado final no Excel com abas organizadas
    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        # Aba 1: Todos os dados unificados
        df_consolidado.to_excel(
            writer, sheet_name="Dados Unificados", index=False
        )

        # Aba 2: Resumo consolidado por Filial/Origem
        resumo_filial = (
            df_consolidado.groupby("Origem_Arquivo")
            .agg({"Quantidade": "sum", "Faturamento_Total": "sum"})
            .reset_index()
        )
        resumo_filial.to_excel(
            writer, sheet_name="Resumo por Filial", index=False
        )

    print(
        f"\n✅ Consolidação concluída! Arquivo '{arquivo_saida}' gerado com sucesso."
    )


if __name__ == "__main__":
    consolidar_planilhas_vendas()