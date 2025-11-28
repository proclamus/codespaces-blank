import os
import sys
import pandas as pd

ARQUIVO = "vendas_mercado_basico.xlsx"


def carregar_arquivo():
    """Tenta carregar o arquivo Excel e retorna um DataFrame ou None se falhar."""
    if not os.path.exists(ARQUIVO):
        print(f"Erro: arquivo '{ARQUIVO}' não encontrado no diretório.")
        return None
    try:
        return pd.read_excel(ARQUIVO)
    except Exception as e:
        print(f"Erro ao abrir o arquivo: {e}")
        return None


def buscar_produto(df):
    """Busca por produtos no DataFrame com base em termo parcial (case-insensitive)."""
    if "Produto" not in df.columns:
        print("Coluna 'Produto' não encontrada no DataFrame.")
        return

    while True:
        termo = input("\nDigite o nome (ou parte) do produto para buscar: ").strip()
        if not termo:
            print("Termo inválido. Tente novamente.")
            continue

        resultados = df[df["Produto"].str.contains(termo, case=False, na=False)]

        if resultados.empty:
            print("\nNenhum produto encontrado.")
        else:
            print("\n=== Resultados da busca ===")
            print(resultados.to_string(index=False))

        print("\n1 - Buscar próximo")
        print("2 - Voltar ao menu inicial")
        print("3 - Sair")
        opc = input("Escolha: ")

        if opc == "1":
            continue
        elif opc == "2":
            return
        elif opc == "3":
            print("Encerrando programa...")
            sys.exit(0)
        else:
            print("Opção inválida.")


def menu_principal(df):
    """Loop principal do menu do programa."""
    while True:
        print("\n=== MENU PRINCIPAL ===")
        print("1 - Buscar produto")
        print("2 - Sair")
        opcao = input("Escolha: ")

        if opcao == "1":
            buscar_produto(df)
        elif opcao == "2":
            print("Encerrando programa...")
            break
        else:
            print("Opção inválida!")


def main():
    df = carregar_arquivo()
    if df is None:
        return
    menu_principal(df)


if __name__ == "__main__":
    main()