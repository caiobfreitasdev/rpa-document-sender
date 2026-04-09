"""
RPA Document Sender — Ponto de entrada principal.

Fluxo:
    1. Autentica UMA vez no Microsoft Graph.
    2. Oferece menu interativo para download e envio.
    3. O token é repassado para cada módulo — sem reautenticação.
"""
from auth import get_graph_token
from sharepoint_dl import iniciar_download
from email_sender import executar_envio_por_regiao


def menu():
    print("\n===== MENU DO ROBÔ =====")
    print("1) Baixar PDFs do SharePoint")
    print("2) Enviar Boletos/NFs por E-mail")
    print("3) Enviar Correção de Competência")
    print("4) Simulação de Envio (sem disparar e-mails)")
    print("5) Sair")


def main():
    print("Autenticando no Microsoft Graph...")
    token = get_graph_token()
    print("Autenticação concluída.\n")

    while True:
        menu()
        escolha = input("\nEscolha uma opção: ").strip()

        if escolha == "1":
            print("\n--- Download de PDFs ---")
            iniciar_download(token)

        elif escolha == "2":
            print("\n--- Envio de E-mails ---")
            executar_envio_por_regiao(token, modo_correcao=False, dry_run=False)

        elif escolha == "3":
            print("\n--- Envio de Correção ---")
            executar_envio_por_regiao(token, modo_correcao=True, dry_run=False)

        elif escolha == "4":
            print("\n--- Simulação (Dry Run) ---")
            executar_envio_por_regiao(token, modo_correcao=False, dry_run=True)

        elif escolha == "5":
            print("\nEncerrando programa...")
            break

        else:
            print("\nOpção inválida, tente novamente.")

    input("\nPressione ENTER para sair...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nErro crítico: {e}")

    input("\nPressione ENTER para sair...")
