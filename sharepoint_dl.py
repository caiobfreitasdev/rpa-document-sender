"""
Download de PDFs do SharePoint via Microsoft Graph API.
"""
import requests
from pathlib import Path
from config import SHAREPOINT_SITE, SHAREPOINT_BASE_FOLDER, DOWNLOADS_DIR

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _get_site_id(token: str) -> str:
    """Resolve o ID do site SharePoint via Microsoft Graph."""
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(f"{GRAPH_BASE}/sites/{SHAREPOINT_SITE}", headers=headers)
    if r.status_code == 200:
        return r.json()["id"]
    raise Exception(f"Erro ao resolver site do SharePoint: {r.status_code} — {r.text}")


def _baixar_pasta(folder_path: str, local_dir: Path, headers: dict, site_id: str) -> int:
    """
    Lista e baixa PDFs recursivamente de uma pasta do SharePoint.
    Cria os diretórios locais automaticamente antes de salvar.
    Retorna a quantidade de PDFs baixados.
    """
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{folder_path}:/children"
    r = requests.get(url, headers=headers)

    if r.status_code != 200:
        print(f"Falha ao listar pasta '{folder_path}': {r.status_code} — {r.json()}")
        return 0

    total = 0

    for item in r.json().get("value", []):
        nome = item["name"]

        if item.get("folder"):
            total += _baixar_pasta(f"{folder_path}/{nome}", local_dir / nome, headers, site_id)

        elif item.get("file") and nome.lower().endswith(".pdf"):
            local_dir.mkdir(parents=True, exist_ok=True)
            destino = local_dir / nome

            pdf_bin = requests.get(item["@microsoft.graph.downloadUrl"])
            destino.write_bytes(pdf_bin.content)

            print(f"  PDF baixado: {destino}")
            total += 1

    return total


def iniciar_download(token: str):
    """
    Fluxo interativo de download:
    1. Pergunta o mês uma vez.
    2. Loop para selecionar regiões até o usuário sair.
    """
    site_id = _get_site_id(token)
    headers = {"Authorization": f"Bearer {token}"}

    mes_nome = input("\nDigite o mês (ex: 12_DEZ): ").strip()

    while True:
        regiao = input("\nQual região deseja acessar? (SP, RJ, NE, SUL): ").strip().upper()

        if regiao not in ["SP", "RJ", "NE", "SUL"]:
            print("Região inválida. Use: SP, RJ, NE ou SUL.")
            continue

        pasta_remota = f"{SHAREPOINT_BASE_FOLDER}/{mes_nome}/{regiao}"
        pasta_local  = DOWNLOADS_DIR / regiao

        print(f"\n  Pasta remota : {pasta_remota}")
        print(f"  Salvando em  : {pasta_local}")

        total = _baixar_pasta(pasta_remota, pasta_local, headers, site_id)
        print(f"\n  Concluído: {total} PDF(s) baixado(s) da região {regiao}.")

        continuar = input("\nDeseja baixar outra região? (S/N): ").strip().upper()
        if continuar == "N":
            print("\nEncerrando download.")
            break
        elif continuar != "S":
            print("\nResposta inválida. Encerrando.")
            break


if __name__ == "__main__":
    from auth import get_graph_token
    iniciar_download(get_graph_token())
