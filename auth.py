"""
Autenticação OAuth2 App-Only — Microsoft Graph API.
Token gerado uma única vez pelo main.py e passado para os módulos.
"""
import requests
from config import TENANT_ID, CLIENT_ID, CLIENT_SECRET

_TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"


def get_graph_token() -> str:
    """Gera token OAuth2 App-Only para autenticação na Microsoft Graph API."""
    r = requests.post(_TOKEN_URL, data={
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    })

    if r.status_code == 200:
        print("Token Microsoft Graph gerado com sucesso.")
        return r.json()["access_token"]

    raise Exception(
        f"Falha ao autenticar no Microsoft Graph: {r.status_code} — {r.text}"
    )
