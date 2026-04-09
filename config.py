"""
Configurações centralizadas do projeto.
Todas as variáveis sensíveis são lidas do arquivo .env — nunca hardcoded aqui.
"""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

def _require(var: str) -> str:
    """Lê variável obrigatória do .env. Falha com mensagem clara se ausente."""
    value = os.getenv(var)
    if not value:
        raise EnvironmentError(
            f"Variável de ambiente obrigatória não definida: '{var}'\n"
            f"Copie o arquivo .env.example para .env e preencha os valores."
        )
    return value

# ─── Azure AD / Microsoft Graph ───────────────────────────────────────────────
TENANT_ID     = _require("TENANT_ID")
CLIENT_ID     = _require("CLIENT_ID")
CLIENT_SECRET = _require("CLIENT_SECRET")

# ─── SharePoint ───────────────────────────────────────────────────────────────
SHAREPOINT_SITE        = _require("SHAREPOINT_SITE")
SHAREPOINT_BASE_FOLDER = _require("SHAREPOINT_BASE_FOLDER")

# ─── E-mail ───────────────────────────────────────────────────────────────────
MAILBOX_REMETENTE = _require("MAILBOX_REMETENTE")

# ─── Caminhos locais ──────────────────────────────────────────────────────────
BASE_DIR       = Path.home() / "RPA_SINERGAS"
EXCEL_PATH     = BASE_DIR / os.getenv("EXCEL_FILENAME", "Base/EMAILS_CLIENTES.xlsx")
THUMB_PATH     = BASE_DIR / os.getenv("THUMBNAIL_FILENAME", "Base/thumbnail.jpg")
DOWNLOADS_DIR  = BASE_DIR / "downloads"
RELATORIOS_DIR = BASE_DIR / "relatorios"

# Garante que as pastas de trabalho existam
DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
RELATORIOS_DIR.mkdir(parents=True, exist_ok=True)
