"""
Envio de Notas Fiscais e Boletos por e-mail via Microsoft Graph API.

Modos disponíveis:
    modo_correcao=False  →  Envio normal da competência
    modo_correcao=True   →  Reenvio com aviso de correção (substitui codenvioerrado.py)

Parâmetros de segurança:
    dry_run=True  →  Simula todo o processo sem enviar nenhum e-mail.
                     Use sempre antes do envio real para conferir.
"""
import base64
import requests
import pandas as pd
from pathlib import Path
from datetime import datetime

from config import MAILBOX_REMETENTE, EXCEL_PATH, THUMB_PATH, DOWNLOADS_DIR, RELATORIOS_DIR
from document_matcher import encontrar_pdfs_do_cliente

GRAPH_SEND_URL = f"https://graph.microsoft.com/v1.0/users/{MAILBOX_REMETENTE}/sendMail"
REGIOES_VALIDAS = {"RJ", "SP", "SUL", "NE"}


# ─── Corpo do e-mail ──────────────────────────────────────────────────────────

def _corpo_normal(content_id: str) -> str:
    return f"""
    <html>
        <body style="margin:0;padding:0;">
            <div style="text-align:center;">
                <img src="cid:{content_id}"
                     style="max-width:100%;height:auto;display:block;margin:0 auto;" />
            </div>
        </body>
    </html>
    """

def _corpo_correcao(competencia: str, content_id: str) -> str:
    aviso = f"""
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#111;
                line-height:1.45;padding:14px 16px;max-width:720px;margin:0 auto 12px auto;">
        <p style="margin:0 0 10px 0;">Prezados,</p>
        <div style="background:#fff4e5;border:1px solid #ffd59e;border-radius:8px;padding:12px 14px;margin:0 0 12px 0;">
            <p style="margin:0;">
                Identificamos que o e-mail enviado anteriormente continha a
                <b>competência incorreta</b>. Por gentileza,
                <b>desconsiderem o envio anterior</b>.
            </p>
            <p style="margin:8px 0 0 0;">
                Seguem, neste envio, os documentos <b>corrigidos</b> referentes
                à competência: <b>{competencia}</b>.
            </p>
        </div>
        <p style="margin:0;">Atenciosamente,<br><b>Equipe Financeira</b></p>
    </div>
    """
    return f"""
    <html>
        <body style="margin:0;padding:0;background:#ffffff;">
            {aviso}
            <div style="text-align:center;">
                <img src="cid:{content_id}"
                     style="max-width:100%;height:auto;display:block;margin:0 auto;" />
            </div>
        </body>
    </html>
    """

def _corpo_sem_thumb(modo_correcao: bool, competencia: str) -> str:
    if modo_correcao:
        return f"<p>Reenvio corrigido referente à competência <b>{competencia}</b>. Desconsidere o e-mail anterior.</p>"
    return "<p>Segue os documentos em anexo.</p>"


# ─── Envio individual ─────────────────────────────────────────────────────────

def _enviar_email(
    token: str,
    assunto: str,
    corpo_html: str,
    to_list: list[dict],
    anexos: list[dict],
) -> int:
    """Envia o e-mail via Graph API. Retorna o status HTTP."""
    payload = {
        "message": {
            "subject": assunto,
            "body": {"contentType": "HTML", "content": corpo_html},
            "toRecipients": to_list,
            "attachments": anexos,
        },
        "saveToSentItems": True,
    }
    r = requests.post(
        GRAPH_SEND_URL,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload,
    )
    return r.status_code


# ─── Fluxo principal ──────────────────────────────────────────────────────────

def executar_envio_por_regiao(token: str, modo_correcao: bool = False, dry_run: bool = False):
    """
    Lê a planilha, encontra os PDFs de cada cliente e envia os e-mails.

    Args:
        token:         Token OAuth2 gerado pelo auth.py.
        modo_correcao: True para envio de correção com aviso no corpo.
        dry_run:       True para simular sem enviar nenhum e-mail.
    """
    if dry_run:
        print("\n[SIMULAÇÃO] Nenhum e-mail será enviado. Apenas conferência.\n")

    competencia = input("Digite a competência (ex: 12/2025): ").strip()

    CONTENT_ID_IMG = "header-img"
    thumb_b64 = None
    if THUMB_PATH.exists():
        thumb_b64 = base64.b64encode(THUMB_PATH.read_bytes()).decode("utf-8")

    relatorio: list[list] = []

    while True:
        print(f"\nRegiões disponíveis: {', '.join(sorted(REGIOES_VALIDAS))}")
        regiao = input("Digite a região que deseja processar: ").strip().upper()

        if regiao not in REGIOES_VALIDAS:
            print("Região inválida.")
            continue

        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=regiao, dtype=str)
        except Exception as e:
            print(f"Erro ao ler a planilha: {e}")
            continue

        df["Posto"] = df["Posto"].astype(str).str.strip()
        pasta_regiao = DOWNLOADS_DIR / regiao

        prefixo_assunto = "CORREÇÃO - " if modo_correcao else ""

        for _, row in df.iterrows():
            nome_posto   = row["Posto"]
            nfse_num     = row.get("NFSe", "")
            destino_raw  = row.get("Email do cliente", "")

            # ── Valida PDFs ──
            arquivos = encontrar_pdfs_do_cliente(nome_posto, pasta_regiao)
            if not arquivos:
                print(f"  [SEM PDF] {nome_posto}")
                relatorio.append([regiao, _agora(), nome_posto, nfse_num, "Sem PDF encontrado", ""])
                continue

            # ── Valida destinatários ──
            to_list = _parsear_emails(destino_raw)
            if not to_list:
                print(f"  [SEM EMAIL] {nome_posto}")
                relatorio.append([regiao, _agora(), nome_posto, nfse_num, "Sem e-mail cadastrado", ""])
                continue

            destinatarios_str = "; ".join(m["emailAddress"]["address"] for m in to_list)
            assunto = (
                f"{prefixo_assunto}NOTA FISCAL DE CONTRATO - "
                f"NF {nfse_num} {nome_posto} - {competencia}"
            )

            # ── Monta corpo e anexos ──
            if thumb_b64 and not modo_correcao:
                corpo = _corpo_normal(CONTENT_ID_IMG)
            elif thumb_b64 and modo_correcao:
                corpo = _corpo_correcao(competencia, CONTENT_ID_IMG)
            else:
                corpo = _corpo_sem_thumb(modo_correcao, competencia)

            anexos = _montar_anexos(thumb_b64, CONTENT_ID_IMG, arquivos)

            # ── Envia ou simula ──
            print(f"  {'[SIM]' if dry_run else '[ENVIANDO]'} {nome_posto} → {destinatarios_str}")

            if dry_run:
                status = "SIMULADO"
            else:
                codigo = _enviar_email(token, assunto, corpo, to_list, anexos)
                status = "Enviado" if codigo == 202 else f"Erro {codigo}"

            relatorio.append([regiao, _agora(), nome_posto, nfse_num, status, destinatarios_str])

        continuar = input("\nDeseja processar outra região? (SIM/NÃO): ").strip().upper()
        if continuar in ("NÃO", "NAO"):
            print("\nEncerrando envio.")
            break

    _salvar_e_enviar_relatorio(token, relatorio, competencia, dry_run)


# ─── Auxiliares ───────────────────────────────────────────────────────────────

def _agora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _parsear_emails(destino_raw: str) -> list[dict]:
    if not isinstance(destino_raw, str) or not destino_raw.strip():
        return []
    return [
        {"emailAddress": {"address": mail.strip()}}
        for mail in destino_raw.replace(",", ";").split(";")
        if mail.strip()
    ]


def _montar_anexos(thumb_b64: str | None, content_id: str, pdfs: list[Path]) -> list[dict]:
    anexos = []
    if thumb_b64:
        anexos.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": THUMB_PATH.name,
            "contentId": content_id,
            "isInline": True,
            "contentType": "image/jpeg",
            "contentBytes": thumb_b64,
        })
    for pdf in pdfs:
        anexos.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": pdf.name,
            "contentBytes": base64.b64encode(pdf.read_bytes()).decode("utf-8"),
        })
    return anexos


def _salvar_e_enviar_relatorio(
    token: str,
    relatorio: list[list],
    competencia: str,
    dry_run: bool,
):
    if not relatorio:
        return

    df_rel = pd.DataFrame(
        relatorio,
        columns=["Região", "Data Envio", "Posto", "NFSe", "Status", "E-mails dos destinatários"],
    )

    prefixo = "SIMULACAO_" if dry_run else ""
    nome_arquivo = f"{prefixo}Relatorio_Envios_{competencia.replace('/', '-')}.xlsx"
    caminho_rel  = RELATORIOS_DIR / nome_arquivo
    df_rel.to_excel(caminho_rel, index=False)
    print(f"\nRelatório salvo em: {caminho_rel}")

    if dry_run:
        print("[SIMULAÇÃO] Relatório não será enviado por e-mail.")
        return

    email_dest = input("Digite o e-mail que receberá o relatório: ").strip()
    relatorio_b64 = base64.b64encode(caminho_rel.read_bytes()).decode("utf-8")

    payload = {
        "message": {
            "subject": f"Relatório Automatizado - NFs e Boletos ({competencia})",
            "body": {"contentType": "Text", "content": "Segue o relatório de envios."},
            "toRecipients": [{"emailAddress": {"address": email_dest}}],
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": nome_arquivo,
                "contentBytes": relatorio_b64,
            }],
        },
        "saveToSentItems": True,
    }

    r = requests.post(
        GRAPH_SEND_URL,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload,
    )

    if r.status_code == 202:
        print("Relatório enviado com sucesso.")
    else:
        print(f"Falha ao enviar relatório: {r.status_code} — {r.text}")


if __name__ == "__main__":
    from auth import get_graph_token
    executar_envio_por_regiao(get_graph_token())
