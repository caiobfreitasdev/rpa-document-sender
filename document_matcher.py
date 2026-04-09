"""
Regra de negócio principal: match entre nome do cliente (planilha) e arquivos PDF.

Lógica:
    O nome do posto lido da planilha é procurado (case-insensitive) no nome
    de cada arquivo PDF presente na pasta da região correspondente.
    Todos os PDFs cujo nome contenha o nome do posto são retornados.
"""
from pathlib import Path


def encontrar_pdfs_do_cliente(nome_posto: str, pasta_regiao: Path) -> list[Path]:
    """
    Retorna a lista de PDFs que correspondem ao nome do posto/cliente.

    Args:
        nome_posto:    Nome do cliente conforme consta na planilha Excel.
        pasta_regiao:  Pasta local onde os PDFs da região foram salvos.

    Returns:
        Lista de Path dos PDFs encontrados. Lista vazia se nenhum corresponder.
    """
    if not pasta_regiao.exists():
        return []

    nome_normalizado = nome_posto.strip().lower()

    return [
        pdf
        for pdf in pasta_regiao.glob("*.pdf")
        if nome_normalizado in pdf.name.lower()
    ]
