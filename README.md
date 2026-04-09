# RPA Document Sender

Automação desenvolvida em Python para envio de Notas Fiscais e Boletos por e-mail, com integração ao **SharePoint** via **Microsoft Graph API**.

O sistema lê uma planilha Excel com dados dos clientes, localiza os documentos correspondentes (por nome do cliente) e realiza o envio automatizado dos e-mails com os anexos corretos — eliminando o processo manual de separação e envio documento a documento.

---

## Funcionalidades

- Download automático de PDFs direto do SharePoint
- Match automático de documentos pelo nome do cliente
- Envio de e-mails com anexos via Microsoft Graph API
- Modo de **correção de competência** (reenvio com aviso automático)
- Modo **simulação (Dry Run)** — confere tudo sem disparar nenhum e-mail
- Geração de relatório Excel com status de cada envio
- Envio automático do relatório por e-mail ao final do processo

---

## Tecnologias utilizadas

- Python 3.11+
- [Microsoft Graph API](https://learn.microsoft.com/pt-br/graph/overview)
- OAuth2 App-Only (Client Credentials Flow)
- `pandas` — leitura da planilha de clientes
- `openpyxl` — geração do relatório de envios
- `requests` — chamadas HTTP à API
- `python-dotenv` — gerenciamento seguro de credenciais

---

## Estrutura do projeto

```
rpa-document-sender/
├── main.py                          # Ponto de entrada — menu interativo
├── auth.py                          # Autenticação OAuth2 Microsoft Graph
├── config.py                        # Configurações centralizadas (lê o .env)
├── sharepoint_dl.py                 # Download de PDFs do SharePoint
├── email_sender.py                  # Envio de e-mails e geração de relatório
├── document_matcher.py              # Regra de negócio: match cliente ↔ PDF
├── requirements.txt
├── .env.example                     # Modelo de configuração (sem dados reais)
├── .gitignore
└── data/
    └── example/
        └── CLIENTES_EXAMPLE.xlsx    # Planilha de exemplo com dados fictícios
```

> As pastas `Base/`, `downloads/` e `relatorios/` são criadas automaticamente na primeira execução e estão protegidas pelo `.gitignore`.

---

## Como configurar

### 1. Clone o repositório

```bash
git clone https://github.com/caiobfreitasdev/rpa-document-sender.git
cd rpa-document-sender
```

### 2. Crie o ambiente virtual e instale as dependências

```bash
python -m venv venv
venv\Scripts\activate      # Windows
pip install -r requirements.txt
```

### 3. Configure as variáveis de ambiente

Copie o arquivo de modelo e preencha com suas credenciais:

```bash
copy .env.example .env
```

Abra o `.env` e preencha:

```env
TENANT_ID=seu-tenant-id
CLIENT_ID=seu-client-id
CLIENT_SECRET=seu-client-secret
SHAREPOINT_SITE=suaempresa.sharepoint.com:/sites/NomeDaSite
SHAREPOINT_BASE_FOLDER=Caminho/Da/Pasta/No/SharePoint
MAILBOX_REMETENTE=financeiro@suaempresa.com.br
```

> As credenciais Azure AD são obtidas no [Portal Azure](https://portal.azure.com) em **App Registrations**.

### 4. Adicione sua planilha de clientes

Coloque o arquivo Excel na pasta `Base/` com a seguinte estrutura de colunas:

| Posto | NFSe | Email do cliente |
|-------|------|-----------------|
| Nome do cliente | Número da NF | email@cliente.com.br |

Cada aba da planilha deve corresponder a uma região: `RJ`, `SP`, `SUL`, `NE`.

Veja o modelo em [`data/example/CLIENTES_EXAMPLE.xlsx`](data/example/CLIENTES_EXAMPLE.xlsx).

---

## Como executar

```bash
python main.py
```

O menu interativo oferece as seguintes opções:

```
===== MENU DO ROBÔ =====
1) Baixar PDFs do SharePoint
2) Enviar Boletos/NFs por E-mail
3) Enviar Correção de Competência
4) Simulação de Envio (sem disparar e-mails)
5) Sair
```

> **Recomendação:** sempre utilize a opção **4 - Simulação** antes do envio real para conferir os PDFs e destinatários de cada cliente.

---

## Como o match de documentos funciona

A regra principal do sistema:

1. Lê o **nome do cliente** na planilha Excel
2. Procura na pasta da região (`downloads/RJ`, `downloads/SP`, etc.) arquivos PDF cujo nome **contenha** o nome do cliente
3. Todos os PDFs encontrados são anexados ao e-mail daquele cliente

Essa lógica está isolada em [`document_matcher.py`](document_matcher.py) para facilitar manutenção e testes.

---

## Autor

Desenvolvido por **Caio Freitas**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Caio%20Freitas-0077B5?style=flat&logo=linkedin)](https://www.linkedin.com/in/caioffreitas/)
