
# Automação de Geração de Briefings com Google Workspace

Este projeto automatiza a criação de documentos de briefing no Google Docs com base em dados armazenados em uma planilha do Google Sheets, utilizando um template padronizado e salvando os arquivos em uma pasta do Google Drive.

## Funcionalidade

- Leitura de dados de projetos do Google Sheets
- Geração de documentos personalizados a partir de um template no Google Docs
- Armazenamento dos documentos em uma pasta específica no Google Drive
- Verificação para evitar duplicidade de documentos
- Interface com alertas e barra de progresso
- Abertura automática da pasta no navegador ao final do processo

## Requisitos

- Python 3.12+
- Conta Google com acesso ao Google Sheets, Docs e Drive
- APIs ativadas no Google Cloud Console:
  - Google Drive API
  - Google Docs API
  - Google Sheets API

## Estrutura de Arquivos

```

├── main.py                 # Script principal da automação
├── client_secret.json      # Credenciais OAuth da conta Google
├── config.json             # IDs de planilha, template e pasta
├── token.pickle            # Token OAuth salvo após autenticação
├── requirements.txt        # Lista de dependências
└── run_automation.bat      # Script de execução automatizada no Windows
```

## Configuração

1. Crie um projeto no Google Cloud Console e ative as APIs necessárias.
2. Configure a tela de consentimento OAuth.
3. Gere um client ID do tipo "Desktop App" e salve como `client_secret.json`.
4. Crie um `config.json` com os seguintes campos:

```json
{
  "sheetsID": "ID_DA_PLANILHA",
  "templateID": "ID_DO_TEMPLATE",
  "outputFolderID": "ID_DA_PASTA"
}
```

## Execução

No Windows, use o script `.bat` incluído para ativar o ambiente virtual, instalar dependências e rodar o programa:

```bat
@echo off
call venv\Scripts\activate
pip install -r requirements.txt
python main.py
pause
```

Na primeira execução, será aberto o navegador para autenticação OAuth com sua conta Google.

## Instalação Manual

```bash
python -m venv venv
call venv\Scripts\activate
pip install -r requirements.txt
python main.py
```


