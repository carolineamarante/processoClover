import json, gspread, webbrowser, os.path, pickle, tkinter as tk
from tqdm import tqdm
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.exceptions import RefreshError
from tkinter import messagebox


# Escopos necessários
scopes = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/spreadsheets.readonly'
]

# Configuração inicial
with open("config.json", "r") as f:
    config = json.load(f)

sheetID = config["sheetsID"]
templateID = config["templateID"]
outputFolderID = config["outputFolderID"]

# Autenticação
creds = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError:
            creds = None
    if not creds:
        flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', scopes)
        creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

# Inicializa serviços Google
client = gspread.authorize(creds)
driveService = build('drive', 'v3', credentials=creds)
docsService = build('docs', 'v1', credentials=creds)

# Leitura da planilha
sheet = client.open_by_key(sheetID).sheet1
data = sheet.get_all_records()

root = tk.Tk()
root.withdraw()

# Filtrar apenas os projetos que ainda não têm documentos criados
projectsToCreate = []
for project in data:
    documentName = f"Briefing - {project['Nome do Projeto']}"
    query = f"name = '{documentName}' and '{outputFolderID}' in parents and trashed = false"
    results = driveService.files().list(q=query).execute()
    if not results["files"]:
        projectsToCreate.append(project)
    else:
        print(f"Documento já existe: {documentName}")


# Exibe a barra de progresso
for project in tqdm(projectsToCreate, desc="Gerando documentos"):
    projectName = project["Nome do Projeto"]
    documentName = f"Briefing - {projectName}"

    # Copia o template
    newDocument = driveService.files().copy(
        fileId=templateID,
        body={'name': documentName, 'parents': [outputFolderID]}
    ).execute()

    if not projectsToCreate:
        messagebox.showinfo("Informação", "Nenhum novo projeto para processar.\nTodos os documentos já foram gerados.")
        exit()

    documentUniqueID = newDocument['id']

    # Substituição de conteúdo
    requests = [
        {"replaceAllText": {
            "containsText": {"text": "{{NOME_DO_PROJETO}}", "matchCase": True},
            "replaceText": project.get("Nome do Projeto", "")
        }},
        {"replaceAllText": {
            "containsText": {"text": "{{RESPONSAVEL}}", "matchCase": True},
            "replaceText": project.get("Responsável", "")
        }},
        {"replaceAllText": {
            "containsText": {"text": "{{PRAZO_FINAL}}", "matchCase": True},
            "replaceText": project.get("Prazo Final", "")
        }},
        {"replaceAllText": {
            "containsText": {"text": "{{DESCRICAO_DO_PROJETO}}", "matchCase": True},
            "replaceText": project.get("Descrição do Projeto", "")
        }},
        {"replaceAllText": {
            "containsText": {"text": "{{OBSERVACOES_ADICIONAIS}}", "matchCase": True},
            "replaceText": project.get("Observações Adicionais", "")
        }},
        {"replaceAllText": {
            "containsText": {"text": "{{STATUS_DO_PROJETO}}", "matchCase": True},
            "replaceText": project.get("Status do Projeto", "")
        }},
        {"replaceAllText": {
            "containsText": {"text": "{{DIAS_ATE_FINALIZACAO}}", "matchCase": True},
            "replaceText": str(project.get("Dias até Finalização", ""))
        }}
    ]

    docsService.documents().batchUpdate(
        documentId=documentUniqueID,
        body={"requests": requests}
    ).execute()

    print(f"Documento criado com sucesso: {documentName}")



# Após finalização de todos os documentos, abre a pasta no navegador
folderURL = f"https://drive.google.com/drive/folders/{outputFolderID}"
webbrowser.open(folderURL)
print("Todos os documentos foram processados. A pasta foi aberta no navegador.")

messagebox.showinfo("Resumo",
                    f"Foram criados {len(projectsToCreate)} documentos de um total de {len(data)} projetos.\n"
                    f"A pasta do Google Drive foi aberta no navegador.")