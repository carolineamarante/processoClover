@echo off
echo [INFO] Ativando ambiente virtual...
call venv\Scripts\activate

echo [INFO] Instalando dependências...
pip install --upgrade pip
pip install -r requirements.txt

echo [INFO] Executando a automação...
python main.py

pause
