@echo off
setlocal
cd /d "%~dp0"
echo Instalando dependencias do projeto...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Nao foi possivel instalar as dependencias.
    pause
    exit /b 1
)
echo.
echo Executando o pipeline do PM3...
python projMens3.py
if errorlevel 1 (
    echo.
    echo A execucao encontrou um erro.
    pause
    exit /b 1
)
echo.
echo Projeto executado com sucesso.
echo Relatorio: documentacao\relatorio_final.pdf
echo Graficos: documentacao\graficos\
pause
