@echo off
REM Script simplificado que funciona igual ao comando manual
setlocal

title Empacotando Extrator de Arquivos

echo ***************************************
echo *   CONSTRUINDO EXECUTAVEL SIMPLES    *
echo ***************************************

REM Verifica se o arquivo principal existe
if not exist "main.py" (
    echo ERRO: Arquivo main.py nao encontrado!
    pause
    exit /b 1
)

REM Verifica se o icone existe
if not exist "icon.ico" (
    echo AVISO: Arquivo icon.ico nao encontrado, continuando sem icone
    set ICON_CMD=
) else (
    set ICON_CMD=--icon=icon.ico
)

REM Ativa o ambiente virtual (se existir)
if exist "env\Scripts\activate.bat" (
    call env\Scripts\activate.bat
)

echo.
echo Executando PyInstaller...
echo Comando: pyinstaller --onefile --windowed %ICON_CMD% main.py

pyinstaller --onefile --windowed %ICON_CMD% main.py

if errorlevel 1 (
    echo.
    echo **************************************
    echo * ERRO: Falha ao construir o executavel
    echo **************************************
    pause
    exit /b 1
)

echo.
echo **************************************
echo * BUILD CONCLUIDO COM SUCESSO!      *
echo * Executavel: dist\main.exe         *
echo **************************************

echo.
pause