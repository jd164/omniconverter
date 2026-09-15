@echo off
title OmniConverter - Painel de Conversao de Chamadas
echo ======================================================================
echo           OmniConverter - Painel Web & Conversor de Chamadas
echo ======================================================================
echo.
echo A iniciar o servidor e abrir o painel web no seu navegador...
echo.

python main.py --web

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERRO] Ocorreu um problema ao iniciar o servidor. Verifique se o Python esta instalado.
    echo Pressione qualquer tecla para sair...
    pause >nul
)
