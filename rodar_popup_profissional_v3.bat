@echo off
setlocal
cd /d "%~dp0"
title Assistente Profissional de Orcamento

pythonw "%~dp0pricing_popup_professional_v3.py"
if errorlevel 1 (
    echo Erro ao iniciar a interface profissional.
    echo Verifique se o Python esta instalado e se o arquivo pricing_popup_professional_v3.py esta na mesma pasta.
    pause
)

endlocal
