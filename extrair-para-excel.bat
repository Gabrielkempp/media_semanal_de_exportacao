@echo off
rem Rotina semanal MDIC: baixa a divulgacao atual, arquiva os arquivos do site e atualiza o Excel.
rem Pode rodar mais de uma vez na semana: reexecutar nao duplica nada.
cd /d "%~dp0"
set PYTHONUTF8=1
python extrair_para_excel.py %*
exit /b %ERRORLEVEL%
