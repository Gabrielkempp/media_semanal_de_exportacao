@echo off
rem Rotina semanal MDIC: baixa a divulgacao atual e grava na tabela do data lake.
rem Pode rodar mais de uma vez na semana: reexecutar nao duplica nada.
cd /d "%~dp0"
set PYTHONUTF8=1
python main.py %*
exit /b %ERRORLEVEL%
