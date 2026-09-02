@echo off
REM Atualiza o painel Torre de Controle - Consumo de Gelo
REM Coloque este .bat na mesma pasta do atualizar_painel_gelo.py e do template.html
cd /d "%~dp0"
python atualizar_painel_gelo.py >> log_atualizacao.txt 2>&1
