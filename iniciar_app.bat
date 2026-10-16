@echo off
cd /d "%~dp0"
echo Iniciando o servidor do aplicativo...
echo Abra o navegador manualmente e acesse http://localhost:8501
echo Mantenha esta janela aberta enquanto estiver usando o aplicativo.
echo.
py -m streamlit run app.py --server.headless=true --server.port=8501
echo.
echo Para parar o aplicativo, pressione Ctrl+C nesta janela.
pause
echo O aplicativo foi encerrado. Se apareceu uma mensagem de erro acima, tire uma foto dela.
