@echo off
echo ==================================================
echo 🚀 INSTALANDO DEPENDENCIAS Y LANZANDO NOMINA APP
echo ==================================================

:: 1. Instalar dependencias
pip install -r requirements.txt

:: 2. Lanzar la aplicación
echo Iniciando aplicacion...
streamlit run app.py

:: 3. LIMPIEZA AUTOMÁTICA AL CERRAR
echo ==================================================
echo 🧹 LIMPIANDO ARCHIVOS TEMPORALES...
echo ==================================================

:: Borrar archivos de un solo uso (ejemplo: los archivos PDF que se procesaron)
if exist "*.pdf" del /q *.pdf

:: Si deseas borrar también la base de datos (SOLO SI ES DE UN SOLO USO)
:: del /q nomina_db.xlsx

echo Proceso finalizado.
pause