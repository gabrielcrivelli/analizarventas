\# Ventas Workbench (Desktop + Web)



Este repo contiene:



\- `/desktop`: app de escritorio en Python/Tkinter para consolidar ventas y generar reportes ejecutable en Windows.

\- `/web`: app web SaaS en Streamlit para analizar Excels de ventas, con múltiples acciones predefinidas (KPIs, tablas, detección de duplicados, etc.).



\## Uso local (web)
cd web
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
streamlit run app_streamlit.py

text

## Uso local (desktop)

cd desktop
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
pyinstaller --onefile --windowed --name "ConsolidadorVentas" ventas_consolidator_gui.py

text
undefined




