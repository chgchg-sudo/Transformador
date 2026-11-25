Extractor ANECA / Sexenios
Aplicación desarrollada en Streamlit para extraer automáticamente las aportaciones científicas de los documentos de aportaciones utilizados en los procesos de evaluación ANECA y generar un fichero Excel estructurado para su análisis y revisión.

Funcionalidades
Carga de documentos Word (.docx).
Extracción automática de aportaciones.
Identificación de:
Tipo de aportación.
Título.
Revista.
ID Persistente.
Referencia APA.
Objetivos perseguidos.
Resultados más relevantes.
Liderazgo.
Indicadores de impacto científico.
Indicadores de impacto social.
Indicadores de ciencia abierta.
Exportación a Excel (.xlsx).
Vista previa de los datos extraídos.
Requisitos
Python 3.10 o superior
Instalación
Clonar el repositorio:

git clone https://github.com/TU_USUARIO/extractor-aneca.git
cd extractor-aneca
Instalar dependencias:

pip install -r requirements.txt
Ejecución
streamlit run app.py
La aplicación estará disponible en:

http://localhost:8501
Despliegue en Streamlit Cloud
Subir el código a GitHub.
Acceder a https://share.streamlit.io
Iniciar sesión con GitHub.
Seleccionar el repositorio.
Indicar app.py como archivo principal.
Pulsar Deploy.
Ficheros del proyecto
.
├── app.py
├── requirements.txt
└── README.md
Dependencias
streamlit
pandas
openpyxl
python-docx
Licencia
Uso académico y de investigación.
