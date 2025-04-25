# Informe CEV v2 (PDF scraper)

Esta es una aplicación web simple construida con Streamlit que permite a los usuarios cargar archivos PDF correspondientes al formato "Informe de Calificación Energética de Vivienda v2" (CEV v2), extraer datos específicos de diferentes páginas y visualizarlos en un formato tabular. Además, permite descargar toda la información extraída en un archivo Excel.

## Descripción

La aplicación utiliza la biblioteca `PyMuPDF (fitz)` para leer y extraer texto de coordenadas específicas dentro de los archivos PDF. Los datos extraídos se organizan utilizando `pandas` DataFrames y se presentan en una interfaz web interactiva creada con `Streamlit`. La interfaz organiza los datos en pestañas correspondientes a las páginas del informe original. Finalmente, ofrece la opción de descargar un resumen completo en formato `.xlsx`.

## Características

* **Carga de Archivos PDF:** Permite subir un archivo PDF (formato Informe CEV v2) a la vez.
* **Validación de Archivos:** Realiza una verificación básica para asegurar que el PDF cargado corresponde al formato esperado.
* **Extracción de Datos:** Extrae datos estructurados de las páginas 1, 2, 3 (Consumos y Envolvente), 4 y 7 del informe. Las páginas 5 y 6 se incluyen como marcadores de posición.
* **Visualización Web:** Muestra los datos extraídos en tablas interactivas dentro de pestañas organizadas por página.
* **Transposición:** Presenta los datos de ciertas páginas (P1, P2, P3 Consumos, P5, P6, P7) transpuestos para facilitar la lectura.
* **Renombrado:** Utiliza nombres descriptivos en español para las filas/columnas de las tablas mostradas.
* **Descarga en Excel:** Permite descargar toda la información extraída (en formato original, no transpuesto y con nombres descriptivos) en un único archivo Excel (`.xlsx`) con múltiples hojas.

## Requisitos

Para ejecutar esta aplicación, necesitarás tener Python 3 instalado, junto con las siguientes bibliotecas:

* streamlit
* pandas
* PyMuPDF (fitz)
* openpyxl (para la generación de Excel)

## Instalación

1.  Clona o descarga este repositorio.
2.  Navega al directorio del proyecto en tu terminal.
3.  Instala las dependencias usando pip:
    ```bash
    pip install streamlit pandas PyMuPDF openpyxl
    ```

## Uso

1.  Asegúrate de tener los archivos `app.py` e `scraping_functions.py` en el mismo directorio.
2.  Abre tu terminal, navega al directorio del proyecto.
3.  Ejecuta la aplicación Streamlit con el siguiente comando:
    ```bash
    streamlit run app.py
    ```
4.  Se abrirá una pestaña en tu navegador web con la aplicación.
5.  Usa la pestaña "Subir Archivo PDF" para seleccionar un archivo PDF de informe CEV v2. La aplicación procesará automáticamente el archivo.
6.  Una vez procesado, navega por las pestañas "Página 1" a "Página 7" para ver los datos extraídos.
7.  Si el procesamiento fue exitoso, aparecerá un botón "Descargar Informe CEV (Excel)" debajo del título principal para descargar el archivo `.xlsx` con todos los datos.

## Estructura del Proyecto

* `app.py`: Contiene el código de la aplicación web Streamlit, la interfaz de usuario, la lógica de carga/procesamiento y la visualización de datos.
* `scraping_functions.py`: Contiene las funciones responsables de extraer los datos de las diferentes páginas del PDF utilizando coordenadas específicas.
* `README.md`: Este archivo.