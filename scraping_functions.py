
from bisect import bisect_left
from typing import Dict, Tuple, Any, List, Union
from functools import lru_cache
import pandas as pd
import fitz  # PyMuPDF
import logging

@lru_cache(maxsize=128)
def normalize_coordinates(
    x: float, 
    y: float, 
    report_width: float, 
    report_height: float, 
    page_width: float, 
    page_height: float
) -> Tuple[float, float]:
    """Normalize coordinates with caching for repeated calculations."""
    rx = (x / report_width) * page_width
    ry = (y / report_height) * page_height
    return rx, ry

def extract_text_from_area(page: fitz.Page, area: Tuple[float, float, float, float]) -> str:
    """
    Extract text from a specific area of a PDF page.

    Args:
        page (fitz.Page): Page object from which to extract text.
        area (tuple): Tuple containing (x1, y1, x2, y2) coordinates of the area to extract text from.

    Returns:
        str: Text extracted from the specified area.
    
    Raises:
        ValueError: If invalid coordinates are provided
        TypeError: If invalid argument types are provided
    """
    if not isinstance(page, fitz.Page):
        raise TypeError("Page argument must be a fitz.Page object")
    
    if not isinstance(area, tuple) or len(area) != 4:
        raise TypeError("Area must be a tuple of 4 coordinates")

    # Constants (defined as class attributes or in a config file in a real application)
    REPORT_WIDTH = 215.9  # mm
    REPORT_HEIGHT = 330.0  # mm
    
    try:
        # Get page dimensions once
        width = page.rect.width
        height = page.rect.height
        
        # Validate coordinates
        x1, y1, x2, y2 = area
        if x1 >= x2 or y1 >= y2:
            raise ValueError("Invalid coordinates: ensure x1 < x2 and y1 < y2")
        
        # Normalize coordinates using cached function
        rx1, ry1 = normalize_coordinates(x1, y1, REPORT_WIDTH, REPORT_HEIGHT, width, height)
        rx2, ry2 = normalize_coordinates(x2, y2, REPORT_WIDTH, REPORT_HEIGHT, width, height)
        
        # Create rectangle and extract text in one operation
        rect = fitz.Rect(rx1, ry1, rx2, ry2)
        return page.get_textbox(rect).strip()
    
    except Exception as e:
        # Log the error in a production environment
        print(f"Error extracting text: {str(e)}")
        logging.warning(f"Error extracting text: {str(e)}")
        return ""

# ------------------------------------------------------------------------------------------------------------
#  Pagina 1    
# ------------------------------------------------------------------------------------------------------------    
def _from_procentaje_ahorro_to_letra(porcentaje_ahorro: float) -> str:
    """
    Convert a savings percentage to a corresponding letter grade using bisect.
    """
    boundaries = [-0.35, -0.1, 0.2, 0.4, 0.55, 0.7, 0.85, 100]
    grades = ['G', 'F', 'E', 'D', 'C', 'B', 'A', 'A+']
    
    idx = bisect_left(boundaries, porcentaje_ahorro)
    return grades[idx] if 0 <= idx < len(grades) else None

def get_informe_cev_v2_pagina1_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 1 of an informe_CEV_v2 PDF report and return it as a dictionary.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        Dict[str, Any]: A dictionary containing the extracted data fields from page 1.
                       Returns an empty dictionary if any error occurs during extraction.
    """
    try:
        # Validate input
        if not isinstance(pdf_report, fitz.Document):
            raise ValueError("Invalid input: pdf_report must be a fitz.Document object.")
        if 0 >= pdf_report.page_count:
            raise ValueError("Invalid page number: Page 1 does not exist in the document.")

        page = pdf_report[0]  # Get page 1 (index 0)

        # Define coordinates as constants for page 1
        COORDINATES: Dict[str, Tuple[float, float, float, float]] = {
            'tipo_evaluacion': (8.3, 10.3, 165.6, 18.8),
            'codigo_evaluacion': (73.1, 20.0, 95.6, 25.1),  
            'region': (28.0, 26.6, 80.0, 31.8),             
            'comuna': (29.2, 33.0, 80.0, 38.2)	,  
            'direccion': (31.3, 39.1, 155.3, 44.3),
            'rol_vivienda_proyecto': (57.4, 45.6, 74.8, 50.8), 
            'tipo_vivienda':(45.9, 51.7, 155.3, 56.9),     
            'superficie_interior_util_m2': (54.2, 58.3, 66.0, 63.5),
            'porcentaje_ahorro': (5.6, 78.6, 165.8, 191.3),            
            'demanda_calefaccion_kwh_m2_ano': (15.6, 220.0, 73.0, 230.0),
            'demanda_enfriamiento_kwh_m2_ano': (90.0, 220.0, 151.5, 230.0),
            'demanda_total_kwh_m2_ano': (167.0, 225.0, 209.0, 245.0),
            'emitida_el': (34.5, 247.5, 57.0, 252.8) 
        }

        # Extract all fields at once
        fields: Dict[str, str] = {
            key: extract_text_from_area(page, coords) if coords else None
            for key, coords in COORDINATES.items()
        }

        porcentaje_ahorro = next((int(item) for item in fields.get('porcentaje_ahorro', '').splitlines() if item.replace('-', '').isdigit()), None)

        # Create a dictionary to return
        result: Dict[str, Any] = {
            'tipo_evaluacion': fields.get('tipo_evaluacion', '').strip(),
            'codigo_evaluacion': fields.get('codigo_evaluacion', '').strip(),
            'region': fields.get('region', '').strip(),
            'comuna': fields.get('comuna', '').strip(),
            'direccion': fields.get('direccion', '').strip(),
            'rol_vivienda_proyecto': fields.get('rol_vivienda_proyecto', '').strip(),
            'tipo_vivienda': fields.get('tipo_vivienda', '').strip(),
            'superficie_interior_util_m2': float(fields.get('superficie_interior_util_m2', '').replace(',', '.').strip()),            
            'porcentaje_ahorro': porcentaje_ahorro,
            'letra_eficiencia_energetica_dem': _from_procentaje_ahorro_to_letra(porcentaje_ahorro/100),
            'demanda_calefaccion_kwh_m2_ano': float(fields.get('demanda_calefaccion_kwh_m2_ano', '').splitlines()[-1].replace(',', '.').strip()),
            'demanda_enfriamiento_kwh_m2_ano': float(fields.get('demanda_enfriamiento_kwh_m2_ano', '').splitlines()[-1].replace(',', '.').strip()),
            'demanda_total_kwh_m2_ano': float(fields.get('demanda_total_kwh_m2_ano', '').splitlines()[-1].replace(',', '.').strip()),
            'emitida_el': fields.get('emitida_el', '').splitlines()[-1].strip()
        
        }

        return result

    except (RuntimeError, IndexError, ValueError) as e:
        logging.error(f"Error in get_informe_cev_v2_pagina1_as_dict: {str(e)}")
        return {}  # Return empty dictionary in case of error
    
def get_informe_cev_v2_pagina1_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """
    Extract data from page 1 of an informe_CEV_v2 PDF report and return it as a Pandas DataFrame.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        pd.DataFrame: A DataFrame containing the extracted data fields from page 1.
                      Returns an empty DataFrame if any error occurs during extraction.
    """
    try:
        # Call the existing function to get the data as a dictionary
        data_dict: Dict[str, Any] = get_informe_cev_v2_pagina1_as_dict(pdf_report)

        # Convert the dictionary to a DataFrame
        df: pd.DataFrame = pd.DataFrame.from_dict(data_dict, orient='index').T
        
        return df

    except Exception as e:
        print(f"Error in get_informe_cev_v2_pagina1_as_dataframe: {str(e)}")
        return pd.DataFrame()  # Return an empty DataFrame in case of error
    


def scrape_informe_cev_v2_pagina1(pdf_report):
    # Informe CEV (v.2) - Page 1

    # pdf_report = fitz.open(pdf_file_path)
    page_number = 0  # Page number (starting from 0)
    page = pdf_report[page_number]

    ### Seccion 1: Datos vivienda y Evaluación

    area_coordinates = (8.3, 10.3, 165.6, 65.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    datos_vivienda = extracted_text.splitlines()[-8:]
    datos_vivienda

    index = ['tipo_evaluacion', 'codigo_evaluacion', 'region', 'comuna', 'direccion', 'rol_vivienda_proyecto', 'tipo_vivienda', 'superficie_interior_util_m2']


    # ### Convert list to dictionary
    _dict = dict(zip(index, datos_vivienda))
    _dict['tipo_evaluacion'] = _dict['tipo_evaluacion'].title()
    _dict['superficie_interior_util_m2'] = float(_dict['superficie_interior_util_m2'].replace(',', '.'))

    # Convert dictionary to DataFrame
    df = pd.DataFrame.from_dict(_dict, orient='index').T
    
    # ### Seccion 2: Letra de eﬁciencia energética - Diseño de arquitectura
    area_coordinates = (5.6, 78.6, 165.8, 191.3)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    porcentaje_ahorro_list = extracted_text.splitlines()

    porcentaje_ahorro = None
    for item in porcentaje_ahorro_list:
        if item.replace('-', '').isdigit():
            porcentaje_ahorro = int(item)
            break
    df['porcentaje_ahorro'] = porcentaje_ahorro
    df['letra_eficiencia_energetica_dem'] = _from_procentaje_ahorro_to_letra(porcentaje_ahorro/100)
    
    ### Section 3: Requerimientos anuales de energía para calefacción y enfriamiento

    ### Subsection 1: Demanda energética para calefacción
    area_coordinates = (15.6, 220.0, 73.0, 230.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_calefaccion_kwh_m2_ano = float(extracted_text.splitlines()[-1].replace(',', '.'))
    df['demanda_calefaccion_kwh_m2_ano'] = demanda_calefaccion_kwh_m2_ano
    
    ### Subsection 2: Demanda energética para enfriamiento
    area_coordinates = (90.0, 220.0, 151.5, 230.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_enfriamiento_kwh_m2_ano = float(extracted_text.splitlines()[-1].replace(',', '.'))
    df['demanda_enfriamiento_kwh_m2_ano'] = demanda_enfriamiento_kwh_m2_ano

    ### Subsection 3: Demanda energética total
    area_coordinates = (167.0, 225.0, 209.0, 245.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_total_kwh_m2_ano = float(extracted_text.splitlines()[-1].replace(',', '.'))
    df['demanda_total_kwh_m2_ano'] = demanda_total_kwh_m2_ano

    # ### Subsection 4: Fecha de Emision
    area_coordinates = (35.5, 247.5, 57.0, 255.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    emitida_el = extracted_text.splitlines()[-1]
    emitida_el
    df['emitida_el'] = emitida_el
    # Close the PDF document
    # pdf_report.close()
    # END
    return df

# ------------------------------------------------------------------------------------------------------------
#  Pagina 2    
# ------------------------------------------------------------------------------------------------------------   

def get_informe_cev_v2_pagina2_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 2 of an informe_CEV_v2 PDF report and return it as a dictionary.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        Dict[str, Any]: A dictionary containing the extracted data fields from page 2.
                       Returns an empty dictionary if any error occurs during extraction.
    """
    try:
        # Validate input
        if not isinstance(pdf_report, fitz.Document):
            raise ValueError("Invalid input: pdf_report must be a fitz.Document object.")
        if 1 >= pdf_report.page_count:
            raise ValueError("Invalid page number: Page 2 does not exist in the document.")

        page = pdf_report[1]  # Get page 2 (index 1)

        # Define coordinates as constants for page 2
        COORDINATES: Dict[str, Tuple[float, float, float, float]] = {
            'region': (40.4,47.4,95.0,51.7),
            'comuna': (40.4,53.2,95.0,57.4),
            'direccion': (40.4,58.9,95.0,63.1),
            'rol_vivienda': (40.4,64.6,95.0,68.9),
            'tipo_vivienda': (40.4,70.2,95.0,74.4),
            'zona_termica': (143.8,47.5,146.1,51.7),
            'superficie_interior_util_m2': (143.8,53.3,150,57.5),
            'solicitado_por': (143.8,58.9,185.5,63.1),
            'evaluado_por': (143.8,64.7,210.5,68.9),
            'codigo_evaluacion': (143.8,70.2,160.6,74.5),
            'demanda_calefaccion_kwh_m2_ano': (99.4,99.6,107.6,105.2),
            'demanda_enfriamiento_kwh_m2_ano': (99.2,120.9,107.5,126.5),
            'demanda_total_kwh_m2_ano': (101.5,135.1,130.6,151.4),
            'demanda_total_bis_kwh_m2_ano': (39.2, 159.8, 122.8, 166.0),
            'demanda_total_referencia_kwh_m2_ano': (16.9, 168.3, 146.2, 173.2),
            'porcentaje_ahorro': (152.0, 162.6, 201.5, 168.7),
            'muro_principal_descripcion': (46.2, 202.2, 184.5, 209.1),
            'muro_principal_exigencia_W_m2_K': (185.5, 204.2, 209.5, 209.1),
            'muro_secundario_descripcion': (46.2, 209.2, 184.5, 215.1),
            'muro_secundario_exigencia_W_m2_K': (185.5,211.5,209.5,214.3),
            'piso_principal_descripcion': (46.7,216.6,184.5,219.4),
            'piso_principal_exigencia_W_m2_K': (185.5, 216.4, 209.5, 223.7),
            'puerta_principal_descripcion': (46.2, 223.5, 184.5, 230.2),
            'puerta_principal_exigencia_W_m2_K': (185.5, 223.9, 209.5, 230.2),
            'techo_principal_descripcion': (46.2, 230.5, 184.5, 237.0),
            'techo_principal_exigencia_W_m2_K': (185.5, 230.5, 209.5, 237.2),
            'techo_secundario_descripcion': (46.2, 237.2, 184.5, 244.1),
            'techo_secundario_exigencia_W_m2_K': (185.5, 237.2, 209.5, 244.1),
            'superficie_vidriada_principal_descripcion': (46.2, 244.2, 184.5, 251.0),
            'superficie_vidriada_principal_exigencia': (185.5, 244.3, 209.5, 251.0),
            'superficie_vidriada_secundaria_descripcion': (46.2, 251.3, 184.5, 258.0),
            'superficie_vidriada_secundaria_exigencia': (185.5, 251.3, 209.5, 258.0),
            'ventilacion_rah_descripcion': (46.2, 258.3, 184.5, 265.0),
            'ventilacion_rah_exigencia': (185.5, 258.3, 209.5, 265.0),
            'infiltraciones_rah_descripcion': (46.2, 265.3, 184.5, 272.0),
            'infiltraciones_rah_exigencia': (185.5, 265.3, 209.5, 272.0)
        }

        # Extract all fields at once
        fields: Dict[str, str] = {
            key: extract_text_from_area(page, coords) if coords else None
            for key, coords in COORDINATES.items()
        }
        
        # Create a dictionary to return
        result: Dict[str, Any] = {
            'region': fields.get('region', '').strip(),
            'comuna': fields.get('comuna', '').strip(),
            'direccion': fields.get('direccion', '').strip(),
            'rol_vivienda': fields.get('rol_vivienda', '').strip(),
            'tipo_vivienda': fields.get('tipo_vivienda', '').strip(),
            'zona_termica': fields.get('zona_termica', '').strip(),
            'superficie_interior_util_m2': float(fields.get('superficie_interior_util_m2', '').replace(',', '.').strip()),
            'solicitado_por': fields.get('solicitado_por', '').strip(),
            'evaluado_por': fields.get('evaluado_por', '').strip(),
            'codigo_evaluacion': fields.get('codigo_evaluacion', '').strip(),
            'demanda_calefaccion_kwh_m2_ano': float(fields.get('demanda_calefaccion_kwh_m2_ano', '').splitlines()[-1].replace(',', '.').strip()),
            'demanda_enfriamiento_kwh_m2_ano': float(fields.get('demanda_enfriamiento_kwh_m2_ano', '').splitlines()[-1].replace(',', '.').strip()),
            'demanda_total_kwh_m2_ano': float(fields.get('demanda_total_kwh_m2_ano', '').splitlines()[-1].replace(',', '.').strip()),
            'demanda_total_bis_kwh_m2_ano': float(fields.get('demanda_total_bis_kwh_m2_ano', '').splitlines()[-1].replace(',', '.').strip()),
            'demanda_total_referencia_kwh_m2_ano': float(fields.get('demanda_total_referencia_kwh_m2_ano', '').splitlines()[-1].replace(',', '.').strip()),
            'porcentaje_ahorro': (lambda x: float(x.replace(',', '.').strip()) if x else None)(fields.get('porcentaje_ahorro', '').splitlines()[-1] if fields.get('porcentaje_ahorro', '').splitlines() else None),
            'muro_principal_descripcion': fields.get('muro_principal_descripcion', '').replace('\n', '').strip(),
            'muro_principal_exigencia_W_m2_K': (lambda x: float(x.replace('[W/m2K]', '').replace(',', '.')) if x else None)(fields.get('muro_principal_exigencia_W_m2_K', '')),
            'muro_secundario_descripcion': fields.get('muro_secundario_descripcion', '').replace('\n', '').strip(),
            'muro_secundario_exigencia_W_m2_K': (lambda x: float(x.replace('[W/m2K]', '').replace(',', '.')) if x else None)(fields.get('muro_secundario_exigencia_W_m2_K', '')),
            'piso_principal_descripcion':  fields.get('piso_principal_descripcion', '').replace('\n', '').strip(),
            'piso_principal_exigencia_W_m2_K': (lambda x: float(x.replace('[W/m2K]', '').replace(',', '.')) if x else None)(fields.get('piso_principal_exigencia_W_m2_K', '')),
            'puerta_principal_descripcion': fields.get('puerta_principal_descripcion', '').replace('\n', '').strip(),
            'puerta_principal_exigencia_W_m2_K': fields.get('puerta_principal_exigencia_W_m2_K', '').replace('\n', '').strip(),
            'techo_principal_descripcion': fields.get('techo_principal_descripcion', '').replace('\n', '').strip(),
            'techo_principal_exigencia_W_m2_K': (lambda x: float(x.replace('[W/m2K]', '').replace(',', '.')) if x else None)(fields.get('techo_principal_exigencia_W_m2_K', '')),
            'techo_secundario_descripcion': fields.get('techo_secundario_descripcion', '').replace('\n', '').strip(),
            'techo_secundario_exigencia_W_m2_K': (lambda x: float(x.replace('[W/m2K]', '').replace(',', '.')) if x else None)(fields.get('techo_secundario_exigencia_W_m2_K', '')),
            'superficie_vidriada_principal_descripcion': fields.get('superficie_vidriada_principal_descripcion', '').replace('\n', '').strip(),
            'superficie_vidriada_principal_exigencia': fields.get('superficie_vidriada_principal_exigencia', '').replace('\n', '').strip(),
            'superficie_vidriada_secundaria_descripcion': fields.get('superficie_vidriada_secundaria_descripcion', '').replace('\n', '').strip(),
            'superficie_vidriada_secundaria_exigencia': fields.get('superficie_vidriada_secundaria_exigencia', '').replace('\n', '').strip(),
            'ventilacion_rah_descripcion': fields.get('ventilacion_rah_descripcion', '').replace('\n', '').strip(),
            'ventilacion_rah_exigencia': fields.get('ventilacion_rah_exigencia', '').replace('\n', '').strip(),
            'infiltraciones_rah_descripcion': fields.get('infiltraciones_rah_descripcion', '').replace('\n', '').strip(),
            'infiltraciones_rah_exigencia': fields.get('infiltraciones_rah_exigencia', '').replace('\n', '').strip()
        }

        return result

    except (RuntimeError, IndexError, ValueError) as e:
        logging.error(f"Error in get_informe_cev_v2_pagina2_as_dict: {str(e)}")
        return {}  # Return empty dictionary in case of error

def get_informe_cev_v2_pagina2_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """
    Extract data from page 2 of an informe_CEV_v2 PDF report and return it as a Pandas DataFrame.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        pd.DataFrame: A DataFrame containing the extracted data fields from page 2.
                      Returns an empty DataFrame if any error occurs during extraction.
    """
    try:
        # Call the existing function to get the data as a dictionary
        data_dict: Dict[str, Any] = get_informe_cev_v2_pagina2_as_dict(pdf_report)

        # Convert the dictionary to a DataFrame
        df: pd.DataFrame = pd.DataFrame.from_dict(data_dict, orient='index').T

        return df

    except Exception as e:
        print(f"Error in get_informe_cev_v2_pagina2_as_dataframe: {str(e)}")
        return pd.DataFrame()  # Return an empty DataFrame in case of error    


def scrape_informe_cev_v2_pagina2(pdf_report):
    # Informe CEV (v.2) - Page 1

    # pdf_report = fitz.open(pdf_file_path)
    page_number = 1  # Page number (starting from 0)
    page = pdf_report[page_number]
    # ## Pagina 2

    # ### Seccion 1: Datos vivienda y Evaluación
    # #### Subsection 1: Izquierda
    area_coordinates = (7.8, 46.3, 96.8, 74.2)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    datos_vivienda = extracted_text.splitlines()[-5:]
    # Swap the elements
    datos_vivienda[2], datos_vivienda[3] = datos_vivienda[3], datos_vivienda[2]

    index = ['region', 'comuna', 'direccion', 'rol_vivienda', 'tipo_vivienda']

    _dict = dict(zip(index, datos_vivienda))
    _dict

    # Convert dictionary to DataFrame
    df1 = pd.DataFrame.from_dict(_dict, orient='index').T

    # #### Subsection 2: Derecha
    area_coordinates = (98.6, 46.3, 209.3, 74.2)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    datos_vivienda = extracted_text.splitlines()[-5:]
    
    # Swap the elements
    datos_vivienda[2], datos_vivienda[3] = datos_vivienda[3], datos_vivienda[2]

    index = ['zona_termica', 'superficie_interior_util_m2', 'solicitado_por', 'evaluado_por', 'codigo_evaluacion']
    _dict = dict(zip(index, datos_vivienda))
    _dict['superficie_interior_util_m2'] = float(_dict['superficie_interior_util_m2'].replace(',', '.'))

    # Convert dictionary to DataFrame
    df2 = pd.DataFrame.from_dict(_dict, orient='index').T

    df = pd.concat([df1, df2], axis=1)

    # ### Seccion 2: Demanda energética promedio según tipología y zona térmica (kWh/m2 año)

    # #### Subsection: Demanda Energetica para Calefaccion
    area_coordinates = (99.1, 99.1, 135.3, 105.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_calefaccion_kwh_m2_ano = extracted_text.split('\n')[-1]
    # df['demanda_calefaccion_kwh_m2_ano'] = float(demanda_calefaccion_kwh_m2_ano.replace(',', '.'))
    df['demanda_calefaccion_kwh_m2_ano'] = float(demanda_calefaccion_kwh_m2_ano.replace(',', '.')) if demanda_calefaccion_kwh_m2_ano.strip().replace(',', '.').replace('.', '', 1).isdigit() else None

    # #### Subsection: Demanda Energetica para Enfriamiento

    area_coordinates = (99.1, 120.5, 135.3, 126.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_enfriamiento_kwh_m2_ano = extracted_text.split('\n')[-1]
    # df['demanda_enfriamiento_kwh_m2_ano'] = float(demanda_enfriamiento_kwh_m2_ano.replace(',', '.'))
    df['demanda_enfriamiento_kwh_m2_ano'] = float(demanda_enfriamiento_kwh_m2_ano.replace(',', '.')) if demanda_enfriamiento_kwh_m2_ano.strip().replace(',', '.').replace('.', '', 1).isdigit() else None

    # #### Subsection: Demanda Energetica Total
    area_coordinates = (99.1, 137.4, 135.3, 150.4)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_total_kwh_m2_ano = extracted_text.split('\n')[-1]
    # df['demanda_total_kwh_m2_ano'] = float(demanda_total_kwh_m2_ano.replace(',', '.'))
    df['demanda_total_kwh_m2_ano'] = float(demanda_total_kwh_m2_ano.replace(',', '.')) if demanda_total_kwh_m2_ano.strip().replace(',', '.').replace('.', '', 1).isdigit() else None

    # ### Seccion 3
    # ### Subsection 1: Demanda Energética Total
    area_coordinates = (39.2, 159.8, 122.8, 166.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_total_bis_kwh_m2_ano = extracted_text.split('\n')[-1]
    # df['demanda_total_bis_kwh_m2_ano'] = float(demanda_total_bis_kwh_m2_ano.replace(',', '.'))
    df['demanda_total_bis_kwh_m2_ano'] = float(demanda_total_bis_kwh_m2_ano.replace(',', '.')) if demanda_total_bis_kwh_m2_ano.strip().replace(',', '.').replace('.', '', 1).isdigit() else None

    # ### Subsection 2: Demanda Energética Total de Referencia
    area_coordinates = (16.9, 168.3, 146.2, 173.2)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_total_referencia_kwh_m2_ano = extracted_text.split('\n')[-1]
    df['demanda_total_referencia_kwh_m2_ano'] = float(demanda_total_referencia_kwh_m2_ano.replace(',', '.'))

    # ### Porcentaje de Ahorro
    area_coordinates = (152.0, 162.6, 201.5, 168.7)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    porcentaje_ahorro = extracted_text.split('\n')[-1]
    
    try:
        df['porcentaje_ahorro'] = float(porcentaje_ahorro.replace(',', '.'))
    except ValueError:
        # If it can't be converted to a float, set the value to None
        df['porcentaje_ahorro'] = None

    # ### Seccion 4: Principales características del Diseño de Arquitectura
    # ### Muro Principal

    # Muro principal: Descripcion
    # Define section coordinates
    area_coordinates = (46.2, 202.2, 184.5, 209.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    muro_principal_descripcion = extracted_text.replace('\n', '')
    df['muro_principal_descripcion'] = muro_principal_descripcion

    # Muro principal: Exigencia
    # Define section coordinates
    area_coordinates = (185.5, 202.2, 209.5, 209.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    muro_principal_exigencia = extracted_text.split('\n')[-1]
    try:
        df['muro_principal_exigencia_W_m2_K'] = float(muro_principal_exigencia.replace('[W/m2K]', '').replace(',', '.'))
    except ValueError:
        df['muro_principal_exigencia_W_m2_K'] = None
        
    # ### Muro Secundario
    # Muro secundario: Descripcion
    # Define section coordinates
    area_coordinates = (46.2, 209.2, 184.5, 215.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    muro_secundario_descripcion = extracted_text.replace('\n', '')
    df['muro_secundario_descripcion'] = muro_secundario_descripcion if muro_secundario_descripcion != '0' else None


    # Muro secundario: Exigencia
    # Define section coordinates
    area_coordinates = (185.5, 202.2, 209.5, 209.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    muro_secundario_exigencia = extracted_text.split('\n')[-1]
    try:
        df['muro_secundario_exigencia_W_m2_K'] = float(muro_secundario_exigencia.replace('[W/m2K]', '').replace(',', '.'))
    except ValueError:
        df['muro_secundario_exigencia_W_m2_K'] = None
        

    # ### Piso Principal
    # Piso principal: Descripcion
    # Define section coordinates
    area_coordinates = (46.2, 216.4, 184.5, 223.7)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    piso_principal_descripcion = extracted_text.replace('\n', '')
    df['piso_principal_descripcion'] = piso_principal_descripcion

    # Piso principal: Exigencia
    area_coordinates = (185.5, 216.4, 209.5, 223.7)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    piso_principal_exigencia = extracted_text.split('\n')[-1]
    try:
        df['piso_principal_exigencia_W_m2_K'] = float(piso_principal_exigencia.replace('[W/m2K]', '').replace(',', '.'))
    except ValueError:
        df['piso_principal_exigencia_W_m2_K'] = None

    # ### Puerta principal
    # Puerta principal: Descripcion
    area_coordinates = (46.2, 223.5, 184.5, 230.2)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    puerta_principal_descripcion = extracted_text.replace('\n', '')
    df['puerta_principal_descripcion'] = puerta_principal_descripcion

    # Puerta principal: Exigencia
    # Define section coordinates
    area_coordinates  = (185.5, 223.9, 209.5, 230.2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    puerta_principal_exigencia = extracted_text.strip()
    df['puerta_principal_exigencia_W_m2_K'] = puerta_principal_exigencia if puerta_principal_exigencia.isalpha() else None

    # ### Techo Principal

    # Techo principal: Descripcion
    area_coordinates = (46.2, 230.5, 184.5, 237.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    techo_principal_descripcion = extracted_text.replace('\n', '')
    df['techo_principal_descripcion'] = techo_principal_descripcion

    # Techo principal: Exigencia
    # Define section coordinates
    area_coordinates  = (185.5, 230.5, 209.5, 237.2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    techo_principal_exigencia = extracted_text.strip()
    try:
        df['techo_principal_exigencia_W_m2_K'] = float(techo_principal_exigencia.replace('[W/m2K]', '').replace(',', '.'))
    except ValueError:
        df['techo_principal_exigencia_W_m2_K'] = None

    # ### Techo Secundario
    # Techo secundario: Descripcion
    area_coordinates = (46.2, 237.2, 184.5, 244.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    techo_secundario_descripcion = extracted_text.replace('\n', '')
    df['techo_secundario_descripcion'] = techo_secundario_descripcion if techo_secundario_descripcion != '0' else None

    # Techo secundario: Exigencia
    # Define section coordinates
    area_coordinates  = (185.5, 237.2, 209.5, 244.1)
    extracted_text = extract_text_from_area(page, area_coordinates)
    techo_secundario_exigencia = extracted_text.strip()
    try:
        df['techo_secundario_exigencia_W_m2_K'] = float(techo_secundario_exigencia.replace('[W/m2K]', '').replace(',', '.'))
    except ValueError:
        df['techo_secundario_exigencia_W_m2_K'] = None

    # ### Superficie vidriada principal
    # Superficie vidriada principal: Descripcion
    area_coordinates = (46.2, 244.2, 184.5, 251.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    superficie_vidriada_principal_descripcion = extracted_text.replace('\n', '')
    df['superficie_vidriada_principal_descripcion'] = superficie_vidriada_principal_descripcion

    # Superficie vidriada principal: Exigencia
    # Define section coordinates
    area_coordinates  = (185.5, 244.3, 209.5, 251.0)
    extracted_text = extract_text_from_area(page, area_coordinates)
    superficie_vidriada_principal_exigencia = extracted_text.strip()
    df['superficie_vidriada_principal_exigencia'] = superficie_vidriada_principal_exigencia.strip() if superficie_vidriada_principal_exigencia.isalpha() else None

    # ### Superficie vidriada secundaria
    # Superficie vidriada secundaria: Descripcion
    area_coordinates = (46.2, 251.3, 184.5, 258.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    superficie_vidriada_secundaria_descripcion = extracted_text.replace('\n', '')
    df['superficie_vidriada_secundaria_descripcion'] = superficie_vidriada_secundaria_descripcion if superficie_vidriada_secundaria_descripcion != '0' else None


    # Superficie vidriada secundaria: Exigencia
    # Define section coordinates
    area_coordinates  = (185.5, 251.3, 209.5, 258.0)
    extracted_text = extract_text_from_area(page, area_coordinates)
    superficie_vidriada_secundaria_exigencia = extracted_text.strip()
    df['superficie_vidriada_secundaria_exigencia'] = superficie_vidriada_secundaria_exigencia.strip() if superficie_vidriada_secundaria_exigencia.isalpha() else None

    # ### Ventilación (RAH)
    # Ventilación (RAH): Descripcion
    area_coordinates = (46.2, 258.3, 184.5, 265.0)
    extracted_text = extract_text_from_area(page, area_coordinates)
    ventilacion_rah_descripcion = extracted_text
    df['ventilacion_rah_descripcion'] = ventilacion_rah_descripcion.replace('\n', '').strip()

    # Ventilación (RAH): Exigencia
    # Define section coordinates
    area_coordinates = (185.5, 258.3, 209.5, 265.0)
    extracted_text = extract_text_from_area(page, area_coordinates)
    ventilacion_rah_exigencia = extracted_text
    df['ventilacion_rah_exigencia'] = ventilacion_rah_exigencia.strip()

    # ### Inﬁltraciones (RAH)
    # Infiltraciones (RAH): Descripcion
    area_coordinates = (46.2, 265.3, 184.5, 272.0)
    extracted_text = extract_text_from_area(page, area_coordinates)
    infiltraciones_rah_descripcion = extracted_text
    df['infiltraciones_rah_descripcion'] = infiltraciones_rah_descripcion.replace('\n', '').strip()

    # Infiltraciones (RAH): Exigencia
    area_coordinates = (185.5, 265.3, 209.5, 272.0)
    extracted_text = extract_text_from_area(page, area_coordinates)
    infiltraciones_rah_exigencia = extracted_text
    df['infiltraciones_rah_exigencia'] = infiltraciones_rah_exigencia.replace('\n', '').strip()
    # Close the PDF document
    # pdf_report.close()

    # ### END
    return df

# ------------------------------------------------------------------------------------------------------------
#  Pagina 3 - Consumos    
# ------------------------------------------------------------------------------------------------------------  

def get_informe_cev_v2_pagina3_consumos_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 3 (consumos) of an informe_CEV_v2 PDF report and return it as a dictionary.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        Dict[str, Any]: A dictionary containing the extracted data fields from page 3 (consumos).
                       Returns an empty dictionary if any error occurs during extraction.
    """
    try:
        # Validate input
        if not isinstance(pdf_report, fitz.Document):
            raise ValueError("Invalid input: pdf_report must be a fitz.Document object.")
        if 2 >= pdf_report.page_count:
            raise ValueError("Invalid page number: Page 3 does not exist in the document.")

        page = pdf_report[2]  # Get page 3 (index 2)

        # Define coordinates as constants for page 3
        COORDINATES: Dict[str, Tuple[float, float, float, float]] = {
            'codigo_evaluacion': (62.3, 30.7, 88.1, 35.1), 
            'agua_caliente_sanitaria_kwh_m2': (78.1, 73.9, 98.0, 76.7),
            'agua_caliente_sanitaria_perc': (98.7, 73.9, 116.3, 76.7),
            'iluminacion_kwh_m2': (79.2, 78.1, 98.3, 81.4),
            'iluminacion_per': (98.7, 78.1, 116.3, 81.4),
            'calefaccion_kwh_m2': (79.2, 82.2, 98.3, 86.6),
            'calefaccion_kwh_per': (98.7, 82.2, 116.3, 86.6),
            'energia_renovable_no_convencional_kwh_m2': (79.2, 87.2, 98.3, 91.0),
            'energia_renovable_no_convencional_per': (98.7, 87.2, 116.3, 91.0),
            'consumo_total_kwh_m2': (118.0, 74.0, 148.0, 86.0),
            'emisiones_kgco2_m2_ano': (171.5, 69.0, 183.5, 74.2),
            'calefaccion_descripcion_proy': (76.6, 101.4, 155.5, 105.3),
            'calefaccion_consumo_proy_kwh': (157.0, 101.4, 196.0, 105.3),
            'calefaccion_consumo_proy_per': (198.0, 101.4, 207.0, 105.3),
            'iluminacion_descripcion_proy': (76.6, 106.2, 155.5, 110.0),
            'iluminacion_consumo_proy_kwh': (157.0, 106.2, 196.0, 110.0),
            'iluminacion_consumo_proy_per': (198.0, 106.2, 207.0, 110.0),
            'agua_caliente_sanitaria_descripcion_proy': (76.6, 111.2, 155.5, 115.0),
            'agua_caliente_sanitaria_consumo_proy_kwh': (157.0, 111.2, 196.0, 115.0),
            'agua_caliente_sanitaria_consumo_proy_per': (198.0, 111.2, 207.0, 115.0),
            'energia_renovable_no_convencional_descripcion_proy': (76.6, 115.8, 155.5, 120.0),
            'energia_renovable_no_convencional_consumo_proy_kwh': (157.0, 115.8, 196.0, 120.0),
            'energia_renovable_no_convencional_consumo_proy_per': (198.0, 115.8, 207.0, 120.0),
            'consumo_total_requerido_proy_kwh': (157.0, 121.0, 196.0, 125.0),
            'calefaccion_descripcion_ref': (76.6, 136.1, 155.5, 140.1),
            'calefaccion_consumo_ref_kwh': (157.0, 136.1, 196.0, 140.1),
            'calefaccion_consumo_ref_per': (198.0, 136.1, 207.0, 140.1),
            'iluminacion_descripcion_ref': (76.6, 140.7, 155.5, 144.7),
            'iluminacion_consumo_ref_kwh': (157.0, 140.7, 196.0, 144.7),
            'iluminacion_consumo_ref_per': (198.0, 140.7, 207.0, 144.7),
            'agua_caliente_sanitaria_descripcion_ref': (76.6, 145.2, 155.5, 149.2),
            'agua_caliente_sanitaria_consumo_ref_kwh': (157.0, 145.2, 196.0, 149.2),
            'agua_caliente_sanitaria_consumo_ref_per': (198.0, 145.2, 207.0, 149.2),
            'energia_renovable_no_convencional_descripcion_ref': (76.6, 150.8, 155.5, 154.8),
            'energia_renovable_no_convencional_consumo_ref_kwh': (157.0, 150.8, 196.0, 154.8),
            'energia_renovable_no_convencional_consumo_ref_per': (198.0, 150.8, 207.0, 154.8),
            'consumo_total_requerido_ref_kwh': (157.0, 156.0, 196.0, 160.0),
            'consumo_ep_calefaccion_kwh': (87.0, 176.0, 104.0, 179.0),
            'consumo_ep_agua_caliente_sanitaria_kwh': (87.0, 180.0, 104.0, 183.5),
            'consumo_ep_iluminacion_kwh': (87.0, 184.0, 104.0, 187.5),
            'consumo_ep_ventiladores_kwh': (87.0, 188.0, 104.0, 191.5),
            'generacion_ep_fotovoltaicos_kwh': (87.0, 199.0, 104.0, 202.5),
            'aporte_fotovoltaicos_consumos_basicos_kwh': (87.0, 203.2, 104.0, 206.0),
            'diferencia_fotovoltaica_para_consumo_kwh': (87.0, 206.9, 104.0, 210.2),
            'aporte_solar_termica_consumos_basicos_kwh': (87.0, 218.0, 104.0, 221.0),
            'aporte_solar_termica_agua_caliente_sanitaria_kwh': (87.0, 222.5, 104.0, 225.5),
            'total_consumo_ep_antes_fotovoltaica_kwh': (192.0, 176.0, 208.0, 179.5),
            'aporte_fotovoltaicos_consumos_basicos_kwh_bis': (192.0, 180.0, 208.0, 183.5),
            'consumos_basicos_a_suplir_kwh': (192.0, 184.3, 208.0, 187.0),
            'consumo_total_ep_obj_kwh': (192.0, 199.0, 208.0, 202.5),
            'consumo_total_ep_ref_kwh': (192.0, 202.8, 208.0, 206.5),
            'coeficiente_energetico_c': (192.0, 207.0, 208.0, 210.5)
        }


        # Extract all fields at once
        fields: Dict[str, str] = {
            key: extract_text_from_area(page, coords) if coords else None
            for key, coords in COORDINATES.items()
        }

        # Create a dictionary to return
        result: Dict[str, Any] = {
            'codigo_evaluacion': fields.get('codigo_evaluacion', '').strip(),
            'agua_caliente_sanitaria_kwh_m2': float(fields.get('agua_caliente_sanitaria_kwh_m2', '').replace(',', '.')) if fields.get('agua_caliente_sanitaria_kwh_m2', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'agua_caliente_sanitaria_perc': float(fields.get('agua_caliente_sanitaria_perc', '').replace(',', '.')) if fields.get('agua_caliente_sanitaria_perc', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'iluminacion_kwh_m2': float(fields.get('iluminacion_kwh_m2', '').replace(',', '.')) if fields.get('iluminacion_kwh_m2', '').replace(',', '.').replace('.', '', 1).isdigit() else None,            
            'iluminacion_per': float(fields.get('iluminacion_per', '').replace(',', '.')) if fields.get('iluminacion_per', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'calefaccion_kwh_m2': float(fields.get('calefaccion_kwh_m2', '').replace(',', '.')) if fields.get('calefaccion_kwh_m2', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'calefaccion_kwh_per': float(fields.get('calefaccion_kwh_per', '').replace(',', '.')) if fields.get('calefaccion_kwh_per', '').replace(',', '.').replace('.', '', 1).isdigit() else None,            
            'energia_renovable_no_convencional_kwh_m2': float(fields.get('energia_renovable_no_convencional_kwh_m2', '').replace(',', '.')) if fields.get('energia_renovable_no_convencional_kwh_m2', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'energia_renovable_no_convencional_per': float(fields.get('energia_renovable_no_convencional_per', '').replace(',', '.')) if fields.get('energia_renovable_no_convencional_per', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'consumo_total_kwh_m2': float(fields.get('consumo_total_kwh_m2', '').replace(',', '.')) if fields.get('consumo_total_kwh_m2', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'emisiones_kgco2_m2_ano': float(fields.get('emisiones_kgco2_m2_ano', '').replace(',', '.')) if fields.get('emisiones_kgco2_m2_ano', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            # Calefacción
            'calefaccion_descripcion_proy': fields.get('calefaccion_descripcion_proy', '').strip(),
            'calefaccion_consumo_proy_kwh': float(fields.get('calefaccion_consumo_proy_kwh', '').split('\n')[-1].replace(',', '.')) if fields.get('calefaccion_consumo_proy_kwh', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            'calefaccion_consumo_proy_per': float(fields.get('calefaccion_consumo_proy_per', '').split('\n')[-1].replace(',', '.')) if fields.get('calefaccion_consumo_proy_per', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            # Iluminación
            'iluminacion_descripcion_proy': fields.get('iluminacion_descripcion_proy', '').split('\n')[0].strip(),
            'iluminacion_consumo_proy_kwh': float(fields.get('iluminacion_consumo_proy_kwh', '').split('\n')[-1].replace(',', '.')) if fields.get('iluminacion_consumo_proy_kwh', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            'iluminacion_consumo_proy_per': float(fields.get('iluminacion_consumo_proy_per', '').split('\n')[-1].replace(',', '.')) if fields.get('iluminacion_consumo_proy_per', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            # Agua Caliente Sanitaria
            'agua_caliente_sanitaria_descripcion_proy': fields.get('agua_caliente_sanitaria_descripcion_proy', '').strip(),
            'agua_caliente_sanitaria_consumo_proy_kwh': float(fields.get('agua_caliente_sanitaria_consumo_proy_kwh', '').split('\n')[-1].replace(',', '.')) if fields.get('agua_caliente_sanitaria_consumo_proy_kwh', '').split('\n')[-1].replace(',','.')  .replace('.', '', 1).isdigit() else None,
            'agua_caliente_sanitaria_consumo_proy_per': float(fields.get('agua_caliente_sanitaria_consumo_proy_per', '').split('\n')[-1].replace(',', '.')) if fields.get('agua_caliente_sanitaria_consumo_proy_per', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            
            # Energía renovable no convencional
            'energia_renovable_no_convencional_descripcion_proy': fields.get('energia_renovable_no_convencional_descripcion_proy', '').strip(),
            'energia_renovable_no_convencional_consumo_proy_kwh': float(fields.get('energia_renovable_no_convencional_consumo_proy_kwh', '').split('\n')[-1].replace(',', '.')) if fields.get('energia_renovable_no_convencional_consumo_proy_kwh', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            'energia_renovable_no_convencional_consumo_proy_per': float(fields.get('energia_renovable_no_convencional_consumo_proy_per', '').split('\n')[-1].replace(',', '.')) if fields.get('energia_renovable_no_convencional_consumo_proy_per', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,

            'consumo_total_requerido_proy_kwh': float(fields.get('consumo_total_requerido_proy_kwh', '').split('\n')[-1].replace(',', '.')) if fields.get('consumo_total_requerido_proy_kwh', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            # ### Equipos de referencia
            # Calefacción
            'calefaccion_descripcion_ref': fields.get('calefaccion_descripcion_ref', '').strip(),
            'calefaccion_consumo_ref_kwh': float(fields.get('calefaccion_consumo_ref_kwh', '').split('\n')[-1].replace(',', '.')) if fields.get('calefaccion_consumo_ref_kwh', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            'calefaccion_consumo_ref_per': float(fields.get('calefaccion_consumo_ref_per', '').split('\n')[-1].replace(',', '.')) if fields.get('calefaccion_consumo_ref_per', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            # Iluminación
            'iluminacion_descripcion_ref': fields.get('iluminacion_descripcion_ref', '').strip(),
            'iluminacion_consumo_ref_kwh': float(fields.get('iluminacion_consumo_ref_kwh', '').split('\n')[-1].replace(',', '.')) if fields.get('iluminacion_consumo_ref_kwh', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            'iluminacion_consumo_ref_per': float(fields.get('iluminacion_consumo_ref_per', '').split('\n')[-1].replace(',', '.')) if fields.get('iluminacion_consumo_ref_per', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            # Agua Caliente Sanitaria
            'agua_caliente_sanitaria_descripcion_ref': fields.get('agua_caliente_sanitaria_descripcion_ref', '').strip(),
            'agua_caliente_sanitaria_consumo_ref_kwh': float(fields.get('agua_caliente_sanitaria_consumo_ref_kwh', '').split('\n')[-1].replace(',', '.')) if fields.get('agua_caliente_sanitaria_consumo_ref_kwh', '').split('\n')[-1].replace(',', '', 1).isdigit() else None, 
            'agua_caliente_sanitaria_consumo_ref_per': float(fields.get('agua_caliente_sanitaria_consumo_ref_per', '').split('\n')[-1].replace(',', '.')) if fields.get('agua_caliente_sanitaria_consumo_ref_per', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None, 
            # Energía renovable no convencional
            'energia_renovable_no_convencional_descripcion_ref': fields.get('energia_renovable_no_convencional_descripcion_ref', '').strip(),
            'energia_renovable_no_convencional_consumo_ref_kwh': float(fields.get('energia_renovable_no_convencional_consumo_ref_kwh', '').split('\n')[-1].replace(',', '.')) if fields.get('energia_renovable_no_convencional_consumo_ref_kwh', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            'energia_renovable_no_convencional_consumo_ref_per': float(fields.get('energia_renovable_no_convencional_consumo_ref_per', '').split('\n')[-1].replace(',', '.')) if fields.get('energia_renovable_no_convencional_consumo_ref_per', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None, 
            'consumo_total_requerido_ref_kwh': float(fields.get('consumo_total_requerido_ref_kwh', '').split('\n')[-1].replace(',', '.')) if fields.get('consumo_total_requerido_ref_kwh', '').split('\n')[-1].replace(',', '.').replace('.', '', 1).isdigit() else None,
            # ### REQUERIMIENTOS DE ENERGÍA (kWh/año)
            # #### CONSUMOS SIN INCLUIR ERNC
            'consumo_ep_calefaccion_kwh': float(fields.get('consumo_ep_calefaccion_kwh', '').replace(',', '.')) if fields.get('consumo_ep_calefaccion_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'consumo_ep_agua_caliente_sanitaria_kwh': float(fields.get('consumo_ep_agua_caliente_sanitaria_kwh', '').replace(',', '.')) if fields.get('consumo_ep_agua_caliente_sanitaria_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'consumo_ep_iluminacion_kwh': float(fields.get('consumo_ep_iluminacion_kwh', '').replace(',', '.')) if fields.get('consumo_ep_iluminacion_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'consumo_ep_ventiladores_kwh': float(fields.get('consumo_ep_ventiladores_kwh', '').replace(',', '.')) if fields.get('consumo_ep_ventiladores_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'generacion_ep_fotovoltaicos_kwh': float(fields.get('generacion_ep_fotovoltaicos_kwh', '').replace(',', '.')) if fields.get('generacion_ep_fotovoltaicos_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'aporte_fotovoltaicos_consumos_basicos_kwh': float(fields.get('aporte_fotovoltaicos_consumos_basicos_kwh', '').replace(',', '.')) if fields.get('aporte_fotovoltaicos_consumos_basicos_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'diferencia_fotovoltaica_para_consumo_kwh': float(fields.get('diferencia_fotovoltaica_para_consumo_kwh', '').replace(',', '.')) if fields.get('diferencia_fotovoltaica_para_consumo_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'aporte_solar_termica_consumos_basicos_kwh': float(fields.get('aporte_solar_termica_consumos_basicos_kwh', '').replace(',', '.')) if fields.get('aporte_solar_termica_consumos_basicos_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'aporte_solar_termica_agua_caliente_sanitaria_kwh': float(fields.get('aporte_solar_termica_agua_caliente_sanitaria_kwh', '').replace(',', '.')) if fields.get('aporte_solar_termica_agua_caliente_sanitaria_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'total_consumo_ep_antes_fotovoltaica_kwh': float(fields.get('total_consumo_ep_antes_fotovoltaica_kwh', '').replace(',', '.')) if fields.get('total_consumo_ep_antes_fotovoltaica_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'aporte_fotovoltaicos_consumos_basicos_kwh_bis': float(fields.get('aporte_fotovoltaicos_consumos_basicos_kwh_bis', '').replace(',', '.')) if fields.get('aporte_fotovoltaicos_consumos_basicos_kwh_bis', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'consumos_basicos_a_suplir_kwh': float(fields.get('consumos_basicos_a_suplir_kwh', '').replace(',', '.')) if fields.get('consumos_basicos_a_suplir_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'consumo_total_ep_obj_kwh': float(fields.get('consumo_total_ep_obj_kwh', '').replace(',', '.')) if fields.get('consumo_total_ep_obj_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,
            'consumo_total_ep_ref_kwh': float(fields.get('consumo_total_ep_ref_kwh', '').replace(',', '.')) if fields.get('consumo_total_ep_ref_kwh', '').replace(',', '.').replace('.', '', 1).isdigit() else None,          
            'coeficiente_energetico_c': float(fields.get('coeficiente_energetico_c', '').replace(',', '.')) if fields.get('coeficiente_energetico_c', '').replace(',', '.').replace('.', '', 1).isdigit() else None            
        }

        return result

    except (RuntimeError, IndexError, ValueError) as e:
        logging.error(f"Error in get_informe_cev_v2_pagina3_consumos_as_dict: {str(e)}")
        return {}  # Return empty dictionary in case of error
    
def get_informe_cev_v2_pagina3_consumos_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """
    Extract data from page 3 (consumos) of an informe_CEV_v2 PDF report and return it as a Pandas DataFrame.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        pd.DataFrame: A DataFrame containing the extracted data fields from page 3 (consumos).
                      Returns an empty DataFrame if any error occurs during extraction.
    """
    try:
        # Call the existing function to get the data as a dictionary
        data_dict: Dict[str, Any] = get_informe_cev_v2_pagina3_consumos_as_dict(pdf_report)

        # Convert the dictionary to a DataFrame
        df: pd.DataFrame = pd.DataFrame.from_dict(data_dict, orient='index').T

        return df

    except Exception as e:
        print(f"Error in get_informe_cev_v2_pagina3_consumos_as_dataframe: {str(e)}")
        return pd.DataFrame()  # Return an empty DataFrame in case of error

def scrape_informe_cev_v2_pagina3_consumos(pdf_report):

    # # Informe CEV (v.2) - Page 3
    #pdf_report = fitz.open(pdf_file_path)
    page_number = 2  # Page number (starting from 0)
    page = pdf_report[page_number]


    # ## Pagina 3
    # Código evaluación energética
    area_coordinates = (62.3, 30.7, 88.1, 35.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    codigo_evaluacion = extracted_text
    codigo_evaluacion

    df = pd.DataFrame(data=[codigo_evaluacion], columns=['codigo_evaluacion'])
 
    # ## DISTRIBUCIÓN DEL CONSUMO ENERGÉTICO ARQUITECTURA + EQUIPOS + TIPO DE ENERGÍA
    ## Agua caliente sanitaria
    area_coordinates = (78.1, 73.9, 98.0, 76.7)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    agua_caliente_sanitaria_kwh_m2 = float(extracted_text.replace(',', '.')) if extracted_text.replace(',', '.').replace('.', '', 1).isdigit() else None
    df['agua_caliente_sanitaria_kwh_m2'] = agua_caliente_sanitaria_kwh_m2

    area_coordinates = (98.7, 73.9, 116.3, 76.7)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    agua_caliente_sanitaria_perc = float(extracted_text.replace(',', '.')) if extracted_text.replace(',', '.').replace('.', '', 1).isdigit() else None
    df['agua_caliente_sanitaria_perc'] = agua_caliente_sanitaria_perc
   
    ## Iluminación
    area_coordinates = (79.2, 78.1, 98.3, 81.4)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    iluminacion_kwh_m2 = float(extracted_text.replace(',', '.')) if extracted_text.replace(',', '.').replace('.', '', 1).isdigit() else None
    df['iluminacion_kwh_m2'] = iluminacion_kwh_m2
   
    area_coordinates = (98.7, 78.1, 116.3, 81.4)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    iluminacion_per = float(extracted_text.replace(',', '.')) if extracted_text.replace(',', '.').replace('.', '', 1).isdigit() else None
    df['iluminacion_per'] = iluminacion_per
  
    # Calefacción
    area_coordinates = (79.2, 82.2, 98.3, 86.6)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    calefaccion_kwh_m2 = float(extracted_text.replace(',', '.')) if extracted_text.replace(',', '.').replace('.', '', 1).isdigit() else None
    df['calefaccion_kwh_m2'] = calefaccion_kwh_m2
   
    area_coordinates = (98.7, 82.2, 116.3, 86.6)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    calefaccion_kwh_per = float(extracted_text.replace(',', '.')) if extracted_text.replace(',', '.').replace('.', '', 1).isdigit() else None
    df['calefaccion_kwh_per'] = calefaccion_kwh_per

    # Energía renovable no convencional
    area_coordinates = (79.2, 87.2, 98.3, 91.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    ernc_kwh_m2 = float(extracted_text.replace(',', '.')) if extracted_text.replace(',', '.').replace('.', '', 1).isdigit() else None
    df['energia_renovable_no_convencional_kwh_m2'] = ernc_kwh_m2
 
    area_coordinates = (98.7, 87.2, 116.3, 91.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    ernc_per = float(extracted_text.replace(',', '.')) if extracted_text.replace(',', '.').replace('.', '', 1).isdigit() else None
    df['energia_renovable_no_convencional_per'] = ernc_per
 
    # Consumo Total por m²
    area_coordinates = (118.0, 74.0, 148.0, 86.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    consumo_kwh_m2 = float(extracted_text.replace(',', '.')) if extracted_text.replace(',', '.').replace('.', '', 1).isdigit() else None
    df['consumo_total_kwh_m2'] = consumo_kwh_m2
  
    # Emisiones de CO2e
    area_coordinates = (171.5, 69.0, 183.5, 74.2)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    emisiones_kgco2_m2_ano = float(extracted_text.replace(',', '.')) if extracted_text.strip().replace(',', '.').replace('.', '', 1).isdigit() else None
    df['emisiones_kgco2_m2_ano'] = emisiones_kgco2_m2_ano
  
    # Calefacción
    area_coordinates = (76.6, 101.4, 155.5, 105.3)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    calefaccion_descripcion_proy = extracted_text.strip()
    df['calefaccion_descripcion_proy'] = calefaccion_descripcion_proy
 
    area_coordinates = (157.0, 101.4, 196.0, 105.3)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    calefaccion_consumo_proy_kwh = extracted_text.split('\n')[-1].replace(',', '.')
    df['calefaccion_consumo_proy_kwh'] = float(calefaccion_consumo_proy_kwh.replace(',', '.')) if calefaccion_consumo_proy_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None
  
    area_coordinates = (198.0, 101.4, 207.0, 105.3)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    calefaccion_consumo_proy_per = extracted_text.split('\n')[-1].replace(',', '.')
    df['calefaccion_consumo_proy_per'] = float(calefaccion_consumo_proy_per) if calefaccion_consumo_proy_per.replace(',', '.').replace('.', '', 1).isdigit() else None
  
    # Iluminación
    area_coordinates = (76.6, 106.2, 155.5, 110.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    iluminacion_descripcion_proy = extracted_text.split('\n')[0].strip()
    df['iluminacion_descripcion_proy'] = iluminacion_descripcion_proy
  
    area_coordinates = (157.0, 106.2, 196.0, 110.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    iluminacion_consumo_proy_kwh = extracted_text.split('\n')[-1].replace(',', '.')
    df['iluminacion_consumo_proy_kwh'] = float(iluminacion_consumo_proy_kwh) if iluminacion_consumo_proy_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None
   
    area_coordinates = (198.0, 106.2, 207.0, 110.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    iluminacion_consumo_proy_per = extracted_text.split('\n')[-1].replace(',', '.')
    df['iluminacion_consumo_proy_per'] = float(iluminacion_consumo_proy_per) if iluminacion_consumo_proy_per.replace(',', '.').replace('.', '', 1).isdigit() else None
    
    # Agua Caliente Sanitaria
    area_coordinates = (76.6, 111.2, 155.5, 115.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    acs_descripcion_proy = extracted_text
    df['agua_caliente_sanitaria_descripcion_proy'] = acs_descripcion_proy
    
    area_coordinates = (157.0, 111.2, 196.0, 115.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    acs_consumo_proy_kwh = extracted_text.split('\n')[-1].replace(',', '.')
    df['agua_caliente_sanitaria_consumo_proy_kwh'] = float(acs_consumo_proy_kwh) if acs_consumo_proy_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None
    
    area_coordinates = (198.0, 111.2, 207.0, 115.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    acs_consumo_proy_per = extracted_text.split('\n')[-1].replace(',', '.')
    df['agua_caliente_sanitaria_consumo_proy_per'] = float(acs_consumo_proy_per) if acs_consumo_proy_per.replace(',', '.').replace('.', '', 1).isdigit() else None
    
    # Energía renovable no convencional
    area_coordinates = (76.6, 115.8, 155.5, 120.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    enrc_descripcion_proy = extracted_text.split('\n')[0].strip()
    df['energia_renovable_no_convencional_descripcion_proy'] = enrc_descripcion_proy
   
    area_coordinates = (157.0, 115.8, 196.0, 120.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    ernc_consumo_proy_kwh = extracted_text.split('\n')[-1].replace(',', '.')
    df['energia_renovable_no_convencional_consumo_proy_kwh'] = float(ernc_consumo_proy_kwh) if ernc_consumo_proy_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (198.0, 115.8, 207.0, 120.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    ernc_consumo_proy_per = extracted_text.split('\n')[-1].replace(',', '.')
    df['energia_renovable_no_convencional_consumo_proy_per'] = float(ernc_consumo_proy_per) if ernc_consumo_proy_per.replace(',', '.').replace('.', '', 1).isdigit() else None
   
    area_coordinates = (157.0, 121.0, 196.0, 125.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    consumo_total_proy_kwh = extracted_text.split('\n')[-1].replace(',', '.')
    df['consumo_total_requerido_proy_kwh'] = float(consumo_total_proy_kwh) if consumo_total_proy_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    # ### Equipos de referencia

   # Calefacción
    area_coordinates = (76.6, 136.1, 155.5, 140.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    calefaccion_descripcion_ref = extracted_text.strip()
    df['calefaccion_descripcion_ref'] = calefaccion_descripcion_ref
 
    area_coordinates = (157.0, 136.1, 196.0, 140.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    calefaccion_consumo_ref_kwh = extracted_text.split('\n')[-1].replace(',', '.')
    df['calefaccion_consumo_ref_kwh'] = float(calefaccion_consumo_ref_kwh) if calefaccion_consumo_ref_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (198.0, 136.1, 207.0, 140.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    calefaccion_consumo_ref_per = extracted_text.split('\n')[-1].replace(',', '.')
    df['calefaccion_consumo_ref_per'] = float(calefaccion_consumo_ref_per) if calefaccion_consumo_ref_per.replace(',', '.').replace('.', '', 1).isdigit() else None
  
    # Iluminación
    area_coordinates = (76.6, 140.7, 155.5, 144.7)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    iluminacion_descripcion_ref = extracted_text.split('\n')[0].strip()
    df['iluminacion_descripcion_ref'] = iluminacion_descripcion_ref

    area_coordinates = (157.0, 140.7, 196.0, 144.7)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    iluminacion_consumo_ref_kwh = extracted_text.split('\n')[-1].replace(',', '.')
    df['iluminacion_consumo_ref_kwh'] = float(iluminacion_consumo_ref_kwh) if iluminacion_consumo_ref_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (198.0, 140.7, 207.0, 144.7)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    iluminacion_consumo_ref_per = extracted_text.split('\n')[-1].replace(',', '.')
    df['iluminacion_consumo_ref_per'] = float(iluminacion_consumo_ref_per) if iluminacion_consumo_ref_per.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    # Agua Caliente Sanitaria
    area_coordinates = (76.6, 145.2, 155.5, 149.2)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    acs_descripcion_ref = extracted_text
    df['agua_caliente_sanitaria_descripcion_ref'] = acs_descripcion_ref
  
    area_coordinates = (157.0, 145.2, 196.0, 149.2)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    acs_consumo_ref_kwh = extracted_text.split('\n')[-1].replace(',', '.')
    df['agua_caliente_sanitaria_consumo_ref_kwh'] = float(acs_consumo_ref_kwh) if acs_consumo_ref_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (198.0, 145.2, 207.0, 149.2)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    acs_consumo_ref_per = extracted_text.split('\n')[-1].replace(',', '.')
    df['agua_caliente_sanitaria_consumo_ref_per'] = float(acs_consumo_ref_per) if acs_consumo_ref_per.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    # Energía renovable no convencional
    area_coordinates = (76.6, 150.8, 155.5, 154.8)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    enrc_descripcion_ref = extracted_text.split('\n')[0].strip()
    df['energia_renovable_no_convencional_descripcion_ref'] = enrc_descripcion_ref
  
    area_coordinates = (157.0, 150.8, 196.0, 154.8)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    ernc_consumo_ref_kwh = extracted_text.split('\n')[-1].replace(',', '.')
    df['energia_renovable_no_convencional_consumo_ref_kwh'] = float(ernc_consumo_ref_kwh) if ernc_consumo_ref_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (198.0, 150.8, 207.0, 154.8)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    ernc_consumo_ref_per = extracted_text.split('\n')[-1].replace(',', '.')
    df['energia_renovable_no_convencional_consumo_ref_per'] = float(ernc_consumo_ref_per) if ernc_consumo_ref_per.replace(',', '.').replace('.', '', 1).isdigit() else None
   
    area_coordinates = (157.0, 156.0, 196.0, 160.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    consumo_total_ref_kwh = extracted_text.split('\n')[-1].replace(',', '.')
    df['consumo_total_requerido_ref_kwh'] = float(consumo_total_ref_kwh) if consumo_total_ref_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None
  
    # ### REQUERIMIENTOS DE ENERGÍA (kWh/año)
    # #### CONSUMOS SIN INCLUIR ERNC

    area_coordinates = (87.0, 176.0, 104.0, 179.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    consumos_sin_incluir_ernc_calef = extracted_text.replace(',', '.')
    df['consumo_ep_calefaccion_kwh'] = float(consumos_sin_incluir_ernc_calef) if consumos_sin_incluir_ernc_calef.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (87.0, 180.0, 104.0, 183.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    consumos_sin_incluir_ernc_acs = extracted_text.replace(',', '.')
    df['consumo_ep_agua_caliente_sanitaria_kwh'] = float(consumos_sin_incluir_ernc_acs) if consumos_sin_incluir_ernc_acs.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (87.0, 184.0, 104.0, 187.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    consumos_sin_incluir_ernc_ilum = extracted_text.replace(',', '.')
    df['consumo_ep_iluminacion_kwh'] = float(consumos_sin_incluir_ernc_ilum) if consumos_sin_incluir_ernc_ilum.replace(',', '.').replace('.', '', 1).isdigit() else None

    area_coordinates = (87.0, 188.0, 104.0, 191.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    consumos_sin_incluir_ernc_vent = extracted_text.replace(',', '.')
    df['consumo_ep_ventiladores_kwh'] = float(consumos_sin_incluir_ernc_vent) if consumos_sin_incluir_ernc_vent.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    # #### GENERACIÓN FOTOVOLTAICA EN LA VIVIENDA
    area_coordinates = (87.0, 199.0, 104.0, 202.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    generacion_fotovoltaica_ep = extracted_text.replace(',', '.')
    df['generacion_ep_fotovoltaicos_kwh'] = float(generacion_fotovoltaica_ep) if generacion_fotovoltaica_ep.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (87.0, 203.2, 104.0, 206.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    generacion_fotovoltaica_aporte = extracted_text.replace(',', '.')
    df['aporte_fotovoltaicos_consumos_basicos_kwh'] = float(generacion_fotovoltaica_aporte) if generacion_fotovoltaica_aporte.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (87.0, 206.9, 104.0, 210.2)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    generacion_fotovoltaica_consumo = extracted_text.replace(',', '.')
    df['diferencia_fotovoltaica_para_consumo_kwh'] = float(generacion_fotovoltaica_consumo) if generacion_fotovoltaica_consumo.replace(',', '.').replace('.', '', 1).isdigit() else None
   
    # #### DISTRIBUCIÓN DEL APORTE DE SOLAR TÉRMICA

    area_coordinates = (87.0, 218.0, 104.0, 221.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    aporte_solar_termica_calef = extracted_text.replace(',', '.')
    df['aporte_solar_termica_consumos_basicos_kwh'] = float(aporte_solar_termica_calef) if aporte_solar_termica_calef.replace(',', '.').replace('.', '', 1).isdigit() else None
   
    area_coordinates = (87.0, 222.5, 104.0, 225.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    aporte_solar_termica_acs = extracted_text.replace(',', '.')
    df['aporte_solar_termica_agua_caliente_sanitaria_kwh'] = float(aporte_solar_termica_acs) if aporte_solar_termica_acs.replace(',', '.').replace('.', '', 1).isdigit() else None

    # #### BALANCE GENERAL DE ENERGÍA
    area_coordinates = (192.0, 176.0, 208.0, 179.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    balance_general_energia_antes = extracted_text.replace(',', '.')
    df['total_consumo_ep_antes_fotovoltaica_kwh'] = float(balance_general_energia_antes) if balance_general_energia_antes.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (192.0, 180.0, 208.0, 183.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    balance_general_energia_aporte_fv = extracted_text.replace(',', '.')
    df['aporte_fotovoltaicos_consumos_basicos_kwh_bis'] = float(balance_general_energia_aporte_fv) if balance_general_energia_aporte_fv.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    area_coordinates = (192.0, 184.3, 208.0, 187.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    balance_general_energia_suplir = extracted_text.replace(',', '.')
    df['consumos_basicos_a_suplir_kwh'] = float(balance_general_energia_suplir) if balance_general_energia_suplir.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    # #### RESUMEN DE CONSUMOS FINALES DE REFERENCIA Y OBJETO
    area_coordinates = (192.0, 199.0, 208.0, 202.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    consumo_total_ep_obj_kwh = extracted_text.replace(',', '.')
    df['consumo_total_ep_obj_kwh'] = float(consumo_total_ep_obj_kwh) if consumo_total_ep_obj_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None

    area_coordinates = (192.0, 202.8, 208.0, 206.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    consumo_total_ep_ref_kwh = extracted_text.replace(',', '.')
    df['consumo_total_ep_ref_kwh'] = float(consumo_total_ep_ref_kwh) if consumo_total_ep_ref_kwh.replace(',', '.').replace('.', '', 1).isdigit() else None

    area_coordinates = (192.0, 207.0, 208.0, 210.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    coeficiente_energetico_c = extracted_text.replace(',', '.')
    df['coeficiente_energetico_c'] = float(coeficiente_energetico_c) if coeficiente_energetico_c.replace(',', '.').replace('.', '', 1).isdigit() else None
 
    # Close the PDF document
    # pdf_report.close()
    return df


# ------------------------------------------------------------------------------------------------------------
#  Pagina 3 - Consumos    
# ------------------------------------------------------------------------------------------------------------  
def get_informe_cev_v2_pagina3_envolvente_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 3 (envolvente) of an informe_CEV_v2 PDF report and return it as a dictionary.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        Dict[str, Any]: A dictionary containing the extracted data fields from page 3 (envolvente).
                       Returns an empty dictionary if any error occurs during extraction.
    """
    try:
        # Validate input
        if not isinstance(pdf_report, fitz.Document):
            raise ValueError("Invalid input: pdf_report must be a fitz.Document object.")
        if 2 >= pdf_report.page_count:
            raise ValueError("Invalid page number: Page 3 does not exist in the document.")

        page = pdf_report[2]  # Get page 3 (index 2)

        # Define coordinates as constants for page 3
        dy = 4.2
        COORDINATES: Dict[str, Tuple[float, float, float, float]] = {
                    'codigo_evaluacion': (62.3, 30.7, 88.1, 35.1),
                    'elementos_opacos_area_m2': (19.5, 245.0, 47.0, 287.5),
                    'elementos_opacos_U_W_m2_K': (47.8, 245.0, 60.5, 287.5),
                    'elementos_traslucidos_area_m2': (68.2, 245.0, 89.5, 283.0),
                    'elementos_traslucidos_U_W_m2_K': (90.4, 245.0, 103.1, 283.0),
                    'P01_W_K': [(115.5, 250.0 + i * dy, 124.5, 253.5 + i * dy) for i in range(8)],
                    'P02_W_K': [(126.2, 250.0 + i * dy, 136.9, 253.5 + i * dy) for i in range(8)],
                    'P03_W_K': [(139.0, 250.0 + i * dy, 148.2, 253.5 + i * dy) for i in range(8)],
                    'P04_W_K': [(149.0, 250.0 + i * dy, 160.0, 253.5 + i * dy) for i in range(8)],
                    'P05_W_K': [(161.3, 250.0 + i * dy, 171.2, 253.5 + i * dy) for i in range(8)],
                    'UA_phiL': [(190.5, 245.5 + i * dy, 201.0, 249.0 + i * dy) for i in range(10)]               
                }

       # Extract all fields at once
        fields: Dict[str, Union[str, List[str]]] = {
            key: (
                [extract_text_from_area(page, coord) for coord in coords] if isinstance(coords, list) else 
                extract_text_from_area(page, coords) if coords else None) for key, coords in COORDINATES.items()
}

        # Create a dictionary to return
        result: Dict[str, Any] = {
            'codigo_evaluacion': fields.get('codigo_evaluacion', '').strip(),
            'orientacion': ['Horiz', 'N', 'NE', 'E', 'SE', 'S', 'SO', 'O', 'NO', 'Pisos'],
            'elementos_opacos_area_m2': [float(x.replace(',', '.')) for x in fields.get('elementos_opacos_area_m2', ['']).splitlines()[-10:]],
            'elementos_opacos_U_W_m2_K': [float(x.replace(',', '.')) for x in fields.get('elementos_opacos_U_W_m2_K', ['']).splitlines()[-10:]],
            'elementos_traslucidos_area_m2': [float(x.replace(',', '.')) for x in fields.get('elementos_traslucidos_area_m2', ['']).splitlines()[-9:]] + [0],
            'elementos_traslucidos_U_W_m2_K': [float(x.replace(',', '.')) for x in fields.get('elementos_traslucidos_U_W_m2_K', ['']).splitlines()[-9:]] + [0],
            'P01_W_K': [0] + [float(x.splitlines()[-1].replace(',', '.')) for x in fields.get('P01_W_K', [''])] + [0],            
            'P02_W_K': [0] + [float(x.splitlines()[-1].replace(',', '.')) for x in fields.get('P02_W_K', [''])] + [0],            
            'P03_W_K': [0] + [float(x.splitlines()[-1].replace(',', '.')) for x in fields.get('P03_W_K', [''])] + [0],            
            'P04_W_K': [0] + [float(x.splitlines()[-1].replace(',', '.')) for x in fields.get('P04_W_K', [''])] + [0],            
            'P05_W_K': [0] + [float(x.splitlines()[-1].replace(',', '.')) for x in fields.get('P05_W_K', [''])] + [0],            
            'UA_phiL': [float(x.splitlines()[-1].replace(',', '.')) for x in fields.get('UA_phiL', [''])],            

            }      

        return result

    except (RuntimeError, IndexError, ValueError) as e:
        logging.error(f"Error in get_informe_cev_v2_pagina3_envolvente_as_dict: {str(e)}")
        return {}  # Return empty dictionary in case of error

def get_informe_cev_v2_pagina3_envolvente_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """
    Extract data from page 3 (envolvente) of an informe_CEV_v2 PDF report and return it as a Pandas DataFrame.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        pd.DataFrame: A DataFrame containing the extracted data fields from page 3 (envolvente).
                      Returns an empty DataFrame if any error occurs during extraction.
    """
    try:
        # Call the existing function to get the data as a dictionary
        data_dict: Dict[str, Any] = get_informe_cev_v2_pagina3_envolvente_as_dict(pdf_report)
        data_dict['codigo_evaluacion'] = [data_dict['codigo_evaluacion']] * 10

        # Convert the dictionary to a DataFrame
        df: pd.DataFrame = pd.DataFrame.from_dict(data_dict, orient='index').T

        # Drop the 'codigo_evaluacion' column
        if "codigo_evaluacion" in df.columns:
            df.drop(columns=["codigo_evaluacion"], inplace=True)

        return df

    except Exception as e:
        print(f"Error in get_informe_cev_v2_pagina3_envolvente_as_dataframe: {str(e)}")
        return pd.DataFrame()  # Return an empty DataFrame in case of error

def scrape_informe_cev_v2_pagina3_envolvente(pdf_report):
    # ### Load the PDF
    #pdf_report = fitz.open(pdf_file_path)
    page_number = 2  # Page number (starting from 0)
    page = pdf_report[page_number]

    # ## Pagina 3
    # Código evaluación energética
    area_coordinates = (62.3, 30.7, 88.1, 35.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    codigo_evaluacion = extracted_text
    df = pd.DataFrame(data=[codigo_evaluacion] * 10, columns=['codigo_evaluacion'])

    # ### Seccion: Resumen Envolvente

    orientacion = ['Horiz', 'N', 'NE', 'E', 'SE', 'S', 'SO', 'O', 'NO', 'Pisos']
    df['orientacion'] = orientacion

    # ### Elementos Opacos
    area_coordinates = (19.5, 245.0, 47.0, 287.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    elementos_opacos_area_m2 = extracted_text.splitlines()[-10:]
    elementos_opacos_area_m2 = [float(item.replace(',', '.')) for item in elementos_opacos_area_m2]
    df['elementos_opacos_area_m2'] = elementos_opacos_area_m2
    
    area_coordinates = (47.8, 245.0, 60.5, 287.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    elementos_opacos_U_W_m2_K = extracted_text.splitlines()[-10:]
    elementos_opacos_U_W_m2_K = [float(item.replace(',', '.')) for item in elementos_opacos_U_W_m2_K]
    df['elementos_opacos_U_W_m2_K'] = elementos_opacos_U_W_m2_K
    df


    # ### Elementos Traslucidos
    area_coordinates = (68.2, 245.0, 89.5, 287.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    elementos_traslucidos_area_m2 = extracted_text.splitlines()[-9:]
    elementos_traslucidos_area_m2 = [float(item.replace(',', '.')) for item in elementos_traslucidos_area_m2]
    elementos_traslucidos_area_m2.append(0)
    df['elementos_traslucidos_area_m2'] = elementos_traslucidos_area_m2

    area_coordinates = (90.4, 245.0, 103.1, 287.5)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    elementos_traslucidos_U_W_m2_K = extracted_text.splitlines()[-9:]
    elementos_traslucidos_U_W_m2_K = [float(item.replace(',', '.')) for item in elementos_traslucidos_U_W_m2_K]
    elementos_traslucidos_U_W_m2_K.append(0)
    df['elementos_traslucidos_U_W_m2_K'] = elementos_traslucidos_U_W_m2_K
   
    # ### Perdidas Puentes Termicos

    # ### P01
    p01_W_K = []
    dy = 3.5
    for n in range(0, 9):
        area_coordinates = (114.1, 250.0+n*dy, 125.5, 253.5+n*dy)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
        extracted_text = extract_text_from_area(page, area_coordinates)
        p01_W_K_i = extracted_text.splitlines()[-1]
        p01_W_K_i = float(p01_W_K_i.replace(',', '.'))
        p01_W_K.append(p01_W_K_i)

    p01_W_K.pop(4)
    p01_W_K.insert(0, 0)
    p01_W_K.append(0)
    df['P01_W_K'] = p01_W_K

    # ### P02
    p02_W_K = []
    dy = 3.5
    for n in range(0, 9):
        area_coordinates = (126.2, 250.0+n*dy, 136.9, 253.5+n*dy)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
        extracted_text = extract_text_from_area(page, area_coordinates)
        p02_W_K_i = extracted_text.splitlines()[-1]
        p02_W_K_i = float(p02_W_K_i.replace(',', '.'))
        p02_W_K.append(p02_W_K_i)

    p02_W_K.pop(4)
    p02_W_K.insert(0, 0)
    p02_W_K.append(0)
    df['P02_W_K'] = p02_W_K

    # ### P03
    p03_W_K = []
    dy = 3.5
    for n in range(0, 9):
        #print(n)
        area_coordinates = (137.0, 250.0+n*dy, 148.2, 253.5+n*dy)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
        extracted_text = extract_text_from_area(page, area_coordinates)
        p03_W_K_i = extracted_text.splitlines()[-1]
        p03_W_K_i = float(p03_W_K_i.replace(',', '.'))
        p03_W_K.append(p03_W_K_i)

    p03_W_K.pop(4)
    p03_W_K.insert(0, 0)
    p03_W_K.append(0)
    df['P03_W_K'] = p03_W_K
   
    # ### P04
    p04_W_K = []
    dy = 3.5
    for n in range(0, 9):
        #print(n)
        area_coordinates = (149.0, 250.0+n*dy, 160.0, 253.5+n*dy)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
        extracted_text = extract_text_from_area(page, area_coordinates)
        p04_W_K_i = extracted_text.splitlines()[-1]
        p04_W_K_i = float(p04_W_K_i.replace(',', '.'))
        p04_W_K.append(p04_W_K_i)

    p04_W_K.pop(4)  
    p04_W_K.insert(0, 0)
    p04_W_K.append(0)
    df['P04_W_K'] = p04_W_K
  
    # ### P05
    p05_W_K = []
    dy = 3.5
    for n in range(0, 9):
        #print(n)
        area_coordinates = (161.3, 250.0+n*dy, 171.2, 253.5+n*dy)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
        extracted_text = extract_text_from_area(page, area_coordinates)
        p05_W_K_i = extracted_text.splitlines()[-1]
        p05_W_K_i = float(p05_W_K_i.replace(',', '.'))
        p05_W_K.append(p05_W_K_i)

    p05_W_K.pop(4) 
    p05_W_K.insert(0, 0)
    p05_W_K.append(0)
    df['P05_W_K'] = p05_W_K
    
    Ht_W_K = []
    dy = 3.5
    for n in range(0, 12):
        #print(n)
        area_coordinates = (189.2, 245.5+n*dy, 201.9, 249.0+n*dy)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
        extracted_text = extract_text_from_area(page, area_coordinates)
        Ht_W_K_i = extracted_text.splitlines()[-1]
        Ht_W_K_i = float(Ht_W_K_i.replace(',', '.'))
        Ht_W_K.append(Ht_W_K_i)
    Ht_W_K.pop(4) 
    Ht_W_K.pop(9) 
    df['UA_phiL'] = Ht_W_K
 
    # df['elementos_opacos_area_m2'] * df['elementos_opacos_U_W_m2_K'] + df['elementos_traslucidos_area_m2'] * df['elementos_traslucidos_U_W_m2_K'] + df['P01_W_K'] + df['P02_W_K'] + df['P03_W_K'] + df['P04_W_K'] + df['P05_W_K']
   
    # Close the PDF document
    # pdf_report.close()
    return df


# ------------------------------------------------------------------------------------------------------------
#  Pagina 4 
# ------------------------------------------------------------------------------------------------------------  
def get_informe_cev_v2_pagina4_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 4 of an informe_CEV_v2 PDF report and return it as a dictionary.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        Dict[str, Any]: A dictionary containing the extracted data fields from page 4.
                       Returns an empty dictionary if any error occurs during extraction.
    """
    try:
        # Validate input
        if not isinstance(pdf_report, fitz.Document):
            raise ValueError("Invalid input: pdf_report must be a fitz.Document object.")
        if 3 >= pdf_report.page_count:
            raise ValueError("Invalid page number: Page 3 does not exist in the document.")

        page = pdf_report[3]  # Get page 4 (index 3)

        # Define coordinates as constants for page 4
        dx = 13.5
        COORDINATES: Dict[str, Tuple[float, float, float, float]] = {            
                    'codigo_evaluacion': (62.3, 30.7, 88.1, 36.1),
                    'demanda_calef_viv_eval_kwh': [(42.0 + i * dx, 139.5, 53.5 + i * dx, 143.5) for i in range(12)],
                    'demanda_calef_viv_ref_kwh': [(42.0 + i * dx, 144.1, 53.5 + i * dx, 147.8) for i in range(12)],
                    'demanda_enfri_viv_eval_kwh': [(42.0 + i * dx, 161.4, 53.5 + i * dx, 165.4) for i in range(12)],
                    'demanda_enfri_viv_ref_kwh': [(42.0 + i * dx, 166.0, 53.5 + i * dx, 168.8) for i in range(12)],
                    'sobrecalentamiento_viv_eval_hr': [(42.0 + i * dx, 254.9, 53.5 + i * dx, 258.6) for i in range(12)],
                    'sobrecalentamiento_viv_ref_hr': [(42.0 + i * dx, 259.6, 53.5 + i * dx, 263.1) for i in range(12)],
                    'sobreenfriamiento_viv_eval_hr': [(42.0 + i * dx, 275.0, 53.5 + i * dx, 278.6) for i in range(12)],
                    'sobreenfriamiento_viv_ref_hr': [(42.0 + i * dx, 279.4, 53.5 + i * dx, 283.1) for i in range(12)]
                }
       

       # Extract all fields at once
        fields: Dict[str, Union[str, List[str]]] = {
            key: (
                [extract_text_from_area(page, coord) for coord in coords] if isinstance(coords, list) else
                extract_text_from_area(page, coords) if coords else None) for key, coords in COORDINATES.items()
          }
        porcentaje_ahorro = next((int(item) for item in fields.get('porcentaje_ahorro', '').splitlines() if item.replace('-', '').isdigit()), None)

        # Create a dictionary to return
        result: Dict[str, Any] = {            
            'codigo_evaluacion': fields.get('codigo_evaluacion', '').strip(),
            'mes_id': [x for x in range(1, 13)],
            'demanda_calef_viv_eval_kwh': [float(x.replace(',', '.')) if x != '' else None for x in fields.get('demanda_calef_viv_eval_kwh', [''])],
            'demanda_calef_viv_ref_kwh': [float(x.replace(',', '.')) if x != '' else None for x in fields.get('demanda_calef_viv_ref_kwh', [''])],
            'demanda_calef_viv_ref_kwh': [float(x.replace(',', '.')) if x != '' else None for x in fields.get('demanda_calef_viv_ref_kwh', [''])],
            'demanda_enfri_viv_eval_kwh': [float(x.replace(',', '.')) if x != '' else None for x in fields.get('demanda_enfri_viv_eval_kwh', [''])],
            'demanda_enfri_viv_ref_kwh': [float(x.replace(',', '.')) if x != '' else None for x in fields.get('demanda_enfri_viv_ref_kwh', [''])],
            'sobrecalentamiento_viv_eval_hr': [float(x.replace(',', '.')) if x != '' else None for x in fields.get('sobrecalentamiento_viv_eval_hr', [''])],
            'sobrecalentamiento_viv_ref_hr': [float(x.replace(',', '.')) if x != '' else None for x in fields.get('sobrecalentamiento_viv_ref_hr', [''])],
            'sobreenfriamiento_viv_eval_hr': [float(x.replace(',', '.')) if x != '' else None for x in fields.get('sobreenfriamiento_viv_eval_hr', [''])],            
            'sobreenfriamiento_viv_ref_hr': [float(x.replace(',', '.')) if x != '' else None for x in fields.get('sobreenfriamiento_viv_ref_hr', [''])],            
        }

        return result

    except (RuntimeError, IndexError, ValueError) as e:
        logging.error(f"Error in get_informe_cev_v2_pagina4_as_dict: {str(e)}")
        return {}  # Return empty dictionary in case of error

def get_informe_cev_v2_pagina4_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """
    Extract data from page 4 of an informe_CEV_v2 PDF report and return it as a Pandas DataFrame.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        pd.DataFrame: A DataFrame containing the extracted data fields from page 4.
                      Returns an empty DataFrame if any error occurs during extraction.
    """
    try:
        # Call the existing function to get the data as a dictionary
        data_dict: Dict[str, Any] = get_informe_cev_v2_pagina4_as_dict(pdf_report)
        data_dict['codigo_evaluacion'] = [data_dict['codigo_evaluacion']] * 12

        # Convert the dictionary to a DataFrame
        df: pd.DataFrame = pd.DataFrame.from_dict(data_dict, orient='index').T
        
        # Drop the 'codigo_evaluacion' column if it exists
        if "codigo_evaluacion" in df.columns:
            df.drop(columns=["codigo_evaluacion"], inplace=True)

        # Convert 'mes_id' values to month names in Spanish
        mes_mapping = {
            1: "Enero",
            2: "Febrero",
            3: "Marzo",
            4: "Abril",
            5: "Mayo",
            6: "Junio",
            7: "Julio",
            8: "Agosto",
            9: "Septiembre",
            10: "Octubre",
            11: "Noviembre",
            12: "Diciembre"
        }

        if 'mes_id' in df.columns:
            df['mes'] = df['mes_id'].map(mes_mapping)
            df.drop(columns=['mes_id'], inplace=True)

            # Move 'mes' to the beginning of the DataFrame
            cols = ['mes'] + [col for col in df.columns if col != 'mes']
            df = df[cols]

        return df

    except Exception as e:
        print(f"Error in get_informe_cev_v2_pagina4_as_dataframe: {str(e)}")
        return pd.DataFrame()  # Return an empty DataFrame in case of error

def scrape_informe_cev_v2_pagina4(pdf_report):
    # # Informe CEV (v.2) - Page 4
    #pdf_report = fitz.open(pdf_file_path)
    page_number = 3  # Page number (starting from 0)
    page = pdf_report[page_number]


    # ## Pagina 4
    area_coordinates = (62.3, 30.7, 88.1, 35.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    codigo_evaluacion = extracted_text

    # Create DataFrame
    df = pd.DataFrame(data=[codigo_evaluacion] * 12, columns=['codigo_evaluacion'])
    df['mes_id'] = range(1, 13)

    # ### Seccion: Demanda Calefaccion
    # #### Vivienda evaluada
    area_coordinates = (41.7, 139.1, 203.4, 143.9)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_calef_viv_eval_kwh = extracted_text.splitlines()
    demanda_calef_viv_eval_kwh = [float(item.replace(',', '.')) for item in demanda_calef_viv_eval_kwh]
    
    # Store a comment column based on length of demanda_calef_viv_eval_kwh
    if len(demanda_calef_viv_eval_kwh) == 12:
        demanda_calef_viv_eval_comment = 'OK'
    else:
        demanda_calef_viv_eval_comment = 'Check!'

    while len(demanda_calef_viv_eval_kwh) < 12:
        demanda_calef_viv_eval_kwh.append(0)
    df['demanda_calef_viv_eval_kwh'] = demanda_calef_viv_eval_kwh
    df['demanda_calef_viv_eval_comment'] = demanda_calef_viv_eval_comment

    # #### Vivienda de referencia
    area_coordinates = (41.7, 144.1, 203.4, 148.2)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_calef_viv_ref_kwh = extracted_text.splitlines()
    demanda_calef_viv_ref_kwh = [float(item.replace(',', '.')) for item in demanda_calef_viv_ref_kwh]

    # Store a comment column based on length of demanda_calef_viv_ref_kwh
    if len(demanda_calef_viv_ref_kwh) == 12:
        demanda_calef_viv_ref_comment = 'OK'
    else:
        demanda_calef_viv_ref_comment = 'Check!'

    while len(demanda_calef_viv_ref_kwh) < 12:
        demanda_calef_viv_ref_kwh.append(0)
    df['demanda_calef_viv_ref_kwh'] = demanda_calef_viv_ref_kwh
    df['demanda_calef_viv_ref_comment'] = demanda_calef_viv_ref_comment

    # ### Seccion: Demanda Enfriamiento
    # Vivienda Evaluada
    area_coordinates = (41.7, 161.1, 203.4, 165.9)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_enfri_viv_eval_kwh = extracted_text.splitlines()
    demanda_enfri_viv_eval_kwh = [float(item.replace(',', '.')) for item in demanda_enfri_viv_eval_kwh]
    
    # Store a comment column based on length of demanda_enfri_viv_eval_kwh
    if len(demanda_enfri_viv_eval_kwh) == 12:
        demanda_enfri_viv_eval_comment = 'OK'
    else:
        demanda_enfri_viv_eval_comment = 'Check!'

    while len(demanda_enfri_viv_eval_kwh) < 12:
        demanda_enfri_viv_eval_kwh.append(0)
    df['demanda_enfri_viv_eval_kwh'] = demanda_enfri_viv_eval_kwh
    df['demanda_enfri_viv_eval_comment'] = demanda_enfri_viv_eval_comment
    
    # Vivienda Referencia
    area_coordinates = (41.7, 166.2, 203.4, 170.3)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    demanda_enfri_viv_ref_kwh = extracted_text.splitlines()
    
    # Convert each item separately, replacing comma with period
    demanda_enfri_viv_ref_kwh = [float(val.replace(',', '.')) for item in demanda_enfri_viv_ref_kwh for val in item.split() if val.strip()]


    # Store a comment column based on length of demanda_enfri_viv_ref_kwh
    if len(demanda_enfri_viv_ref_kwh) == 12:
        demanda_enfri_viv_ref_comment = 'OK'
    else:
        demanda_enfri_viv_ref_comment = 'Check!'

    while len(demanda_enfri_viv_ref_kwh) < 12:
        demanda_enfri_viv_ref_kwh.append(0)
    df['demanda_enfri_viv_ref_kwh'] = demanda_enfri_viv_ref_kwh
    df['demanda_enfri_viv_ref_comment'] = demanda_enfri_viv_ref_comment

    # ### Seccion: Sobrecalentamiento
    # Vivienda Evaluada
    area_coordinates = (41.7, 254.9, 203.4, 258.9)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    sobrecalentamiento_viv_eval_hr = extracted_text.splitlines()
    sobrecalentamiento_viv_eval_hr = [float(item.replace(',', '.')) for item in sobrecalentamiento_viv_eval_hr]

    # Store a comment column based on length of sobrecalentamiento_viv_eval_hr
    if len(sobrecalentamiento_viv_eval_hr) == 12:
        sobrecalentamiento_viv_eval_comment = 'OK'
    else:
        sobrecalentamiento_viv_eval_comment = 'Check!'

    while len(sobrecalentamiento_viv_eval_hr) < 12:
        sobrecalentamiento_viv_eval_hr.append(0)
    df['sobrecalentamiento_viv_eval_hr'] = sobrecalentamiento_viv_eval_hr
    df['sobrecalentamiento_viv_eval_comment'] = sobrecalentamiento_viv_eval_comment


    # Vivienda Referencia
    area_coordinates = (41.7, 259.0, 203.4, 262.9)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    sobrecalentamiento_viv_ref_hr = extracted_text.splitlines()
    sobrecalentamiento_viv_ref_hr = [float(item.replace(',', '.')) for item in sobrecalentamiento_viv_ref_hr]

    # Store a comment column based on length of sobrecalentamiento_viv_ref_hr
    if len(sobrecalentamiento_viv_ref_hr) == 12:
        sobrecalentamiento_viv_ref_comment = 'OK'
    else:
        sobrecalentamiento_viv_ref_comment = 'Check!'

    while len(sobrecalentamiento_viv_ref_hr) < 12:
        sobrecalentamiento_viv_ref_hr.append(0)
    df['sobrecalentamiento_viv_ref_hr'] = sobrecalentamiento_viv_ref_hr
    df['sobrecalentamiento_viv_ref_comment'] = sobrecalentamiento_viv_ref_comment


    # ### Seccion: Sobreenfriamiento
    # Vivienda Evaluada
    area_coordinates = (41.7, 275.0, 203.4, 278.6)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    sobreenfriamiento_viv_eval_hr = extracted_text.splitlines()
    sobreenfriamiento_viv_eval_hr = [float(item.replace(',', '.')) for item in sobreenfriamiento_viv_eval_hr]

    # Store a comment column based on length of sobreenfriamiento_viv_eval_hr
    if len(sobreenfriamiento_viv_eval_hr) == 12:
        sobreenfriamiento_viv_eval_comment = 'OK'
    else:
        sobreenfriamiento_viv_eval_comment = 'Check!'

    while len(sobreenfriamiento_viv_eval_hr) < 12:
        sobreenfriamiento_viv_eval_hr.append(0)
    df['sobreenfriamiento_viv_eval_hr'] = sobreenfriamiento_viv_eval_hr
    df['sobreenfriamiento_viv_eval_comment'] = sobreenfriamiento_viv_eval_comment
 
    # Vivienda Referencia
    area_coordinates = (41.7, 279.0, 203.4, 282.9)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    sobreenfriamiento_viv_ref_hr = extracted_text.splitlines()
    sobreenfriamiento_viv_ref_hr = [float(item.replace(',', '.')) for item in sobreenfriamiento_viv_ref_hr]

    # Store a comment column based on length of sobreenfriamiento_viv_ref_hr
    if len(sobreenfriamiento_viv_ref_hr) == 12:
        sobreenfriamiento_viv_ref_comment = 'OK'
    else:
        sobreenfriamiento_viv_ref_comment = 'Check!'

    while len(sobreenfriamiento_viv_ref_hr) < 12:
        sobreenfriamiento_viv_ref_hr.append(0)
    df['sobreenfriamiento_viv_ref_hr'] = sobreenfriamiento_viv_ref_hr
    df['sobreenfriamiento_viv_ref_comment'] = sobreenfriamiento_viv_ref_comment
    return df
# ------------------------------------------------------------------------------------------------------------
#  Pagina 5 
# ------------------------------------------------------------------------------------------------------------ 

def get_informe_cev_v2_pagina5_as_dataframe(pdf_report):    
    #pdf_report = fitz.open(pdf_file_path)
    page_number = 4  # Page number (starting from 0)
    page = pdf_report[page_number]

    # ## Pagina 5
    # Código evaluación energética
    area_coordinates = (62.3, 30.7, 88.1, 35.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    codigo_evaluacion = extracted_text
    codigo_evaluacion

    df = pd.DataFrame(data=[codigo_evaluacion], columns=['codigo_evaluacion'])
    df['content'] = None    
    return df
# ------------------------------------------------------------------------------------------------------------
#  Pagina 6
# ------------------------------------------------------------------------------------------------------------ 

def get_informe_cev_v2_pagina6_as_dataframe(pdf_report):
    #pdf_report = fitz.open(pdf_file_path)
    page_number = 5  # Page number (starting from 0)
    page = pdf_report[page_number]

    # ## Pagina 6
    # Código evaluación energética
    area_coordinates = (62.3, 30.7, 88.1, 35.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    codigo_evaluacion = extracted_text
    codigo_evaluacion

    df = pd.DataFrame(data=[codigo_evaluacion], columns=['codigo_evaluacion'])
    df['content'] = None
    
    return df

# ------------------------------------------------------------------------------------------------------------
#  Pagina 7
# ------------------------------------------------------------------------------------------------------------ 

def get_informe_cev_v2_pagina7_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 7 of an informe_CEV_v2 PDF report and return it as a dictionary.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        Dict[str, Any]: A dictionary containing the extracted data fields from page 7.
                       Returns an empty dictionary if any error occurs during extraction.
    """
    try:
        # Validate input
        if not isinstance(pdf_report, fitz.Document):
            raise ValueError("Invalid input: pdf_report must be a fitz.Document object.")
        if 6 >= pdf_report.page_count:
            raise ValueError("Invalid page number: Page 7 does not exist in the document.")
            
        page = pdf_report[6]

        # Define coordinates as constants
        COORDINATES: Dict[str, Tuple[float, float, float, float]] = {
            'codigo_evaluacion': (63.5, 30.9, 84.0, 36.1),
            'mandante_nombre': (27.5, 90.6, 96.0, 94.7),
            'mandante_rut': (27.5, 95.4, 96.0, 99.4),
            'evaluador_nombre': (131.1, 90.6, 205.0, 94.7),
            'evaluador_rut': (131.1, 95.4, 196.7, 99.4),
            'evaluador_rol_minvu': (150.0, 99.9, 166.0, 103.7)	
        }

        # Extract all fields at once
        fields: Dict[str, str] = {
            key: extract_text_from_area(page, coords)
            for key, coords in COORDINATES.items()
        }

        # Create a dictionary to return
        result: Dict[str, Any] = {
            'codigo_evaluacion': fields.get('codigo_evaluacion', '').strip(),
            'mandante_nombre': fields.get('mandante_nombre', '').strip(),
            'mandante_rut': fields.get('mandante_rut', '').strip(),
            'evaluador_nombre': fields.get('evaluador_nombre', '').strip(),
            'evaluador_rut': fields.get('evaluador_rut', '').strip(),
            'evaluador_rol_minvu': fields.get('evaluador_rol_minvu', '').strip()
        }
              
        return result
    except (RuntimeError, IndexError, ValueError) as e:
        logging.error(f"Error in get_informe_cev_v2_pagina7_as_dict: {str(e)}")
        return {}  # Return empty dictionary in case of error


def get_informe_cev_v2_pagina7_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """
    Extract data from page 7 of an informe_CEV_v2 PDF report and return it as a Pandas DataFrame.

    Args:
        pdf_report (fitz.Document): The loaded PyMuPDF Document object representing the PDF report.

    Returns:
        pd.DataFrame: A DataFrame containing the extracted data fields from page 7.
                      Returns an empty DataFrame if any error occurs during extraction.
    """
    try:
        # Call the existing function to get the data as a dictionary
        data_dict: Dict[str, Any] = get_informe_cev_v2_pagina7_as_dict(pdf_report)

        # Convert the dictionary to a DataFrame
        df: pd.DataFrame = pd.DataFrame.from_dict(data_dict, orient='index').T  
        
        return df

    except Exception as e:
        print(f"Error in get_informe_cev_v2_pagina7_as_dataframe: {str(e)}")
        return pd.DataFrame()  # Return an empty DataFrame in case of error
    

def scrape_informe_cev_v2_pagina7(pdf_report):

    #pdf_report = fitz.open(pdf_file_path)
    page_number = 6  # Page number (starting from 0)
    page = pdf_report[page_number]


    # ## Pagina 7
    # Código evaluación energética
    area_coordinates = (62.3, 30.7, 88.1, 35.1)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    codigo_evaluacion = extracted_text

    df = pd.DataFrame(data=[codigo_evaluacion], columns=['codigo_evaluacion'])

    # Mandante
    area_coordinates = (27.3, 90.0, 97.0, 98.0)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    extracted_text = extracted_text.split('\n')
    df['mandante_nombre']  = extracted_text[0]
    df['mandante_rut']  = extracted_text[1]

    # Evaluador Energético
    area_coordinates = (130.4, 90.0, 208.7, 103.7)  # Coordinates of the area to extract text from: (x1, y1, x2, y2)
    extracted_text = extract_text_from_area(page, area_coordinates)
    extracted_text = extracted_text.split('\n')
    df['evaluador_nombre']  = extracted_text[-3]
    df['evaluador_rut']  = extracted_text[-2]
    df['evaluador_rol_minvu']  = extracted_text[-1]
    # Close the PDF document
    # pdf_report.close()
    return df