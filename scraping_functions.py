from bisect import bisect_left
from typing import Dict, Tuple, Any, List, Union, Optional # Added Optional
from functools import lru_cache
import pandas as pd
import fitz  # PyMuPDF
import logging

# Configure logging (ensure it's configured somewhere, e.g., here or in app.py)
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

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
    try:
        rx = (x / report_width) * page_width
        ry = (y / report_height) * page_height
        return rx, ry
    except ZeroDivisionError:
        logging.error("Report width or height cannot be zero for normalization.")
        return 0.0, 0.0

def extract_text_from_area(page: fitz.Page, area: Tuple[float, float, float, float]) -> str:
    """
    Extract text from a specific area of a PDF page. Robust error handling.
    """
    if not isinstance(page, fitz.Page):
        logging.error("Invalid page object provided to extract_text_from_area.")
        # Consider raising TypeError for programming errors
        return "" # Return empty string on error

    if not isinstance(area, tuple) or len(area) != 4:
        logging.error(f"Invalid area format provided: {area}. Must be a tuple of 4 coordinates.")
        return ""

    REPORT_WIDTH = 215.9  # mm
    REPORT_HEIGHT = 330.0  # mm

    try:
        page_rect = page.rect
        if page_rect is None:
             logging.error("Could not get page rectangle.")
             return ""
        width = page_rect.width
        height = page_rect.height

        if width <= 0 or height <= 0:
            logging.error(f"Invalid page dimensions in extract_text_from_area: width={width}, height={height}")
            return ""

        x1, y1, x2, y2 = area
        if not all(isinstance(coord, (int, float)) for coord in area):
             logging.error(f"Coordinates must be numeric in extract_text_from_area: {area}")
             return "" # Or raise ValueError

        if x1 >= x2 or y1 >= y2:
            logging.warning(f"Invalid coordinates provided: {area}. Ensure x1 < x2 and y1 < y2.")
            return ""

        # Normalize coordinates
        rx1, ry1 = normalize_coordinates(x1, y1, REPORT_WIDTH, REPORT_HEIGHT, width, height)
        rx2, ry2 = normalize_coordinates(x2, y2, REPORT_WIDTH, REPORT_HEIGHT, width, height)

        # Ensure normalized coordinates create a valid rectangle
        if rx1 >= rx2 or ry1 >= ry2:
             logging.warning(f"Normalized coordinates resulted in invalid rectangle: ({rx1}, {ry1}, {rx2}, {ry2}) from area {area}")
             return ""

        rect = fitz.Rect(rx1, ry1, rx2, ry2)
        extracted_text = page.get_textbox(rect) # Use get_textbox for better layout preservation than get_text
        return extracted_text.strip() if extracted_text else ""

    except ZeroDivisionError:
        # This might occur if normalize_coordinates fails due to zero report dimensions
        logging.error("Division by zero error during coordinate normalization in extract_text_from_area.")
        return ""
    except Exception as e:
        logging.error(f"Unexpected error extracting text from area {area}: {e}", exc_info=True)
        return ""


# --- Helper Functions ---

def safe_float_convert(text: Optional[str], default: Any = None) -> Union[float, None]:
    """Safely converts a string to a float, handling potential errors and None input."""
    if text is None or text == '':
        return default
    try:
        # Handle potential thousand separators (.) before decimal comma (,) typical in Spanish locale
        cleaned_text = text.replace('.', '').replace(',', '.').strip()
        # Basic check if it looks like a number (allows negative, scientific notation)
        return float(cleaned_text)
    except (ValueError, TypeError):
        logging.warning(f"Could not convert '{text}' to float.")
        return default

def _from_procentaje_ahorro_to_letra(porcentaje_ahorro_decimal: Optional[float]) -> Optional[str]:
    """
    Convert a savings percentage (as a float, e.g., 0.75 for 75%) to a corresponding letter grade.
    Handles None input gracefully.
    """
    if porcentaje_ahorro_decimal is None:
        return None
    # Boundaries defined based on the percentage value (e.g., -0.35 corresponds to -35%)
    boundaries = [-0.35, -0.1, 0.2, 0.4, 0.55, 0.7, 0.85, 100.0] # Using 100.0 for clarity beyond A+
    grades = ['G', 'F', 'E', 'D', 'C', 'B', 'A', 'A+']

    try:
        idx = bisect_left(boundaries, porcentaje_ahorro_decimal)
        if 0 <= idx < len(grades):
            return grades[idx]
        else:
             # Handle cases where percentage is outside the defined boundaries (e.g., > 100 or extremely negative)
             logging.warning(f"Percentage {porcentaje_ahorro_decimal*100}% resulted in out-of-bounds grade index {idx}.")
             if idx >= len(grades): return grades[-1] # Assign highest grade if above top boundary
             else: return grades[0] # Assign lowest grade if below bottom boundary

    except TypeError:
         logging.error(f"Invalid type for percentage: {porcentaje_ahorro_decimal}. Cannot determine grade.")
         return None


# ------------------------------------------------------------------------------------------------------------
#  Pagina 1
# ------------------------------------------------------------------------------------------------------------

def get_informe_cev_v2_pagina1_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 1 of an informe_CEV_v2 PDF report and return it as a dictionary.
    Uses safe float conversion.
    """
    result: Dict[str, Any] = {}
    try:
        if not isinstance(pdf_report, fitz.Document): raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 1: raise ValueError("PDF has no pages.")
        page = pdf_report[0]

        COORDINATES: Dict[str, Tuple[float, float, float, float]] = {
            'tipo_evaluacion': (8.3, 10.3, 165.6, 18.8),
            'codigo_evaluacion': (73.1, 20.0, 95.6, 25.1),
            'region': (28.0, 26.6, 80.0, 31.8),
            'comuna': (29.2, 33.0, 80.0, 38.2),
            'direccion': (31.3, 39.1, 155.3, 44.3),
            'rol_vivienda_proyecto': (57.4, 45.6, 74.8, 50.8),
            'tipo_vivienda':(45.9, 51.7, 155.3, 56.9),
            'superficie_interior_util_m2': (54.2, 58.3, 66.0, 63.5),
            'porcentaje_ahorro_raw': (5.6, 78.6, 165.8, 191.3), # Raw text block
            'demanda_calefaccion_kwh_m2_ano_raw': (15.6, 220.0, 73.0, 230.0), # Raw text block
            'demanda_enfriamiento_kwh_m2_ano_raw': (90.0, 220.0, 151.5, 230.0), # Raw text block
            'demanda_total_kwh_m2_ano_raw': (167.0, 225.0, 209.0, 245.0), # Raw text block
            'emitida_el_raw': (34.5, 247.5, 57.0, 252.8) # Raw text block
        }

        fields: Dict[str, str] = { k: extract_text_from_area(page, v) for k, v in COORDINATES.items() }

        # Post-processing with safe conversion
        porcentaje_ahorro_str = next((line for line in fields.get('porcentaje_ahorro_raw', '').splitlines() if line.replace('-', '').isdigit()), None)
        porcentaje_ahorro_int = int(porcentaje_ahorro_str) if porcentaje_ahorro_str is not None else None
        porcentaje_ahorro_decimal = float(porcentaje_ahorro_int / 100.0) if porcentaje_ahorro_int is not None else None

        demanda_cal_str = fields.get('demanda_calefaccion_kwh_m2_ano_raw', '').splitlines()
        demanda_enf_str = fields.get('demanda_enfriamiento_kwh_m2_ano_raw', '').splitlines()
        demanda_tot_str = fields.get('demanda_total_kwh_m2_ano_raw', '').splitlines()
        emitida_str = fields.get('emitida_el_raw', '').splitlines()

        result = {
            'tipo_evaluacion': fields.get('tipo_evaluacion', '').strip(),
            'codigo_evaluacion': fields.get('codigo_evaluacion', '').strip(),
            'region': fields.get('region', '').strip(),
            'comuna': fields.get('comuna', '').strip(),
            'direccion': fields.get('direccion', '').strip(),
            'rol_vivienda_proyecto': fields.get('rol_vivienda_proyecto', '').strip(),
            'tipo_vivienda': fields.get('tipo_vivienda', '').strip(),
            'superficie_interior_util_m2': safe_float_convert(fields.get('superficie_interior_util_m2')),
            'porcentaje_ahorro': porcentaje_ahorro_int, # Keep as integer %
            'letra_eficiencia_energetica_dem': _from_procentaje_ahorro_to_letra(porcentaje_ahorro_decimal),
            'demanda_calefaccion_kwh_m2_ano': safe_float_convert(demanda_cal_str[-1] if demanda_cal_str else None),
            'demanda_enfriamiento_kwh_m2_ano': safe_float_convert(demanda_enf_str[-1] if demanda_enf_str else None),
            'demanda_total_kwh_m2_ano': safe_float_convert(demanda_tot_str[-1] if demanda_tot_str else None),
            'emitida_el': emitida_str[-1].strip() if emitida_str else None
        }
        return result

    except (IndexError, ValueError, TypeError) as e:
        logging.error(f"Error processing Page 1 dictionary: {e}", exc_info=True)
        return {}

def get_informe_cev_v2_pagina1_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts data from page 1 into a Pandas DataFrame."""
    data_dict = get_informe_cev_v2_pagina1_as_dict(pdf_report)
    if not data_dict: return pd.DataFrame()
    try:
        return pd.DataFrame.from_dict(data_dict, orient='index').T
    except Exception as e:
         logging.error(f"Failed to convert page 1 dict to DataFrame: {e}", exc_info=True)
         return pd.DataFrame()


# ------------------------------------------------------------------------------------------------------------
#  Pagina 2
# ------------------------------------------------------------------------------------------------------------

def get_informe_cev_v2_pagina2_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 2 of an informe_CEV_v2 PDF report and return it as a dictionary.
    Uses safe float conversion.
    """
    result: Dict[str, Any] = {}
    try:
        if not isinstance(pdf_report, fitz.Document): raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 2: raise ValueError("PDF has less than 2 pages.")
        page = pdf_report[1]

        COORDINATES: Dict[str, Tuple[float, float, float, float]] = {
            'region': (40.4,47.4,95.0,51.7), 'comuna': (40.4,53.2,95.0,57.4), 'direccion': (40.4,58.9,95.0,63.1), 'rol_vivienda': (40.4,64.6,95.0,68.9), 'tipo_vivienda': (40.4,70.2,95.0,74.4),
            'zona_termica': (143.8,47.5,146.1,51.7), 'superficie_interior_util_m2_raw': (143.8,53.3,150,57.5), 'solicitado_por': (143.8,58.9,185.5,63.1), 'evaluado_por': (143.8,64.7,210.5,68.9), 'codigo_evaluacion': (143.8,70.2,160.6,74.5),
            'demanda_calefaccion_kwh_m2_ano_raw': (99.4,99.6,107.6,105.2), 'demanda_enfriamiento_kwh_m2_ano_raw': (99.2,120.9,107.5,126.5), 'demanda_total_kwh_m2_ano_raw': (101.5,135.1,130.6,151.4),
            'demanda_total_bis_kwh_m2_ano_raw': (39.2, 159.8, 122.8, 166.0), 'demanda_total_referencia_kwh_m2_ano_raw': (16.9, 168.3, 146.2, 173.2), 'porcentaje_ahorro_raw': (152.0, 162.6, 201.5, 168.7),
            'muro_principal_descripcion': (46.2, 202.2, 184.5, 209.1), 'muro_principal_exigencia_raw': (185.5, 204.2, 209.5, 209.1),
            'muro_secundario_descripcion': (46.2, 209.2, 184.5, 215.1), 'muro_secundario_exigencia_raw': (185.5,211.5,209.5,214.3),
            'piso_principal_descripcion': (46.7,216.6,184.5,219.4), 'piso_principal_exigencia_raw': (185.5, 216.4, 209.5, 223.7),
            'puerta_principal_descripcion': (46.2, 223.5, 184.5, 230.2), 'puerta_principal_exigencia_raw': (185.5, 223.9, 209.5, 230.2), # Textual
            'techo_principal_descripcion': (46.2, 230.5, 184.5, 237.0), 'techo_principal_exigencia_raw': (185.5, 230.5, 209.5, 237.2),
            'techo_secundario_descripcion': (46.2, 237.2, 184.5, 244.1), 'techo_secundario_exigencia_raw': (185.5, 237.2, 209.5, 244.1),
            'superficie_vidriada_principal_descripcion': (46.2, 244.2, 184.5, 251.0), 'superficie_vidriada_principal_exigencia': (185.5, 244.3, 209.5, 251.0), # Textual
            'superficie_vidriada_secundaria_descripcion': (46.2, 251.3, 184.5, 258.0), 'superficie_vidriada_secundaria_exigencia': (185.5, 251.3, 209.5, 258.0), # Textual
            'ventilacion_rah_descripcion': (46.2, 258.3, 184.5, 265.0), 'ventilacion_rah_exigencia': (185.5, 258.3, 209.5, 265.0), # Textual
            'infiltraciones_rah_descripcion': (46.2, 265.3, 184.5, 272.0), 'infiltraciones_rah_exigencia': (185.5, 265.3, 209.5, 272.0) # Textual
        }

        fields: Dict[str, str] = { k: extract_text_from_area(page, v) for k, v in COORDINATES.items() }

        # Helper lambdas for cleaner processing
        get_last_line = lambda key: fields.get(key, '').splitlines()[-1].strip() if fields.get(key) else None
        get_last_line_float = lambda key: safe_float_convert(get_last_line(key))
        clean_desc = lambda key: fields.get(key, '').replace('\n', ' ').strip()
        clean_exigencia_float = lambda key: safe_float_convert(fields.get(key, '').replace('[W/m2K]', '').strip())

        result = {
            'region': clean_desc('region'), 'comuna': clean_desc('comuna'), 'direccion': clean_desc('direccion'), 'rol_vivienda': clean_desc('rol_vivienda'), 'tipo_vivienda': clean_desc('tipo_vivienda'),
            'zona_termica': clean_desc('zona_termica'), 'superficie_interior_util_m2': safe_float_convert(fields.get('superficie_interior_util_m2_raw')), 'solicitado_por': clean_desc('solicitado_por'), 'evaluado_por': clean_desc('evaluado_por'), 'codigo_evaluacion': clean_desc('codigo_evaluacion'),
            'demanda_calefaccion_kwh_m2_ano': get_last_line_float('demanda_calefaccion_kwh_m2_ano_raw'), 'demanda_enfriamiento_kwh_m2_ano': get_last_line_float('demanda_enfriamiento_kwh_m2_ano_raw'), 'demanda_total_kwh_m2_ano': get_last_line_float('demanda_total_kwh_m2_ano_raw'),
            'demanda_total_bis_kwh_m2_ano': get_last_line_float('demanda_total_bis_kwh_m2_ano_raw'), 'demanda_total_referencia_kwh_m2_ano': get_last_line_float('demanda_total_referencia_kwh_m2_ano_raw'), 'porcentaje_ahorro': get_last_line_float('porcentaje_ahorro_raw'),
            'muro_principal_descripcion': clean_desc('muro_principal_descripcion'), 'muro_principal_exigencia_W_m2_K': clean_exigencia_float('muro_principal_exigencia_raw'),
            'muro_secundario_descripcion': clean_desc('muro_secundario_descripcion'), 'muro_secundario_exigencia_W_m2_K': clean_exigencia_float('muro_secundario_exigencia_raw'),
            'piso_principal_descripcion': clean_desc('piso_principal_descripcion'), 'piso_principal_exigencia_W_m2_K': clean_exigencia_float('piso_principal_exigencia_raw'),
            'puerta_principal_descripcion': clean_desc('puerta_principal_descripcion'), 'puerta_principal_exigencia': clean_desc('puerta_principal_exigencia_raw'), # Textual
            'techo_principal_descripcion': clean_desc('techo_principal_descripcion'), 'techo_principal_exigencia_W_m2_K': clean_exigencia_float('techo_principal_exigencia_raw'),
            'techo_secundario_descripcion': clean_desc('techo_secundario_descripcion'), 'techo_secundario_exigencia_W_m2_K': clean_exigencia_float('techo_secundario_exigencia_raw'),
            'superficie_vidriada_principal_descripcion': clean_desc('superficie_vidriada_principal_descripcion'), 'superficie_vidriada_principal_exigencia': clean_desc('superficie_vidriada_principal_exigencia'), # Textual
            'superficie_vidriada_secundaria_descripcion': clean_desc('superficie_vidriada_secundaria_descripcion'), 'superficie_vidriada_secundaria_exigencia': clean_desc('superficie_vidriada_secundaria_exigencia'), # Textual
            'ventilacion_rah_descripcion': clean_desc('ventilacion_rah_descripcion'), 'ventilacion_rah_exigencia': clean_desc('ventilacion_rah_exigencia'), # Textual
            'infiltraciones_rah_descripcion': clean_desc('infiltraciones_rah_descripcion'), 'infiltraciones_rah_exigencia': clean_desc('infiltraciones_rah_exigencia') # Textual
        }
        return result

    except (IndexError, ValueError, TypeError) as e:
        logging.error(f"Error processing Page 2 dictionary: {e}", exc_info=True)
        return {}

def get_informe_cev_v2_pagina2_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts data from page 2 into a Pandas DataFrame."""
    data_dict = get_informe_cev_v2_pagina2_as_dict(pdf_report)
    if not data_dict: return pd.DataFrame()
    try:
        return pd.DataFrame.from_dict(data_dict, orient='index').T
    except Exception as e:
        logging.error(f"Failed to convert page 2 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()


# ------------------------------------------------------------------------------------------------------------
#  Pagina 3 - Consumos
# ------------------------------------------------------------------------------------------------------------

def get_informe_cev_v2_pagina3_consumos_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 3 (consumos) of an informe_CEV_v2 PDF report and return it as a dictionary.
    Uses safe float conversion.
    """
    result: Dict[str, Any] = {}
    try:
        if not isinstance(pdf_report, fitz.Document): raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 3: raise ValueError("PDF has less than 3 pages.")
        page = pdf_report[2]

        COORDINATES: Dict[str, Tuple[float, float, float, float]] = {
            'codigo_evaluacion': (62.3, 30.7, 88.1, 35.1),
            'agua_caliente_sanitaria_kwh_m2_raw': (78.1, 73.9, 98.0, 76.7), 'agua_caliente_sanitaria_perc_raw': (98.7, 73.9, 116.3, 76.7),
            'iluminacion_kwh_m2_raw': (79.2, 78.1, 98.3, 81.4), 'iluminacion_per_raw': (98.7, 78.1, 116.3, 81.4),
            'calefaccion_kwh_m2_raw': (79.2, 82.2, 98.3, 86.6), 'calefaccion_kwh_per_raw': (98.7, 82.2, 116.3, 86.6),
            'energia_renovable_no_convencional_kwh_m2_raw': (79.2, 87.2, 98.3, 91.0), 'energia_renovable_no_convencional_per_raw': (98.7, 87.2, 116.3, 91.0),
            'consumo_total_kwh_m2_raw': (118.0, 74.0, 148.0, 86.0), 'emisiones_kgco2_m2_ano_raw': (171.5, 69.0, 183.5, 74.2),
            'calefaccion_descripcion_proy': (76.6, 101.4, 155.5, 105.3), 'calefaccion_consumo_proy_kwh_raw': (157.0, 101.4, 196.0, 105.3), 'calefaccion_consumo_proy_per_raw': (198.0, 101.4, 207.0, 105.3),
            'iluminacion_descripcion_proy': (76.6, 106.2, 155.5, 110.0), 'iluminacion_consumo_proy_kwh_raw': (157.0, 106.2, 196.0, 110.0), 'iluminacion_consumo_proy_per_raw': (198.0, 106.2, 207.0, 110.0),
            'agua_caliente_sanitaria_descripcion_proy': (76.6, 111.2, 155.5, 115.0), 'agua_caliente_sanitaria_consumo_proy_kwh_raw': (157.0, 111.2, 196.0, 115.0), 'agua_caliente_sanitaria_consumo_proy_per_raw': (198.0, 111.2, 207.0, 115.0),
            'energia_renovable_no_convencional_descripcion_proy': (76.6, 115.8, 155.5, 120.0), 'energia_renovable_no_convencional_consumo_proy_kwh_raw': (157.0, 115.8, 196.0, 120.0), 'energia_renovable_no_convencional_consumo_proy_per_raw': (198.0, 115.8, 207.0, 120.0),
            'consumo_total_requerido_proy_kwh_raw': (157.0, 121.0, 196.0, 125.0),
            'calefaccion_descripcion_ref': (76.6, 136.1, 155.5, 140.1), 'calefaccion_consumo_ref_kwh_raw': (157.0, 136.1, 196.0, 140.1), 'calefaccion_consumo_ref_per_raw': (198.0, 136.1, 207.0, 140.1),
            'iluminacion_descripcion_ref': (76.6, 140.7, 155.5, 144.7), 'iluminacion_consumo_ref_kwh_raw': (157.0, 140.7, 196.0, 144.7), 'iluminacion_consumo_ref_per_raw': (198.0, 140.7, 207.0, 144.7),
            'agua_caliente_sanitaria_descripcion_ref': (76.6, 145.2, 155.5, 149.2), 'agua_caliente_sanitaria_consumo_ref_kwh_raw': (157.0, 145.2, 196.0, 149.2), 'agua_caliente_sanitaria_consumo_ref_per_raw': (198.0, 145.2, 207.0, 149.2),
            'energia_renovable_no_convencional_descripcion_ref': (76.6, 150.8, 155.5, 154.8), 'energia_renovable_no_convencional_consumo_ref_kwh_raw': (157.0, 150.8, 196.0, 154.8), 'energia_renovable_no_convencional_consumo_ref_per_raw': (198.0, 150.8, 207.0, 154.8),
            'consumo_total_requerido_ref_kwh_raw': (157.0, 156.0, 196.0, 160.0),
            'consumo_ep_calefaccion_kwh_raw': (87.0, 176.0, 104.0, 179.0), 'consumo_ep_agua_caliente_sanitaria_kwh_raw': (87.0, 180.0, 104.0, 183.5), 'consumo_ep_iluminacion_kwh_raw': (87.0, 184.0, 104.0, 187.5), 'consumo_ep_ventiladores_kwh_raw': (87.0, 188.0, 104.0, 191.5),
            'generacion_ep_fotovoltaicos_kwh_raw': (87.0, 199.0, 104.0, 202.5), 'aporte_fotovoltaicos_consumos_basicos_kwh_raw': (87.0, 203.2, 104.0, 206.0), 'diferencia_fotovoltaica_para_consumo_kwh_raw': (87.0, 206.9, 104.0, 210.2),
            'aporte_solar_termica_consumos_basicos_kwh_raw': (87.0, 218.0, 104.0, 221.0), 'aporte_solar_termica_agua_caliente_sanitaria_kwh_raw': (87.0, 222.5, 104.0, 225.5),
            'total_consumo_ep_antes_fotovoltaica_kwh_raw': (192.0, 176.0, 208.0, 179.5), 'aporte_fotovoltaicos_consumos_basicos_kwh_bis_raw': (192.0, 180.0, 208.0, 183.5), 'consumos_basicos_a_suplir_kwh_raw': (192.0, 184.3, 208.0, 187.0),
            'consumo_total_ep_obj_kwh_raw': (192.0, 199.0, 208.0, 202.5), 'consumo_total_ep_ref_kwh_raw': (192.0, 202.8, 208.0, 206.5), 'coeficiente_energetico_c_raw': (192.0, 207.0, 208.0, 210.5)
        }

        fields: Dict[str, str] = { k: extract_text_from_area(page, v) for k, v in COORDINATES.items() }

        get_float = lambda key: safe_float_convert(fields.get(key))
        get_last_line_float = lambda key: safe_float_convert(fields.get(key, '').splitlines()[-1] if fields.get(key) else None)
        clean_desc = lambda key: fields.get(key, '').replace('\n', ' ').strip()

        result = {
            'codigo_evaluacion': clean_desc('codigo_evaluacion'),
            'agua_caliente_sanitaria_kwh_m2': get_float('agua_caliente_sanitaria_kwh_m2_raw'), 'agua_caliente_sanitaria_perc': get_float('agua_caliente_sanitaria_perc_raw'),
            'iluminacion_kwh_m2': get_float('iluminacion_kwh_m2_raw'), 'iluminacion_per': get_float('iluminacion_per_raw'),
            'calefaccion_kwh_m2': get_float('calefaccion_kwh_m2_raw'), 'calefaccion_kwh_per': get_float('calefaccion_kwh_per_raw'),
            'energia_renovable_no_convencional_kwh_m2': get_float('energia_renovable_no_convencional_kwh_m2_raw'), 'energia_renovable_no_convencional_per': get_float('energia_renovable_no_convencional_per_raw'),
            'consumo_total_kwh_m2': get_float('consumo_total_kwh_m2_raw'), 'emisiones_kgco2_m2_ano': get_float('emisiones_kgco2_m2_ano_raw'),
            'calefaccion_descripcion_proy': clean_desc('calefaccion_descripcion_proy'), 'calefaccion_consumo_proy_kwh': get_last_line_float('calefaccion_consumo_proy_kwh_raw'), 'calefaccion_consumo_proy_per': get_last_line_float('calefaccion_consumo_proy_per_raw'),
            'iluminacion_descripcion_proy': clean_desc('iluminacion_descripcion_proy'), 'iluminacion_consumo_proy_kwh': get_last_line_float('iluminacion_consumo_proy_kwh_raw'), 'iluminacion_consumo_proy_per': get_last_line_float('iluminacion_consumo_proy_per_raw'),
            'agua_caliente_sanitaria_descripcion_proy': clean_desc('agua_caliente_sanitaria_descripcion_proy'), 'agua_caliente_sanitaria_consumo_proy_kwh': get_last_line_float('agua_caliente_sanitaria_consumo_proy_kwh_raw'), 'agua_caliente_sanitaria_consumo_proy_per': get_last_line_float('agua_caliente_sanitaria_consumo_proy_per_raw'),
            'energia_renovable_no_convencional_descripcion_proy': clean_desc('energia_renovable_no_convencional_descripcion_proy'), 'energia_renovable_no_convencional_consumo_proy_kwh': get_last_line_float('energia_renovable_no_convencional_consumo_proy_kwh_raw'), 'energia_renovable_no_convencional_consumo_proy_per': get_last_line_float('energia_renovable_no_convencional_consumo_proy_per_raw'),
            'consumo_total_requerido_proy_kwh': get_last_line_float('consumo_total_requerido_proy_kwh_raw'),
            'calefaccion_descripcion_ref': clean_desc('calefaccion_descripcion_ref'), 'calefaccion_consumo_ref_kwh': get_last_line_float('calefaccion_consumo_ref_kwh_raw'), 'calefaccion_consumo_ref_per': get_last_line_float('calefaccion_consumo_ref_per_raw'),
            'iluminacion_descripcion_ref': clean_desc('iluminacion_descripcion_ref'), 'iluminacion_consumo_ref_kwh': get_last_line_float('iluminacion_consumo_ref_kwh_raw'), 'iluminacion_consumo_ref_per': get_last_line_float('iluminacion_consumo_ref_per_raw'),
            'agua_caliente_sanitaria_descripcion_ref': clean_desc('agua_caliente_sanitaria_descripcion_ref'), 'agua_caliente_sanitaria_consumo_ref_kwh': get_last_line_float('agua_caliente_sanitaria_consumo_ref_kwh_raw'), 'agua_caliente_sanitaria_consumo_ref_per': get_last_line_float('agua_caliente_sanitaria_consumo_ref_per_raw'),
            'energia_renovable_no_convencional_descripcion_ref': clean_desc('energia_renovable_no_convencional_descripcion_ref'), 'energia_renovable_no_convencional_consumo_ref_kwh': get_last_line_float('energia_renovable_no_convencional_consumo_ref_kwh_raw'), 'energia_renovable_no_convencional_consumo_ref_per': get_last_line_float('energia_renovable_no_convencional_consumo_ref_per_raw'),
            'consumo_total_requerido_ref_kwh': get_last_line_float('consumo_total_requerido_ref_kwh_raw'),
            'consumo_ep_calefaccion_kwh': get_float('consumo_ep_calefaccion_kwh_raw'), 'consumo_ep_agua_caliente_sanitaria_kwh': get_float('consumo_ep_agua_caliente_sanitaria_kwh_raw'), 'consumo_ep_iluminacion_kwh': get_float('consumo_ep_iluminacion_kwh_raw'), 'consumo_ep_ventiladores_kwh': get_float('consumo_ep_ventiladores_kwh_raw'),
            'generacion_ep_fotovoltaicos_kwh': get_float('generacion_ep_fotovoltaicos_kwh_raw'), 'aporte_fotovoltaicos_consumos_basicos_kwh': get_float('aporte_fotovoltaicos_consumos_basicos_kwh_raw'), 'diferencia_fotovoltaica_para_consumo_kwh': get_float('diferencia_fotovoltaica_para_consumo_kwh_raw'),
            'aporte_solar_termica_consumos_basicos_kwh': get_float('aporte_solar_termica_consumos_basicos_kwh_raw'), 'aporte_solar_termica_agua_caliente_sanitaria_kwh': get_float('aporte_solar_termica_agua_caliente_sanitaria_kwh_raw'),
            'total_consumo_ep_antes_fotovoltaica_kwh': get_float('total_consumo_ep_antes_fotovoltaica_kwh_raw'), 'aporte_fotovoltaicos_consumos_basicos_kwh_bis': get_float('aporte_fotovoltaicos_consumos_basicos_kwh_bis_raw'), 'consumos_basicos_a_suplir_kwh': get_float('consumos_basicos_a_suplir_kwh_raw'),
            'consumo_total_ep_obj_kwh': get_float('consumo_total_ep_obj_kwh_raw'), 'consumo_total_ep_ref_kwh': get_float('consumo_total_ep_ref_kwh_raw'), 'coeficiente_energetico_c': get_float('coeficiente_energetico_c_raw')
        }
        return result

    except (IndexError, ValueError, TypeError) as e:
        logging.error(f"Error processing Page 3 (Consumos) dictionary: {e}", exc_info=True)
        return {}

def get_informe_cev_v2_pagina3_consumos_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts consumption data from page 3 into a Pandas DataFrame."""
    data_dict = get_informe_cev_v2_pagina3_consumos_as_dict(pdf_report)
    if not data_dict: return pd.DataFrame()
    try:
        return pd.DataFrame.from_dict(data_dict, orient='index').T
    except Exception as e:
        logging.error(f"Failed to convert page 3 consumos dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()


# ------------------------------------------------------------------------------------------------------------
#  Pagina 3 - Envolvente
# ------------------------------------------------------------------------------------------------------------

def get_informe_cev_v2_pagina3_envolvente_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extracts envelope data from page 3 into a dictionary (structured for DataFrame).
    Uses safe float conversion.
    """
    data_list: Dict[str, List[Any]] = {}
    try:
        if not isinstance(pdf_report, fitz.Document): raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 3: raise ValueError("PDF has less than 3 pages.")
        page = pdf_report[2]

        dy = 4.2; num_orientations = 10; num_puentes_termicos = 8; puente_termico_start_y = 250.0
        orientations = ['Horiz', 'N', 'NE', 'E', 'SE', 'S', 'SO', 'O', 'NO', 'Pisos']

        COORDINATES_BLOCKS: Dict[str, Tuple[float, float, float, float]] = {
            'codigo_eval_coords': (62.3, 30.7, 88.1, 35.1),
            'opacos_area_coords': (19.5, 245.0, 47.0, 245.0 + (num_orientations * dy)),
            'opacos_U_coords': (47.8, 245.0, 60.5, 245.0 + (num_orientations * dy)),
            'traslucidos_area_coords': (68.2, 245.0, 89.5, 245.0 + ((num_orientations -1) * dy)),
            'traslucidos_U_coords': (90.4, 245.0, 103.1, 245.0 + ((num_orientations -1) * dy)),
            'ua_phiL_coords': (190.5, 245.5, 201.0, 245.5 + (num_orientations * dy))
        }

        PT_COORDS_BASE: Dict[str, Tuple[float, float]] = {
            'P01_W_K': (115.5, 124.5), 'P02_W_K': (126.2, 136.9), 'P03_W_K': (139.0, 148.2),
            'P04_W_K': (149.0, 160.0), 'P05_W_K': (161.3, 171.2)
        }

        # --- Extract Single Value ---
        codigo_evaluacion = extract_text_from_area(page, COORDINATES_BLOCKS['codigo_eval_coords']).strip()

        # --- Extract Columnar Data Blocks ---
        opacos_area_text = extract_text_from_area(page, COORDINATES_BLOCKS['opacos_area_coords'])
        opacos_U_text = extract_text_from_area(page, COORDINATES_BLOCKS['opacos_U_coords'])
        traslucidos_area_text = extract_text_from_area(page, COORDINATES_BLOCKS['traslucidos_area_coords'])
        traslucidos_U_text = extract_text_from_area(page, COORDINATES_BLOCKS['traslucidos_U_coords'])
        ua_phiL_text = extract_text_from_area(page, COORDINATES_BLOCKS['ua_phiL_coords'])

        # --- Extract Puente Termico Data ---
        puentes_termicos_text: Dict[str, List[str]] = {key: [] for key in PT_COORDS_BASE}
        for key, (x1, x2) in PT_COORDS_BASE.items():
            for i in range(num_puentes_termicos):
                y1 = puente_termico_start_y + i * dy; y2 = y1 + 3.5
                pt_coord = (x1, y1, x2, y2)
                text_lines = extract_text_from_area(page, pt_coord).splitlines()
                puentes_termicos_text[key].append(text_lines[-1] if text_lines else '')

        # --- Process and Structure Data ---
        data_list['codigo_evaluacion'] = [codigo_evaluacion] * num_orientations
        data_list['orientacion'] = orientations

        opacos_area_lines = opacos_area_text.splitlines()[-num_orientations:]
        opacos_U_lines = opacos_U_text.splitlines()[-num_orientations:]
        data_list['elementos_opacos_area_m2'] = [safe_float_convert(line) for line in opacos_area_lines]
        data_list['elementos_opacos_U_W_m2_K'] = [safe_float_convert(line) for line in opacos_U_lines]

        traslucidos_area_lines = traslucidos_area_text.splitlines()[-(num_orientations-1):]
        traslucidos_U_lines = traslucidos_U_text.splitlines()[-(num_orientations-1):]
        data_list['elementos_traslucidos_area_m2'] = [safe_float_convert(line) for line in traslucidos_area_lines] + [None]
        data_list['elementos_traslucidos_U_W_m2_K'] = [safe_float_convert(line) for line in traslucidos_U_lines] + [None]

        for key, lines in puentes_termicos_text.items():
            float_values = [safe_float_convert(line) for line in lines]
            data_list[key] = [None] + float_values + [None] # Pad first and last

        ua_phiL_lines = ua_phiL_text.splitlines()[-num_orientations:]
        data_list['UA_phiL'] = [safe_float_convert(line) for line in ua_phiL_lines]

        # Validate list lengths
        for key, lst in data_list.items():
            if len(lst) != num_orientations:
                logging.warning(f"Length mismatch for {key} (Envolvente): expected {num_orientations}, got {len(lst)}. Padding.")
                data_list[key].extend([None] * (num_orientations - len(lst)))

        return data_list

    except (IndexError, ValueError, TypeError) as e:
        logging.error(f"Error processing Page 3 (Envolvente) dictionary: {e}", exc_info=True)
        return {}

def get_informe_cev_v2_pagina3_envolvente_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts envelope data from page 3 into a Pandas DataFrame."""
    data_dict_of_lists = get_informe_cev_v2_pagina3_envolvente_as_dict(pdf_report)
    if not data_dict_of_lists: return pd.DataFrame()
    try:
        # Directly create DataFrame from the dictionary of lists
        df = pd.DataFrame(data_dict_of_lists)
        # Drop the duplicated codigo_evaluacion column generated during dict creation if exists
        if "codigo_evaluacion" in df.columns:
            df = df.drop(columns=["codigo_evaluacion"])
        return df
    except ValueError as ve:
         logging.error(f"ValueError creating DataFrame for Page 3 Envolvente (likely unequal list lengths): {ve}", exc_info=True)
         return pd.DataFrame()
    except Exception as e:
        logging.error(f"Failed to convert page 3 envolvente dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()


# ------------------------------------------------------------------------------------------------------------
#  Pagina 4
# ------------------------------------------------------------------------------------------------------------

def get_informe_cev_v2_pagina4_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extracts monthly data from page 4 into a dictionary (structured for DataFrame).
    Uses safe float conversion.
    """
    data_list: Dict[str, List[Any]] = {}
    try:
        if not isinstance(pdf_report, fitz.Document): raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 4: raise ValueError("PDF has less than 4 pages.")
        page = pdf_report[3]
        num_months = 12; months = list(range(1, num_months + 1))

        codigo_eval_coords = (62.3, 30.7, 88.1, 36.1)
        dx = 13.5; base_x = 42.0; col_width = 11.5

        Y_COORDS: Dict[str, Tuple[float, float]] = {
            'demanda_calef_viv_eval_kwh': (139.5, 143.5), 'demanda_calef_viv_ref_kwh': (144.1, 147.8),
            'demanda_enfri_viv_eval_kwh': (161.4, 165.4), 'demanda_enfri_viv_ref_kwh': (166.0, 168.8),
            'sobrecalentamiento_viv_eval_hr': (254.9, 258.6), 'sobrecalentamiento_viv_ref_hr': (259.6, 263.1),
            'sobreenfriamiento_viv_eval_hr': (275.0, 278.6), 'sobreenfriamiento_viv_ref_hr': (279.4, 283.1)
        }

        codigo_evaluacion = extract_text_from_area(page, codigo_eval_coords).strip()
        data_list['codigo_evaluacion'] = [codigo_evaluacion] * num_months
        data_list['mes_id'] = months

        for key, (y1, y2) in Y_COORDS.items():
            monthly_values_text: List[str] = []
            for i in range(num_months):
                x1 = base_x + i * dx; x2 = x1 + col_width
                month_coord = (x1, y1, x2, y2)
                text = extract_text_from_area(page, month_coord)
                monthly_values_text.append(text)
            data_list[key] = [safe_float_convert(val) for val in monthly_values_text]

        # Validate list lengths
        for key, lst in data_list.items():
            if len(lst) != num_months:
                logging.warning(f"Length mismatch for {key} (Page 4): expected {num_months}, got {len(lst)}. Padding.")
                data_list[key].extend([None] * (num_months - len(lst)))

        return data_list

    except (IndexError, ValueError, TypeError) as e:
        logging.error(f"Error processing Page 4 dictionary: {e}", exc_info=True)
        return {}

def get_informe_cev_v2_pagina4_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts monthly data from page 4 into a Pandas DataFrame."""
    data_dict_of_lists = get_informe_cev_v2_pagina4_as_dict(pdf_report)
    if not data_dict_of_lists: return pd.DataFrame()
    try:
        df = pd.DataFrame(data_dict_of_lists)
        # Drop the duplicated codigo_evaluacion column
        if "codigo_evaluacion" in df.columns:
            df = df.drop(columns=["codigo_evaluacion"])

        # Convert mes_id to Spanish month names
        mes_mapping = { 1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre" }
        if 'mes_id' in df.columns:
            df['mes'] = df['mes_id'].map(mes_mapping)
            df = df.drop(columns=['mes_id'])
            # Move 'mes' column to the beginning
            cols = ['mes'] + [col for col in df.columns if col != 'mes']
            df = df[cols]
        return df
    except ValueError as ve:
         logging.error(f"ValueError creating DataFrame for Page 4 (likely unequal list lengths): {ve}", exc_info=True)
         return pd.DataFrame()
    except Exception as e:
        logging.error(f"Failed to convert page 4 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()


# ------------------------------------------------------------------------------------------------------------
#  Pagina 5
# ------------------------------------------------------------------------------------------------------------
def get_informe_cev_v2_pagina5_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """Creates placeholder dict for page 5."""
    result: Dict[str, Any] = {}
    try:
        if not isinstance(pdf_report, fitz.Document): raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 5: raise ValueError("PDF has less than 5 pages.")
        page = pdf_report[4]
        codigo_eval_coords = (62.3, 30.7, 88.1, 35.1) # Standard coord
        codigo_evaluacion = extract_text_from_area(page, codigo_eval_coords).strip()
        result = {
            'codigo_evaluacion': codigo_evaluacion,
            'content_note': 'Extracción de datos específicos para Página 5 no implementada (contenido gráfico).'
        }
        return result
    except Exception as e:
        logging.error(f"Error accessing Page 5 dictionary: {e}", exc_info=True)
        return {}

def get_informe_cev_v2_pagina5_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Creates placeholder DataFrame for page 5."""
    data_dict = get_informe_cev_v2_pagina5_as_dict(pdf_report)
    if not data_dict: return pd.DataFrame()
    try:
        # Need to wrap single values in lists for DataFrame creation
        return pd.DataFrame({k: [v] for k, v in data_dict.items()})
    except Exception as e:
        logging.error(f"Failed to convert page 5 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()

# ------------------------------------------------------------------------------------------------------------
#  Pagina 6
# ------------------------------------------------------------------------------------------------------------
def get_informe_cev_v2_pagina6_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """Creates placeholder dict for page 6."""
    result: Dict[str, Any] = {}
    try:
        if not isinstance(pdf_report, fitz.Document): raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 6: raise ValueError("PDF has less than 6 pages.")
        page = pdf_report[5]
        codigo_eval_coords = (62.3, 30.7, 88.1, 35.1) # Standard coord
        codigo_evaluacion = extract_text_from_area(page, codigo_eval_coords).strip()
        result = {
            'codigo_evaluacion': codigo_evaluacion,
            'content_note': 'Extracción de datos específicos para Página 6 no implementada (contenido gráfico).'
        }
        return result
    except Exception as e:
        logging.error(f"Error accessing Page 6 dictionary: {e}", exc_info=True)
        return {}

def get_informe_cev_v2_pagina6_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Creates placeholder DataFrame for page 6."""
    data_dict = get_informe_cev_v2_pagina6_as_dict(pdf_report)
    if not data_dict: return pd.DataFrame()
    try:
        return pd.DataFrame({k: [v] for k, v in data_dict.items()})
    except Exception as e:
        logging.error(f"Failed to convert page 6 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()

# ------------------------------------------------------------------------------------------------------------
#  Pagina 7
# ------------------------------------------------------------------------------------------------------------

def get_informe_cev_v2_pagina7_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 7 of an informe_CEV_v2 PDF report and return it as a dictionary.
    """
    result: Dict[str, Any] = {}
    try:
        if not isinstance(pdf_report, fitz.Document): raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 7: raise ValueError("PDF has less than 7 pages.")
        page = pdf_report[6]

        COORDINATES: Dict[str, Tuple[float, float, float, float]] = {
            'codigo_evaluacion': (63.5, 30.9, 84.0, 36.1), 'mandante_nombre': (27.5, 90.6, 96.0, 94.7),
            'mandante_rut': (27.5, 95.4, 96.0, 99.4), 'evaluador_nombre': (131.1, 90.6, 205.0, 94.7),
            'evaluador_rut': (131.1, 95.4, 196.7, 99.4), 'evaluador_rol_minvu': (150.0, 99.9, 166.0, 103.7)
        }

        fields: Dict[str, str] = { k: extract_text_from_area(page, v) for k, v in COORDINATES.items() }

        result = {
            'codigo_evaluacion': fields.get('codigo_evaluacion', '').strip(),
            'mandante_nombre': fields.get('mandante_nombre', '').strip(),
            'mandante_rut': fields.get('mandante_rut', '').strip(),
            'evaluador_nombre': fields.get('evaluador_nombre', '').strip(),
            'evaluador_rut': fields.get('evaluador_rut', '').strip(),
            'evaluador_rol_minvu': fields.get('evaluador_rol_minvu', '').strip()
        }
        return result

    except (IndexError, ValueError, TypeError) as e:
        logging.error(f"Error processing Page 7 dictionary: {e}", exc_info=True)
        return {}


def get_informe_cev_v2_pagina7_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts evaluation details from page 7 into a Pandas DataFrame."""
    data_dict = get_informe_cev_v2_pagina7_as_dict(pdf_report)
    if not data_dict: return pd.DataFrame()
    try:
        return pd.DataFrame.from_dict(data_dict, orient='index').T
    except Exception as e:
        logging.error(f"Failed to convert page 7 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()