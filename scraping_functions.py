from bisect import bisect_left
from typing import Dict, Tuple, Any, List, Union, Optional
from functools import lru_cache
import pandas as pd
import fitz  # PyMuPDF
import logging
import io
import re
from PIL import Image
import pytesseract
import numpy as np
import cv2

# ==============================================================================
# 1. FUNCIONES DE AYUDA (Helpers)
# ==============================================================================


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
        logging.error(
            "Report width or height cannot be zero for normalization.")
        return 0.0, 0.0


def get_page_coordinates(page_num: int) -> Dict[str, Tuple[float, float, float, float]]:
    """
    Get coordinates for each page based on the page number.

    Args:
        page_num: Page number (0-indexed)

    Returns:
        Dictionary with coordinates for the specified page
    """

    # Página 1 (índice 0)
    if page_num == 0:
        return {
            'tipo_evaluacion': (8.3, 9.0, 165.6, 18.8),
            'codigo_evaluacion': (73.1, 20.0, 97.1, 25.1),
            'region': (28.0, 26.6, 165.3, 31.8),
            'comuna': (29.2, 33.0, 165.3, 38.2),
            'direccion': (31.3, 39.1, 165.3, 44.3),
            'rol_vivienda_proyecto': (55.5, 45.6, 165.3, 50.8),
            'tipo_vivienda': (45.9, 51.7, 165.3, 56.9),
            'superficie_interior_util_m2': (53.0, 58.3, 70.0, 63.5),
            'porcentaje_ahorro_raw': (5.6, 78.6, 165.8, 191.3),
            'demanda_calefaccion_kwh_m2_ano_raw': (15.6, 220.0, 73.0, 230.0),
            'demanda_enfriamiento_kwh_m2_ano_raw': (90.0, 220.0, 151.5, 230.0),
            'demanda_total_kwh_m2_ano_raw': (167.0, 225.0, 209.0, 245.0),
            'emitida_el_raw': (34.5, 247.0, 57.5, 252.8)
        }

    # Página 2 (índice 1)
    elif page_num == 1:
        return {
            'region': (40.4, 47.4, 95.0, 51.7),
            'comuna': (40.4, 53.2, 95.0, 57.4),
            'direccion': (40.4, 58.9, 95.0, 63.1),
            'rol_vivienda': (40.4, 64.6, 95.0, 68.9),
            'tipo_vivienda': (40.4, 70.2, 95.0, 74.4),
            'zona_termica': (143.1, 47.5, 151.0, 51.7),
            'superficie_interior_util_m2_raw': (143.1, 53.3, 151.0, 57.5),
            'solicitado_por': (143.1, 58.9, 210.5, 63.1),
            'evaluado_por': (143.1, 64.7, 210.5, 68.9),
            'codigo_evaluacion': (143.1, 70.2, 163.0, 74.5),
            'demanda_calefaccion_kwh_m2_ano_raw': (98.7, 98.7, 109.5, 105.2),
            'demanda_enfriamiento_kwh_m2_ano_raw': (98.7, 120.9, 109.5, 126.5),
            'demanda_total_kwh_m2_ano_raw': (98.6, 137.0, 136.0, 149.5),
            'demanda_total_bis_kwh_m2_ano_raw': (39.2, 159.8, 122.8, 166.0),
            'demanda_total_referencia_kwh_m2_ano_raw': (16.9, 168.3, 146.2, 173.2),
            'porcentaje_ahorro_raw': (151.0, 162.6, 201.5, 168.7),
            'muro_principal_descripcion': (46.2, 202.5, 184.5, 209.2),
            'muro_principal_exigencia_raw': (185.5, 202.5, 209.5, 209.2),
            'muro_secundario_descripcion': (46.2, 209.5, 184.5, 216.2),
            'muro_secundario_exigencia_raw': (185.5, 209.5, 209.5, 216.2),
            'piso_principal_descripcion': (46.2, 216.5, 184.5, 223.2),
            'piso_principal_exigencia_raw': (185.5, 216.5, 209.5, 223.2),
            'puerta_principal_descripcion': (46.2, 223.5, 184.5, 230.2),
            'puerta_principal_exigencia_raw': (185.5, 223.5, 209.5, 230.2),
            'techo_principal_descripcion': (46.2, 230.5, 184.5, 237.0),
            'techo_principal_exigencia_raw': (185.5, 230.5, 209.5, 237.0),
            'techo_secundario_descripcion': (46.2, 237.6, 184.5, 244.1),
            'techo_secundario_exigencia_raw': (185.5, 237.6, 209.5, 244.1),
            'superficie_vidriada_principal_descripcion': (46.2, 244.6, 184.5, 251.2),
            'superficie_vidriada_principal_exigencia': (185.5, 244.6, 209.5, 251.2),
            'superficie_vidriada_secundaria_descripcion': (46.2, 251.6, 184.5, 258.2),
            'superficie_vidriada_secundaria_exigencia': (185.5, 251.6, 209.5, 258.2),
            'ventilacion_rah_descripcion': (46.2, 258.6, 184.5, 265.2),
            'ventilacion_rah_exigencia': (185.5, 258.6, 209.5, 265.2),
            'infiltraciones_rah_descripcion': (46.2, 265.6, 184.5, 272.2),
            'infiltraciones_rah_exigencia': (185.5, 265.6, 209.5, 272.2)
        }

    # Página 3 (índice 2) - Consumos
    elif page_num == 2:
        return {
            'codigo_evaluacion': (62.3, 30.7, 88.1, 36.0),
            'agua_caliente_sanitaria_kwh_m2_raw': (79.2, 73.4, 98.3, 77.0),
            'agua_caliente_sanitaria_per_raw': (99.4, 73.4, 117.3, 77.0),
            'iluminacion_kwh_m2_raw': (79.2, 77.7, 98.3, 81.9),
            'iluminacion_per_raw': (98.7, 77.7, 117.3, 81.9),
            'calefaccion_kwh_m2_raw': (79.2, 82.3, 98.3, 86.5),
            'calefaccion_kwh_per_raw': (98.7, 82.3, 117.3, 86.5),
            'energia_renovable_no_convencional_kwh_m2_raw': (79.2, 87.0, 98.3, 91.2),
            'energia_renovable_no_convencional_per_raw': (98.7, 87.0, 117.3, 91.2),
            'consumo_total_kwh_m2_raw': (118.0, 74.0, 149.3, 86.0),
            'emisiones_kgco2_m2_ano_raw': (171.5, 69.0, 184.3, 74.2),
            'calefaccion_descripcion_proy': (76.6, 101.4, 155.5, 105.3),
            'calefaccion_consumo_proy_kwh_raw': (157.0, 101.4, 196.0, 105.3),
            'calefaccion_consumo_proy_per_raw': (198.0, 101.4, 209.0, 105.3),
            'iluminacion_descripcion_proy': (76.6, 106.2, 155.5, 110.0),
            'iluminacion_consumo_proy_kwh_raw': (157.0, 106.2, 196.0, 110.0),
            'iluminacion_consumo_proy_per_raw': (198.0, 106.2, 209.0, 110.0),
            'agua_caliente_sanitaria_descripcion_proy': (76.6, 111.2, 155.5, 115.0),
            'agua_caliente_sanitaria_consumo_proy_kwh_raw': (157.0, 111.2, 196.0, 115.0),
            'agua_caliente_sanitaria_consumo_proy_per_raw': (198.0, 111.2, 209.0, 115.0),
            'energia_renovable_no_convencional_descripcion_proy': (76.6, 115.8, 155.5, 120.0),
            'energia_renovable_no_convencional_consumo_proy_kwh_raw': (157.0, 115.8, 196.0, 120.0),
            'energia_renovable_no_convencional_consumo_proy_per_raw': (198.0, 115.8, 209.0, 120.0),
            'consumo_total_requerido_proy_kwh_raw': (157.0, 121.0, 196.0, 125.0),
            'calefaccion_descripcion_ref': (76.6, 136.1, 155.5, 140.1),
            'calefaccion_consumo_ref_kwh_raw': (157.0, 136.1, 196.0, 140.1),
            'calefaccion_consumo_ref_per_raw': (198.0, 136.1, 209.0, 140.1),
            'iluminacion_descripcion_ref': (76.6, 140.7, 155.5, 144.7),
            'iluminacion_consumo_ref_kwh_raw': (157.0, 140.7, 196.0, 144.7),
            'iluminacion_consumo_ref_per_raw': (198.0, 140.7, 209.0, 144.7),
            'agua_caliente_sanitaria_descripcion_ref': (76.6, 145.5, 155.5, 149.9),
            'agua_caliente_sanitaria_consumo_ref_kwh_raw': (157.0, 145.5, 196.0, 149.9),
            'agua_caliente_sanitaria_consumo_ref_per_raw': (198.0, 145.5, 209.0, 149.9),
            'energia_renovable_no_convencional_descripcion_ref': (76.6, 150.3, 155.5, 155.1),
            'energia_renovable_no_convencional_consumo_ref_kwh_raw': (157.0, 150.3, 196.0, 155.1),
            'energia_renovable_no_convencional_consumo_ref_per_raw': (198.0, 150.3, 209.0, 155.1),
            'consumo_total_requerido_ref_kwh_raw': (157.0, 155.5, 196.0, 161.0),
            # CONSUMOS SIN INCLUIR ERNC
            'consumo_ep_calefaccion_kwh_raw': (87.0, 176.0, 104.0, 179.5),
            'consumo_ep_agua_caliente_sanitaria_kwh_raw': (87.0, 180.0, 104.0, 183.5),
            'consumo_ep_iluminacion_kwh_raw': (87.0, 184.0, 104.0, 187.5),
            'consumo_ep_ventiladores_kwh_raw': (87.0, 188.0, 104.0, 191.5),
            # GENERACIÓN FOTOVOLTAICA EN LA VIVIENDA
            'generacion_ep_fotovoltaicos_kwh_raw': (87.0, 199.0, 104.0, 202.3),
            'aporte_fotovoltaicos_consumos_basicos_kwh_raw': (87.0, 202.8, 104.0, 206.4),
            'diferencia_fotovoltaica_para_consumo_kwh_raw': (87.0, 206.9, 104.0, 210.2),
            # DISTRIBUCIÓN DEL APORTE DE SOLAR TÉRMICA
            'aporte_solar_termica_calefaccion_kwh_raw': (87.0, 218.5, 104.0, 222.0),
            'aporte_solar_termica_agua_caliente_sanitaria_kwh_raw': (87.0, 222.5, 104.0, 225.8),
            # BALANCE GENERAL DE ENERGÍA
            'total_consumo_ep_antes_fotovoltaica_kwh_raw': (192.0, 176.0, 208.0, 179.5),
            'aporte_fotovoltaicos_consumos_basicos_kwh_bis_raw': (192.0, 180.0, 208.0, 183.5),
            'consumos_basicos_a_suplir_kwh_raw': (192.0, 183.9, 208.0, 187.3),
            # RESUMEN DE CONSUMOS FINALES DE REFERENCIA Y OBJETO
            'consumo_total_ep_obj_kwh_raw': (192.0, 199.0, 208.0, 202.5),
            'consumo_total_ep_ref_kwh_raw': (192.0, 202.8, 208.0, 206.5),
            'coeficiente_energetico_c_raw': (192.0, 207.0, 208.0, 210.5),
            # Coordenadas de envolvente también están en página 3
            'opacos_area_coords': (19.8, 245.6, 47.6, 287.3),
            'opacos_u_coords': (48.7, 245.6, 60.8, 287.3),
            'traslucidos_area_coords': (68.4, 245.6, 89.7, 283.0),
            'traslucidos_u_coords': (90.8, 245.6, 103.1, 283.0),
            'puentes_termicos_coords': (115.2, 249.8, 171.8, 283.1),
            'ua_phil_coords': (189.5, 245.6, 201.9, 287.3)
        }

    # Página 4 (índice 3) - Datos mensuales
    elif page_num == 3:
        coordinates = {'codigo_evaluacion': (62.3, 30.7, 88.1, 36.0)}

        # Coordenadas mensuales (12 columnas)
        dx = 13.5
        base_x = 42.0
        col_width = 12.0
        Y_COORDS = {
            'demanda_calef_viv_eval_kwh': (139.5, 143.5),
            'demanda_calef_viv_ref_kwh': (144.0, 148.0),
            'demanda_enfri_viv_eval_kwh': (161.4, 165.5),
            'demanda_enfri_viv_ref_kwh': (166.0, 170.2),
            'sobrecalentamiento_viv_eval_hr': (254.8, 258.8),
            'sobrecalentamiento_viv_ref_hr': (259.5, 263.4),
            'sobreenfriamiento_viv_eval_hr': (274.8, 278.8),
            'sobreenfriamiento_viv_ref_hr': (279.2, 283.3)
        }

        for key, (y1, y2) in Y_COORDS.items():
            for i in range(12):  # 12 meses
                x1 = base_x + i * dx
                x2 = x1 + col_width
                coordinates[f'{key}_mes_{i+1}'] = (x1, y1, x2, y2)

        return coordinates

    elif page_num == 4:
        coordinates = {
            'codigo_evaluacion': (62.3, 30.7, 88.1, 36.0),
            # Coordenadas para columna completa de Enero (ajustar según PDF real)
            # x1, y1, x2, y2 - columna completa
            'columna_enero': (46.5, 189.7, 62.0, 243.2),
            # Coordenadas para columna completa de Julio (ajustar según PDF real)
            # x1, y1, x2, y2 - columna completa
            'columna_julio': (76.5, 189.7, 92.0, 243.2)
        }
        return coordinates

    # Página 6 (índice 5)
    elif page_num == 5:
        return {
            'codigo_evaluacion': (62.3, 30.7, 88.1, 36.0),
            'enero': (64.5, 97.7, 173.2, 103.1),
            'abril': (64.6, 152.5, 173.8, 157.8),
            'julio': (66.4, 211.8, 174.0, 217.5),
            'octubre': (65.8, 271.0, 174.5, 276.4),
        }

    # Página 7 (índice 6)
    elif page_num == 6:
        return {
            'codigo_evaluacion': (62.3, 30.7, 88.1, 36.0),
            'mandante_nombre': (27.5, 90.6, 96.0, 94.5),
            'mandante_rut': (27.5, 95.2, 96.0, 99.0),
            'evaluador_nombre': (131.1, 90.6, 209.3, 94.5),
            'evaluador_rut': (131.1, 95.2, 209.3, 99.0),
            'evaluador_rol_minvu': (150.0, 99.9, 171.0, 103.7)
        }

    else:
        return {}


def draw_extraction_rectangles(pdf_report: fitz.Document, page_num: int, coordinates: Dict[str, Tuple[float, float, float, float]] = None, output_path: str = None) -> fitz.Document:
    """
    Draw rectangles on a specific page of the PDF to visualize the extraction areas.
    Uses normalized coordinates to match the scale used in extract_text_from_area.

    Args:
        pdf_report: fitz.Document object
        page_num: Page number (0-indexed)
        coordinates: Dictionary with field names and their coordinates. If None, uses get_page_coordinates()
        output_path: Optional path to save the modified PDF

    Returns:
        fitz.Document: The modified document with rectangles drawn
    """
    try:
        if not isinstance(pdf_report, fitz.Document):
            raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) <= page_num:
            raise ValueError(f"PDF has less than {page_num + 1} pages.")

        # Usar get_page_coordinates si no se proporcionan coordenadas
        if coordinates is None:
            coordinates = get_page_coordinates(page_num)

        if not coordinates:
            logging.warning(f"No coordinates found for page {page_num + 1}")
            return pdf_report

        # Constantes de normalización
        REPORT_WIDTH = 215.9  # mm
        REPORT_HEIGHT = 330.0  # mm

        # Color único para todo el informe (azul por defecto)
        DEFAULT_COLOR = (0, 0, 1)

        # Obtener la página especificada
        page = pdf_report[page_num]

        # Obtener dimensiones de la página
        page_rect = page.rect
        if page_rect is None:
            raise ValueError("Could not get page rectangle.")

        page_width = page_rect.width
        page_height = page_rect.height

        if page_width <= 0 or page_height <= 0:
            raise ValueError(
                f"Invalid page dimensions: width={page_width}, height={page_height}")

        # Dibujar rectángulos para cada coordenada
        rectangles_drawn = 0
        for field_name, coords in coordinates.items():
            try:
                x1, y1, x2, y2 = coords

                # Validar coordenadas originales
                if not all(isinstance(coord, (int, float)) for coord in coords):
                    logging.warning(
                        f"Invalid coordinates for {field_name}: {coords}")
                    continue

                if x1 >= x2 or y1 >= y2:
                    logging.warning(
                        f"Invalid rectangle for {field_name}: {coords}")
                    continue

                # Normalizar coordenadas
                rx1, ry1 = normalize_coordinates(
                    x1, y1, REPORT_WIDTH, REPORT_HEIGHT, page_width, page_height)
                rx2, ry2 = normalize_coordinates(
                    x2, y2, REPORT_WIDTH, REPORT_HEIGHT, page_width, page_height)

                # Verificar coordenadas normalizadas
                if rx1 >= rx2 or ry1 >= ry2:
                    logging.warning(
                        f"Normalized coordinates invalid for {field_name}: ({rx1}, {ry1}, {rx2}, {ry2}) from original {coords}")
                    continue

                # Crear rectángulo normalizado
                rect = fitz.Rect(rx1, ry1, rx2, ry2)

                # Dibujar rectángulo con línea fina (0.75) y color único
                page.draw_rect(rect, color=DEFAULT_COLOR, width=0.75)
                rectangles_drawn += 1

            except Exception as e:
                logging.error(f"Error drawing rectangle for {field_name}: {e}")
                continue

        logging.info(f"Page {page_num + 1}: Successfully drew {rectangles_drawn} rectangles out of {len(coordinates)} defined coordinates.")

        # Guardar si se especifica una ruta
        if output_path:
            pdf_report.save(output_path)
            logging.info(f"PDF con rectángulos guardado en: {output_path}")

        return pdf_report

    except Exception as e:
        logging.error(f"Error drawing rectangles on page {page_num + 1}: {e}")
        logging.error(f"Error drawing rectangles on page {page_num + 1}: {e}")
        return pdf_report


def draw_all_pages_rectangles(pdf_report: fitz.Document, output_path: str = None) -> fitz.Document:
    """
    Draw rectangles on all pages of the PDF report using coordinates defined by get_page_coordinates().

    Args:
        pdf_report: fitz.Document object
        output_path: Optional path to save the modified PDF

    Returns:
        fitz.Document: The modified document with rectangles drawn on all pages
    """
    try:
        if not isinstance(pdf_report, fitz.Document):
            raise TypeError("Input must be a fitz.Document object.")

        total_rectangles = 0

        # Procesar cada página
        for page_num in range(min(7, len(pdf_report))):  # Máximo 7 páginas
            coordinates = get_page_coordinates(page_num)

            if coordinates:
                draw_extraction_rectangles(pdf_report, page_num, coordinates)
                total_rectangles += len(coordinates)
            else:
                logging.warning(f"No coordinates defined for page {page_num + 1}")

        logging.info(f"Total rectangles drawn across all pages: {total_rectangles}")

        # Guardar si se especifica una ruta
        if output_path:
            pdf_report.save(output_path)
            logging.info(f"PDF completo con rectángulos guardado en: {output_path}")

        return pdf_report

    except Exception as e:
        logging.error(f"Error drawing rectangles on all pages: {e}")
        logging.error(f"Error drawing rectangles on all pages: {e}")
        return pdf_report


def extract_text_from_area(page: fitz.Page, area: Tuple[float, float, float, float]) -> str:
    """
    Extract text from a specific area of a PDF page. Robust error handling.
    """
    if not isinstance(page, fitz.Page):
        logging.error(
            "Invalid page object provided to extract_text_from_area.")
        return ""

    if not isinstance(area, tuple) or len(area) != 4:
        logging.error(
            f"Invalid area format provided: {area}. Must be a tuple of 4 coordinates.")
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
            logging.error(
                f"Invalid page dimensions in extract_text_from_area: width={width}, height={height}")
            return ""

        x1, y1, x2, y2 = area
        if not all(isinstance(coord, (int, float)) for coord in area):
            logging.error(
                f"Coordinates must be numeric in extract_text_from_area: {area}")
            return ""

        if x1 >= x2 or y1 >= y2:
            logging.warning(
                f"Invalid coordinates provided: {area}. Ensure x1 < x2 and y1 < y2.")
            return ""

        # Normalize coordinates
        rx1, ry1 = normalize_coordinates(
            x1, y1, REPORT_WIDTH, REPORT_HEIGHT, width, height)
        rx2, ry2 = normalize_coordinates(
            x2, y2, REPORT_WIDTH, REPORT_HEIGHT, width, height)

        # Ensure normalized coordinates create a valid rectangle
        if rx1 >= rx2 or ry1 >= ry2:
            logging.warning(
                f"Normalized coordinates resulted in invalid rectangle: ({rx1}, {ry1}, {rx2}, {ry2}) from area {area}")
            return ""

        rect = fitz.Rect(rx1, ry1, rx2, ry2)
        extracted_text = page.get_textbox(rect)
        return extracted_text.strip() if extracted_text else ""

    except ZeroDivisionError:
        logging.error(
            "Division by zero error during coordinate normalization in extract_text_from_area.")
        return ""
    except Exception as e:
        logging.error(
            f"Unexpected error extracting text from area {area}: {e}", exc_info=True)
        return ""


def _get_text_from_image_area(page: fitz.Page, crop_box: fitz.Rect) -> str:
    """
    Extrae texto de un área específica usando OCR.
    """
    try:
        # 1. Capturar con DPI alto para máxima claridad
        pix = page.get_pixmap(dpi=600, clip=crop_box)
        img_data = pix.tobytes("png")

        # Convertir los datos de la imagen a un formato que OpenCV pueda usar
        np_arr = np.frombuffer(img_data, np.uint8)
        img_cv = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)

        # 2. Reducir la imagen a la mitad (50% de su tamaño)
        # Esto a menudo la devuelve al tamaño "ideal" para Tesseract
        nueva_altura = int(img_cv.shape[0] * 0.5)
        nueva_anchura = int(img_cv.shape[1] * 0.5)
        img_reescalada = cv2.resize(
            img_cv, (nueva_anchura, nueva_altura), interpolation=cv2.INTER_AREA)

        # 3. Realizar OCR en la imagen reescalada
        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(
            img_reescalada, config=custom_config)
        return text
    except Exception as e:
        logging.error(f"Error durante el OCR en la página {page.number}: {e}")
        return ""




def _parse_hourly_temps_from_text_to_dict(raw_text: str, month_name: str) -> Dict[str, List[float]]:
    """
    Procesa el texto raw de un mes y retorna un diccionario con las temperaturas.

    Args:
        raw_text: Texto extraído con OCR
        month_name: Nombre del mes ('enero', 'abril', 'julio', 'octubre')

    Returns:
        Dict con keys f't_ext_{month_name}' y f't_int_{month_name}', cada una con 24 valores
    """
    try:
        # Extraer números del texto
        text_without_decimals = raw_text.replace('.', '')
        digit_sequences = re.findall(r'\d+', text_without_decimals)
        decimal_numbers = [
            float(seq) / 10 for seq in digit_sequences if seq.isdigit()]
        # Aplicar la función lambda para limitar numero de digitos antes del punto decimal a maximo 2
        decimal_numbers = list(map(lambda x: float(str(x).split('.')[0][-2:] + '.' + str(x).split(
            '.')[1]) if '.' in str(x) and len(str(x).split('.')[0]) > 2 else x, decimal_numbers))
        # === PASO 1: SEPARACIÓN ===
        lista_inicial = separar_string(raw_text)

        # === PASO 2: PROCESAMIENTO ===
        lista_digitos = procesar_lista_completa(lista_inicial)

        # === PASO 3: CONVERSIÓN Y DIVISIÓN ===

        decimal_numbers = convertir_y_dividir(lista_digitos)

        if len(decimal_numbers) < 48:
            logging.warning(
                f"Se esperaban 48 valores de temp para {month_name}, pero se encontraron {len(decimal_numbers)}.")
            # Rellenar con None si faltan valores
            while len(decimal_numbers) < 48:
                decimal_numbers.append(None)

        # Dividir en temperaturas exteriores e interiores
        temp_ext_values = decimal_numbers[0:24]
        temp_int_values = decimal_numbers[24:48]

        return {
            f't_ext_{month_name}': temp_ext_values,
            f't_int_{month_name}': temp_int_values
        }

    except Exception as e:
        logging.error(f"Error procesando temperaturas para {month_name}: {e}")
        return {
            f't_ext_{month_name}': [None] * 24,
            f't_int_{month_name}': [None] * 24
        }

# --- Helper Functions ---


def safe_float_convert(text: Optional[str], default: Any = None) -> Union[float, None]:
    """
    Safely converts a string to a float, handling different locale conventions.
    
    Supports:
    - Spanish/European format: 1.234,56
    - US/Standard format: 1,234.56
    - Mixed multi-line OCR output (tries to find a valid number line by line).
    
    Args:
        text: The string to convert.
        default: Value to return if conversion fails.
        
    Returns:
        Converted float or the default value.
    """
    if text is None or text == '':
        return default

    # Handle multi-line text by splitting and processing each line
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return default

    for line in lines:
        try:
            # Basic cleaning
            cleaned = line.strip()
            
            # If both , and . exist (e.g. 1.234,56 or 1,234.56)
            if ',' in cleaned and '.' in cleaned:
                if cleaned.find('.') < cleaned.find(','): # 1.234,56
                    cleaned = cleaned.replace('.', '').replace(',', '.')
                else: # 1,234.56
                    cleaned = cleaned.replace(',', '')
            # If only comma exists (e.g. 1234,56)
            elif ',' in cleaned:
                cleaned = cleaned.replace(',', '.')
            # If only dot exists (e.g. 1.234 or 1234.56)
            elif '.' in cleaned:
                parts = cleaned.split('.')
                # If there's exactly one dot and it's not followed by exactly 3 digits (e.g. 75.5)
                # it's likely a decimal dot.
                if len(parts) == 2 and len(parts[1]) != 3:
                    pass # Keep the dot as decimal separator
                # If there are multiple dots OR one dot followed by 3 digits (e.g. 1.000 or 1.000.000)
                # verify that all parts after the first have length 3 (thousand separator pattern)
                elif all(len(p) == 3 for p in parts[1:]):
                    cleaned = cleaned.replace('.', '') # Treat as thousands separator
                else:
                    # Invalid or ambiguous dot usage
                    raise ValueError(f"Ambiguous or invalid numeric format: {cleaned}")

            return float(cleaned)
        except (ValueError, TypeError):
            continue

    logging.warning(f"Could not convert '{text}' to float.")
    return default


def _from_procentaje_ahorro_to_letra(porcentaje_ahorro_decimal: Optional[float]) -> Optional[str]:
    """
    Convert a savings percentage (as a float, e.g., 0.75 for 75%) to a corresponding letter grade.
    Handles None input gracefully.
    """
    if porcentaje_ahorro_decimal is None:
        return None
    boundaries = [-0.35, -0.1, 0.2, 0.4, 0.55, 0.7, 0.85, 100.0]
    grades = ['G', 'F', 'E', 'D', 'C', 'B', 'A', 'A+']

    try:
        idx = bisect_left(boundaries, porcentaje_ahorro_decimal)
        if 0 <= idx < len(grades):
            return grades[idx]
        else:
            logging.warning(
                f"Percentage {porcentaje_ahorro_decimal*100}% resulted in out-of-bounds grade index {idx}.")
            if idx >= len(grades):
                return grades[-1]
            else:
                return grades[0]

    except TypeError:
        logging.error(
            f"Invalid type for percentage: {porcentaje_ahorro_decimal}. Cannot determine grade.")
        return None


def separar_string(texto):
    """
    Separa un string por espacios, saltos de línea o pestañas si no es un número.
    """
    elementos = re.split(r'[\s\n\t]+', texto)
    return [elem.strip() for elem in elementos if elem.strip()]


def extraer_numeros_de_texto(texto):
    """
    Extrae todos los números (enteros y decimales) de un string,
    manejando comas y puntos.
    """
    # Patrón para encontrar números (considerando comas y puntos como separadores)
    patron = r'\d+(?:[.,]\d+)*'
    numeros = re.findall(patron, texto)

    # Limpiar cada número encontrado
    resultados = []
    for num in numeros:
        # Usar la misma lógica de safe_float_convert para normalizar
        normalizado = num.replace('.', '') if '.' in num and (len(num.split('.')[-1]) == 3) else num
        normalizado = normalizado.replace(',', '.') if ',' in normalizado else normalizado
        # Eliminar cualquier punto restante que no sea decimal
        if normalizado.count('.') > 1:
             normalizado = normalizado.replace('.', '', normalizado.count('.') - 1)
        
        # Limpieza final para asegurar que sea un string numérico puro para la lógica posterior del scraper
        try:
            float(normalizado)
            # El scraper espera strings sin puntos si son enteros, o con punto decimal
            final = normalizado.replace('.', '') if '.' in normalizado and float(normalizado).is_integer() else normalizado
            resultados.append(final)
        except ValueError:
            continue

    return resultados


def procesar_lista_completa(lista_strings):
    """
    Procesa cada elemento de la lista:
    1. Intenta convertir a float (sin espacios)
    2. Si falla, aplica lógica de espacios según cantidad de dígitos
    3. Si no encuentra números, ELIMINA el elemento
    """
    lista_procesada = []

    for elemento in lista_strings:
        # Detectar caso especial: números separados por espacios
        if ' ' in elemento and re.search(r'\d+\s+\d+', elemento):
            resultado_espacios = manejar_numeros_con_espacios(elemento)
            if resultado_espacios:
                lista_procesada.extend(resultado_espacios)
            continue

        sin_espacios = elemento.replace(' ', '')

        try:
            # Intentar convertir a float
            float(sin_espacios)
            # Si tiene éxito, crear string sin espacios y sin puntos
            resultado = sin_espacios.replace('.', '')
            lista_procesada.append(resultado)
            logging.debug(f"✓ '{elemento}' -> '{resultado}'")

        except ValueError:
            # Si falla, extraer números individuales
            numeros_extraidos = extraer_numeros_de_elemento(elemento)

            if numeros_extraidos:
                logging.debug(
                    f"✗ '{elemento}' -> Extrayendo números: {numeros_extraidos}")
                # Agregar cada número como elemento separado
                for numero in numeros_extraidos:
                    lista_procesada.append(numero)
            else:
                # Si no se encuentran números, ELIMINAR
                logging.debug(f"🗑️ '{elemento}' -> Sin números válidos, ELIMINADO")

    return lista_procesada


def manejar_numeros_con_espacios(elemento):
    """
    Maneja elementos con números separados por espacios según reglas específicas:
    - Si ambos números tienen >= 2 dígitos: separar ['41', '36']
    - Si alguno tiene < 2 dígitos: unir '4136'
    """
    # Encontrar todos los números en el elemento
    numeros = re.findall(r'\d+\.?\d*', elemento)

    if len(numeros) < 2:
        # Si hay menos de 2 números, usar lógica normal
        return extraer_numeros_de_elemento(elemento)

    # Verificar cantidad de dígitos en cada número
    digitos_por_numero = []
    numeros_sin_punto = []

    for num in numeros:
        try:
            float(num)  # Verificar que sea válido
            num_sin_punto = num.replace('.', '')
            numeros_sin_punto.append(num_sin_punto)
            digitos_por_numero.append(len(num_sin_punto))
        except ValueError:
            continue

    if len(numeros_sin_punto) < 2:
        return extraer_numeros_de_elemento(elemento)

    # Aplicar regla: todos deben tener >= 2 dígitos para separar
    todos_tienen_2_o_mas = all(digitos >= 2 for digitos in digitos_por_numero)

    if todos_tienen_2_o_mas:
        # Separar números
        logging.debug(
            f"🔄 '{elemento}' -> Separando (todos >= 2 dígitos): {numeros_sin_punto}")
        return numeros_sin_punto
    else:
        # Unir números (remover espacios)
        numero_unido = ''.join(numeros_sin_punto)
        logging.debug(f"🔗 '{elemento}' -> Uniendo (alguno < 2 dígitos): '{numero_unido}'")
        return [numero_unido]


def extraer_numeros_de_elemento(elemento):
    """
    Extrae números individuales de un elemento (función auxiliar)
    """
    return extraer_numeros_de_texto(elemento)


def convertir_y_dividir(lista_digitos):
    """
    Convierte cada elemento de la lista a float y lo divide por 10
    """
    lista_final = []

    for elemento in lista_digitos:
        try:
            # Convertir a float
            numero_float = float(elemento)
            # Dividir por 10
            resultado = numero_float / 10
            lista_final.append(resultado)
            logging.debug(f"'{elemento}' -> {numero_float} -> {resultado}")

        except ValueError:
            logging.warning(f"⚠️ Error: '{elemento}' no se pudo convertir a float")
            # Opcional: agregar None o saltar el elemento
            # lista_final.append(None)  # Si quieres mantener la posición

    return lista_final

# ------------------------------------------------------------------------------------------------------------
#  Pagina 1
# ------------------------------------------------------------------------------------------------------------


def get_informe_cev_v2_pagina1_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 1 of an informe_CEV_v2 PDF report and return it as a dictionary.
    Uses safe float conversion and get_page_coordinates for consistency.
    """
    result: Dict[str, Any] = {}
    try:
        if not isinstance(pdf_report, fitz.Document):
            raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 1:
            raise ValueError("PDF has no pages.")
        page = pdf_report[0]

        # Usar get_page_coordinates para obtener las coordenadas
        COORDINATES = get_page_coordinates(0)

        fields: Dict[str, str] = {k: extract_text_from_area(
            page, v) for k, v in COORDINATES.items()}

        # Post-processing with safe conversion
        porcentaje_ahorro_str = next((line for line in fields.get(
            'porcentaje_ahorro_raw', '').splitlines() if line.replace('-', '').isdigit()), None)
        porcentaje_ahorro_int = int(
            porcentaje_ahorro_str) if porcentaje_ahorro_str is not None else None
        porcentaje_ahorro_decimal = float(
            porcentaje_ahorro_int / 100.0) if porcentaje_ahorro_int is not None else None

        demanda_cal_str = fields.get(
            'demanda_calefaccion_kwh_m2_ano_raw', '').splitlines()
        demanda_enf_str = fields.get(
            'demanda_enfriamiento_kwh_m2_ano_raw', '').splitlines()
        demanda_tot_str = fields.get(
            'demanda_total_kwh_m2_ano_raw', '').splitlines()
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
            'porcentaje_ahorro': porcentaje_ahorro_int,
            'letra_eficiencia_energetica_dem': _from_procentaje_ahorro_to_letra(porcentaje_ahorro_decimal),
            'demanda_calefaccion_kwh_m2_ano': safe_float_convert(demanda_cal_str[-1] if demanda_cal_str else None),
            'demanda_enfriamiento_kwh_m2_ano': safe_float_convert(demanda_enf_str[-1] if demanda_enf_str else None),
            'demanda_total_kwh_m2_ano': safe_float_convert(demanda_tot_str[-1] if demanda_tot_str else None),
            'emitida_el': emitida_str[-1].strip() if emitida_str else None
        }
        return result

    except (IndexError, ValueError, TypeError) as e:
        logging.error(
            f"Error processing Page 1 dictionary: {e}", exc_info=True)
        return {}


def get_informe_cev_v2_pagina1_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts data from page 1 into a Pandas DataFrame."""
    data_dict = get_informe_cev_v2_pagina1_as_dict(pdf_report)
    if not data_dict:
        return pd.DataFrame()
    try:
        df = pd.DataFrame.from_dict(data_dict, orient='index').T
        if "codigo_evaluacion" in df.columns:
            df = df.drop(columns=["codigo_evaluacion"])
        return df
    except Exception as e:
        logging.error(
            f"Failed to convert page 1 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()

# ------------------------------------------------------------------------------------------------------------
#  Pagina 2
# ------------------------------------------------------------------------------------------------------------


def get_informe_cev_v2_pagina2_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 2 of an informe_CEV_v2 PDF report and return it as a dictionary.
    Uses safe float conversion and get_page_coordinates for consistency.
    """
    result: Dict[str, Any] = {}
    try:
        if not isinstance(pdf_report, fitz.Document):
            raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 2:
            raise ValueError("PDF has less than 2 pages.")
        page = pdf_report[1]

        # Usar get_page_coordinates para obtener las coordenadas
        COORDINATES = get_page_coordinates(1)

        fields: Dict[str, str] = {k: extract_text_from_area(
            page, v) for k, v in COORDINATES.items()}

        # Helper lambdas for cleaner processing
        def get_last_line(key): return fields.get(
            key, '').splitlines()[-1].strip() if fields.get(key) else None

        def get_last_line_float(
            key): return safe_float_convert(get_last_line(key))

        def clean_desc(key): return fields.get(
            key, '').replace('\n', ' ').strip()

        def clean_exigencia_float(key: str) -> Optional[float]:
            """
            Convierte valores de exigencia técnica (rango 0-5, un dígito antes del decimal).
            Maneja tanto punto como coma decimal automáticamente.
            """
            try:
                raw_text = fields.get(key, '').replace('[W/m2K]', '').strip()
                if not raw_text:
                    return None

                # Simplemente reemplazar coma por punto y convertir
                # Como sabemos que son valores pequeños (0-5), no hay separadores de miles
                cleaned_text = raw_text.replace(',', '.')
                return float(cleaned_text)

            except (ValueError, TypeError):
                logging.warning(
                    f"Could not convert exigencia value '{raw_text}' to float for key '{key}'.")
                return None

        result = {
            'region': clean_desc('region'),
            'comuna': clean_desc('comuna'),
            'direccion': clean_desc('direccion'),
            'rol_vivienda': clean_desc('rol_vivienda'),
            'tipo_vivienda': clean_desc('tipo_vivienda'),
            'zona_termica': clean_desc('zona_termica'),
            'superficie_interior_util_m2': safe_float_convert(fields.get('superficie_interior_util_m2_raw')),
            'solicitado_por': clean_desc('solicitado_por'),
            'evaluado_por': clean_desc('evaluado_por'),
            'codigo_evaluacion': clean_desc('codigo_evaluacion'),
            'demanda_calefaccion_kwh_m2_ano': get_last_line_float('demanda_calefaccion_kwh_m2_ano_raw'),
            'demanda_enfriamiento_kwh_m2_ano': get_last_line_float('demanda_enfriamiento_kwh_m2_ano_raw'),
            'demanda_total_kwh_m2_ano': get_last_line_float('demanda_total_kwh_m2_ano_raw'),
            'demanda_total_bis_kwh_m2_ano': get_last_line_float('demanda_total_bis_kwh_m2_ano_raw'),
            'demanda_total_referencia_kwh_m2_ano': get_last_line_float('demanda_total_referencia_kwh_m2_ano_raw'),
            'porcentaje_ahorro': get_last_line_float('porcentaje_ahorro_raw'),
            'muro_principal_descripcion': clean_desc('muro_principal_descripcion'),
            'muro_principal_exigencia_w_m2_k': clean_exigencia_float('muro_principal_exigencia_raw'),
            'muro_secundario_descripcion': clean_desc('muro_secundario_descripcion'),
            'muro_secundario_exigencia_w_m2_k': clean_exigencia_float('muro_secundario_exigencia_raw'),
            'piso_principal_descripcion': clean_desc('piso_principal_descripcion'),
            'piso_principal_exigencia_w_m2_k': clean_exigencia_float('piso_principal_exigencia_raw'),
            'puerta_principal_descripcion': clean_desc('puerta_principal_descripcion'),
            'puerta_principal_exigencia_w_m2_k': clean_desc('puerta_principal_exigencia_raw'),
            'techo_principal_descripcion': clean_desc('techo_principal_descripcion'),
            'techo_principal_exigencia_w_m2_k': clean_exigencia_float('techo_principal_exigencia_raw'),
            'techo_secundario_descripcion': clean_desc('techo_secundario_descripcion'),
            'techo_secundario_exigencia_w_m2_k': clean_exigencia_float('techo_secundario_exigencia_raw'),
            'superficie_vidriada_principal_descripcion': clean_desc('superficie_vidriada_principal_descripcion'),
            'superficie_vidriada_principal_exigencia': clean_desc('superficie_vidriada_principal_exigencia'),
            'superficie_vidriada_secundaria_descripcion': clean_desc('superficie_vidriada_secundaria_descripcion'),
            'superficie_vidriada_secundaria_exigencia': clean_desc('superficie_vidriada_secundaria_exigencia'),
            'ventilacion_rah_descripcion': clean_desc('ventilacion_rah_descripcion'),
            'ventilacion_rah_exigencia': clean_desc('ventilacion_rah_exigencia'),
            'infiltraciones_rah_descripcion': clean_desc('infiltraciones_rah_descripcion'),
            'infiltraciones_rah_exigencia': clean_desc('infiltraciones_rah_exigencia')
        }
        return result

    except (IndexError, ValueError, TypeError) as e:
        logging.error(
            f"Error processing Page 2 dictionary: {e}", exc_info=True)
        return {}


def get_informe_cev_v2_pagina2_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts data from page 2 into a Pandas DataFrame."""
    data_dict = get_informe_cev_v2_pagina2_as_dict(pdf_report)
    if not data_dict:
        return pd.DataFrame()
    try:
        df = pd.DataFrame.from_dict(data_dict, orient='index').T
        if "codigo_evaluacion" in df.columns:
            df = df.drop(columns=["codigo_evaluacion"])
        return df
    except Exception as e:
        logging.error(
            f"Failed to convert page 2 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()


# ------------------------------------------------------------------------------------------------------------
#  Pagina 3 - Consumos
# ------------------------------------------------------------------------------------------------------------

def get_informe_cev_v2_pagina3_consumos_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 3 (consumos) of an informe_CEV_v2 PDF report and return it as a dictionary.
    Uses safe float conversion and get_page_coordinates for consistency.
    """
    result: Dict[str, Any] = {}
    try:
        if not isinstance(pdf_report, fitz.Document):
            raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 3:
            raise ValueError("PDF has less than 3 pages.")
        page = pdf_report[2]

        # Usar get_page_coordinates para obtener las coordenadas
        COORDINATES = get_page_coordinates(2)

        fields: Dict[str, str] = {k: extract_text_from_area(
            page, v) for k, v in COORDINATES.items()}

        def get_float(key): return safe_float_convert(fields.get(key))

        def get_last_line_float(key): return safe_float_convert(
            fields.get(key, '').splitlines()[-1] if fields.get(key) else None)
        def clean_desc(key): return fields.get(
            key, '').replace('\n', ' ').strip()

        result = {
            'codigo_evaluacion': clean_desc('codigo_evaluacion'),
            'agua_caliente_sanitaria_kwh_m2': get_float('agua_caliente_sanitaria_kwh_m2_raw'),
            'agua_caliente_sanitaria_per': get_float('agua_caliente_sanitaria_per_raw'),
            'iluminacion_kwh_m2': get_float('iluminacion_kwh_m2_raw'),
            'iluminacion_per': get_float('iluminacion_per_raw'),
            'calefaccion_kwh_m2': get_float('calefaccion_kwh_m2_raw'),
            'calefaccion_kwh_per': get_float('calefaccion_kwh_per_raw'),
            'energia_renovable_no_convencional_kwh_m2': get_float('energia_renovable_no_convencional_kwh_m2_raw'),
            'energia_renovable_no_convencional_per': get_float('energia_renovable_no_convencional_per_raw'),
            'consumo_total_kwh_m2': get_float('consumo_total_kwh_m2_raw'),
            'emisiones_kgco2_m2_ano': get_float('emisiones_kgco2_m2_ano_raw'),
            'calefaccion_descripcion_proy': clean_desc('calefaccion_descripcion_proy'),
            'calefaccion_consumo_proy_kwh': get_last_line_float('calefaccion_consumo_proy_kwh_raw'),
            'calefaccion_consumo_proy_per': get_last_line_float('calefaccion_consumo_proy_per_raw'),
            'iluminacion_descripcion_proy': clean_desc('iluminacion_descripcion_proy'),
            'iluminacion_consumo_proy_kwh': get_last_line_float('iluminacion_consumo_proy_kwh_raw'),
            'iluminacion_consumo_proy_per': get_last_line_float('iluminacion_consumo_proy_per_raw'),
            'agua_caliente_sanitaria_descripcion_proy': clean_desc('agua_caliente_sanitaria_descripcion_proy'),
            'agua_caliente_sanitaria_consumo_proy_kwh': get_last_line_float('agua_caliente_sanitaria_consumo_proy_kwh_raw'),
            'agua_caliente_sanitaria_consumo_proy_per': get_last_line_float('agua_caliente_sanitaria_consumo_proy_per_raw'),
            'energia_renovable_no_convencional_descripcion_proy': clean_desc('energia_renovable_no_convencional_descripcion_proy'),
            'energia_renovable_no_convencional_consumo_proy_kwh': get_last_line_float('energia_renovable_no_convencional_consumo_proy_kwh_raw'),
            'energia_renovable_no_convencional_consumo_proy_per': get_last_line_float('energia_renovable_no_convencional_consumo_proy_per_raw'),
            'consumo_total_requerido_proy_kwh': get_last_line_float('consumo_total_requerido_proy_kwh_raw'),
            'calefaccion_descripcion_ref': clean_desc('calefaccion_descripcion_ref'),
            'calefaccion_consumo_ref_kwh': get_last_line_float('calefaccion_consumo_ref_kwh_raw'),
            'calefaccion_consumo_ref_per': get_last_line_float('calefaccion_consumo_ref_per_raw'),
            'iluminacion_descripcion_ref': clean_desc('iluminacion_descripcion_ref'),
            'iluminacion_consumo_ref_kwh': get_last_line_float('iluminacion_consumo_ref_kwh_raw'),
            'iluminacion_consumo_ref_per': get_last_line_float('iluminacion_consumo_ref_per_raw'),
            'agua_caliente_sanitaria_descripcion_ref': clean_desc('agua_caliente_sanitaria_descripcion_ref'),
            'agua_caliente_sanitaria_consumo_ref_kwh': get_last_line_float('agua_caliente_sanitaria_consumo_ref_kwh_raw'),
            'agua_caliente_sanitaria_consumo_ref_per': get_last_line_float('agua_caliente_sanitaria_consumo_ref_per_raw'),
            'energia_renovable_no_convencional_descripcion_ref': clean_desc('energia_renovable_no_convencional_descripcion_ref'),
            'energia_renovable_no_convencional_consumo_ref_kwh': get_last_line_float('energia_renovable_no_convencional_consumo_ref_kwh_raw'),
            'energia_renovable_no_convencional_consumo_ref_per': get_last_line_float('energia_renovable_no_convencional_consumo_ref_per_raw'),
            'consumo_total_requerido_ref_kwh': get_last_line_float('consumo_total_requerido_ref_kwh_raw'),
            # CONSUMOS SIN INCLUIR ERNC
            'consumo_ep_calefaccion_kwh': get_float('consumo_ep_calefaccion_kwh_raw'),
            'consumo_ep_agua_caliente_sanitaria_kwh': get_float('consumo_ep_agua_caliente_sanitaria_kwh_raw'),
            'consumo_ep_iluminacion_kwh': get_float('consumo_ep_iluminacion_kwh_raw'),
            'consumo_ep_ventiladores_kwh': get_float('consumo_ep_ventiladores_kwh_raw'),
            # GENERACIÓN FOTOVOLTAICA EN LA VIVIENDA
            'generacion_ep_fotovoltaicos_kwh': get_float('generacion_ep_fotovoltaicos_kwh_raw'),
            'aporte_fotovoltaicos_consumos_basicos_kwh': get_float('aporte_fotovoltaicos_consumos_basicos_kwh_raw'),
            'diferencia_fotovoltaica_para_consumo_kwh': get_float('diferencia_fotovoltaica_para_consumo_kwh_raw'),
            # DISTRIBUCIÓN DEL APORTE DE SOLAR TÉRMICA
            'aporte_solar_termica_calefaccion_kwh': get_float('aporte_solar_termica_calefaccion_kwh_raw'),
            'aporte_solar_termica_agua_caliente_sanitaria_kwh': get_float('aporte_solar_termica_agua_caliente_sanitaria_kwh_raw'),
            # BALANCE GENERAL DE ENERGÍA
            'total_consumo_ep_antes_fotovoltaica_kwh': get_float('total_consumo_ep_antes_fotovoltaica_kwh_raw'),
            'aporte_fotovoltaicos_consumos_basicos_kwh_bis': get_float('aporte_fotovoltaicos_consumos_basicos_kwh_bis_raw'),
            'consumos_basicos_a_suplir_kwh': get_float('consumos_basicos_a_suplir_kwh_raw'),
            # RESUMEN DE CONSUMOS FINALES DE REFERENCIA Y OBJETO
            'consumo_total_ep_obj_kwh': get_float('consumo_total_ep_obj_kwh_raw'),
            'consumo_total_ep_ref_kwh': get_float('consumo_total_ep_ref_kwh_raw'),
            'coeficiente_energetico_c': get_float('coeficiente_energetico_c_raw')
        }
        return result

    except (IndexError, ValueError, TypeError) as e:
        logging.error(
            f"Error processing Page 3 (Consumos) dictionary: {e}", exc_info=True)
        return {}


def get_informe_cev_v2_pagina3_consumos_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts consumption data from page 3 into a Pandas DataFrame."""
    data_dict = get_informe_cev_v2_pagina3_consumos_as_dict(pdf_report)
    if not data_dict:
        return pd.DataFrame()
    try:
        df = pd.DataFrame.from_dict(data_dict, orient='index').T
        if "codigo_evaluacion" in df.columns:
            df = df.drop(columns=["codigo_evaluacion"])
        return df
    except Exception as e:
        logging.error(
            f"Failed to convert page 3 consumos dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()

# ------------------------------------------------------------------------------------------------------------
#  Pagina 3 - Envolvente
# ------------------------------------------------------------------------------------------------------------


def get_informe_cev_v2_pagina3_envolvente_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extracts envelope data from page 3 into a dictionary (structured for DataFrame).
    Uses safe float conversion and get_page_coordinates for consistency.
    """
    data_list: Dict[str, List[Any]] = {}
    try:
        if not isinstance(pdf_report, fitz.Document):
            raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 3:
            raise ValueError("PDF has less than 3 pages.")
        page = pdf_report[2]

        dy = 4.2
        num_orientations = 10
        num_puentes_termicos = 8
        puente_termico_start_y = 250.0
        orientations = ['Horiz', 'N', 'NE', 'E',
                        'SE', 'S', 'SO', 'O', 'NO', 'Pisos']

        # Usar get_page_coordinates para obtener las coordenadas base
        COORDINATES = get_page_coordinates(2)

        # Extraer coordenadas específicas de envolvente
        COORDINATES_BLOCKS = {
            'codigo_eval_coords': COORDINATES.get('codigo_evaluacion', (62.3, 30.7, 88.1, 36.0)),
            'opacos_area_coords': COORDINATES.get('opacos_area_coords', (19.8, 245.6, 47.6, 287.3)),
            'opacos_u_coords': COORDINATES.get('opacos_u_coords', (48.7, 245.6, 60.8, 287.3)),
            'traslucidos_area_coords': COORDINATES.get('traslucidos_area_coords', (68.4, 245.6, 89.7, 283.0)),
            'traslucidos_u_coords': COORDINATES.get('traslucidos_u_coords', (90.8, 245.6, 103.1, 283.0)),
            'ua_phil_coords': COORDINATES.get('ua_phil_coords', (189.5, 245.6, 201.9, 287.3))
        }

        PT_COORDS_BASE: Dict[str, Tuple[float, float]] = {
            'p01_w_k': (115.5, 124.5), 'p02_w_k': (126.2, 136.9), 'p03_w_k': (139.0, 148.2),
            'p04_w_k': (149.0, 160.0), 'p05_w_k': (161.3, 171.2)
        }

        # --- Extract Single Value ---
        codigo_evaluacion = extract_text_from_area(
            page, COORDINATES_BLOCKS['codigo_eval_coords']).strip()

        # --- Extract Columnar Data Blocks ---
        opacos_area_text = extract_text_from_area(
            page, COORDINATES_BLOCKS['opacos_area_coords'])
        opacos_U_text = extract_text_from_area(
            page, COORDINATES_BLOCKS['opacos_u_coords'])
        traslucidos_area_text = extract_text_from_area(
            page, COORDINATES_BLOCKS['traslucidos_area_coords'])
        traslucidos_U_text = extract_text_from_area(
            page, COORDINATES_BLOCKS['traslucidos_u_coords'])

        # --- Extract Puente Termico Data ---
        puentes_termicos_text: Dict[str, List[str]] = {
            key: [] for key in PT_COORDS_BASE}
        for key, (x1, x2) in PT_COORDS_BASE.items():
            for i in range(num_puentes_termicos):
                y1 = puente_termico_start_y + i * dy
                y2 = y1 + 3.5
                pt_coord = (x1, y1, x2, y2)
                text_lines = extract_text_from_area(
                    page, pt_coord).splitlines()
                puentes_termicos_text[key].append(
                    text_lines[-1] if text_lines else '')

        # --- Extract UA_phiL usando extracción individual por fila ---
        ua_phiL_values = []
        dy_ua = 3.5  # dy específico para UA_phiL (diferente del dy general)
        for n in range(0, 12):
            area_coordinates = (189.2, 245.5 + n * dy_ua,
                                201.9, 249.0 + n * dy_ua)
            extracted_text = extract_text_from_area(page, area_coordinates)
            if extracted_text:
                ua_phiL_line = extracted_text.splitlines()[-1]
                try:
                    ua_phiL_value = float(ua_phiL_line.replace(',', '.'))
                    ua_phiL_values.append(ua_phiL_value)
                except (ValueError, TypeError):
                    ua_phiL_values.append(None)
            else:
                ua_phiL_values.append(None)

        # Eliminar elementos en posiciones 4 y 9
        if len(ua_phiL_values) > 4:
            ua_phiL_values.pop(4)  # Luego el menor
        if len(ua_phiL_values) > 9:
            ua_phiL_values.pop(9)  # Eliminar primero el índice mayor

        # Asegurar que tenemos exactamente 10 valores
        while len(ua_phiL_values) < num_orientations:
            ua_phiL_values.append(None)
        # Truncar si hay más de 10
        ua_phiL_values = ua_phiL_values[:num_orientations]

        # --- Process and Structure Data ---
        data_list['codigo_evaluacion'] = [codigo_evaluacion] * num_orientations
        data_list['orientacion'] = orientations

        opacos_area_lines = opacos_area_text.splitlines()[-num_orientations:]
        opacos_U_lines = opacos_U_text.splitlines()[-num_orientations:]
        data_list['elementos_opacos_area_m2'] = [
            safe_float_convert(line) for line in opacos_area_lines]
        data_list['elementos_opacos_u_w_m2_k'] = [
            safe_float_convert(line) for line in opacos_U_lines]

        traslucidos_area_lines = traslucidos_area_text.splitlines(
        )[-(num_orientations-1):]
        traslucidos_U_lines = traslucidos_U_text.splitlines(
        )[-(num_orientations-1):]
        data_list['elementos_traslucidos_area_m2'] = [
            safe_float_convert(line) for line in traslucidos_area_lines] + [None]
        data_list['elementos_traslucidos_u_w_m2_k'] = [
            safe_float_convert(line) for line in traslucidos_U_lines] + [None]

        for key, lines in puentes_termicos_text.items():
            float_values = [safe_float_convert(line) for line in lines]
            data_list[key] = [None] + float_values + \
                [None]  # Pad first and last

        data_list['ua_phil'] = ua_phiL_values

        # Validate list lengths
        for key, lst in data_list.items():
            if len(lst) != num_orientations:
                logging.warning(
                    f"Length mismatch for {key} (Envolvente): expected {num_orientations}, got {len(lst)}. Padding.")
                data_list[key].extend([None] * (num_orientations - len(lst)))

        return data_list

    except (IndexError, ValueError, TypeError) as e:
        logging.error(
            f"Error processing Page 3 (Envolvente) dictionary: {e}", exc_info=True)
        return {}


def get_informe_cev_v2_pagina3_envolvente_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts envelope data from page 3 into a Pandas DataFrame."""
    data_dict_of_lists = get_informe_cev_v2_pagina3_envolvente_as_dict(
        pdf_report)
    if not data_dict_of_lists:
        return pd.DataFrame()
    try:
        df = pd.DataFrame(data_dict_of_lists)
        if "codigo_evaluacion" in df.columns:
            df = df.drop(columns=["codigo_evaluacion"])
        return df
    except ValueError as ve:
        logging.error(
            f"ValueError creating DataFrame for Page 3 Envolvente (likely unequal list lengths): {ve}", exc_info=True)
        return pd.DataFrame()
    except Exception as e:
        logging.error(
            f"Failed to convert page 3 envolvente dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()

# ------------------------------------------------------------------------------------------------------------
#  Pagina 4
# ------------------------------------------------------------------------------------------------------------


def get_informe_cev_v2_pagina4_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extracts monthly data from page 4 into a dictionary (structured for DataFrame).
    Uses safe float conversion and get_page_coordinates for consistency.
    """
    data_list: Dict[str, List[Any]] = {}
    try:
        if not isinstance(pdf_report, fitz.Document):
            raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 4:
            raise ValueError("PDF has less than 4 pages.")
        page = pdf_report[3]
        num_months = 12
        months = list(range(1, num_months + 1))

        # Usar get_page_coordinates para obtener las coordenadas
        COORDINATES = get_page_coordinates(3)

        codigo_evaluacion = extract_text_from_area(
            page, COORDINATES['codigo_evaluacion']).strip()
        data_list['codigo_evaluacion'] = [codigo_evaluacion] * num_months
        data_list['mes_id'] = months

        # Extraer datos mensuales usando las coordenadas generadas
        monthly_fields = [key for key in COORDINATES.keys()
                          if key.endswith('_mes_1')]
        base_fields = [key.replace('_mes_1', '') for key in monthly_fields]

        for base_field in base_fields:
            monthly_values = []
            for month in range(1, 13):
                field_key = f'{base_field}_mes_{month}'
                if field_key in COORDINATES:
                    text = extract_text_from_area(page, COORDINATES[field_key])
                    monthly_values.append(safe_float_convert(text))
                else:
                    monthly_values.append(None)
            data_list[base_field] = monthly_values

        # Validate list lengths
        for key, lst in data_list.items():
            if len(lst) != num_months:
                logging.warning(
                    f"Length mismatch for {key} (Page 4): expected {num_months}, got {len(lst)}. Padding.")
                data_list[key].extend([None] * (num_months - len(lst)))

        return data_list

    except (IndexError, ValueError, TypeError) as e:
        logging.error(
            f"Error processing Page 4 dictionary: {e}", exc_info=True)
        return {}


def get_informe_cev_v2_pagina4_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """Extracts monthly data from page 4 into a Pandas DataFrame."""
    data_dict_of_lists = get_informe_cev_v2_pagina4_as_dict(pdf_report)
    if not data_dict_of_lists:
        return pd.DataFrame()
    try:
        df = pd.DataFrame(data_dict_of_lists)
        if "codigo_evaluacion" in df.columns:
            df = df.drop(columns=["codigo_evaluacion"])

        # Convert mes_id to Spanish month names
        mes_mapping = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
                       7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
        if 'mes_id' in df.columns:
            df['mes'] = df['mes_id'].map(mes_mapping)
            df = df.drop(columns=['mes_id'])
            cols = ['mes'] + [col for col in df.columns if col != 'mes']
            df = df[cols]
        return df
    except ValueError as ve:
        logging.error(
            f"ValueError creating DataFrame for Page 4 (likely unequal list lengths): {ve}", exc_info=True)
        return pd.DataFrame()
    except Exception as e:
        logging.error(
            f"Failed to convert page 4 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()

# ------------------------------------------------------------------------------------------------------------
#  Pagina 5
# ------------------------------------------------------------------------------------------------------------


def get_informe_cev_v2_pagina5_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 5 of an informe_CEV_v2 PDF report and return it as a dictionary.
    Page 5 contains a transposed table with data for January and July across 10 energy parameters.
    Uses column-based extraction with specific logic for 19-value pattern.
    """
    data_list: Dict[str, List[Any]] = {}
    try:
        if not isinstance(pdf_report, fitz.Document):
            raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 5:
            raise ValueError("PDF has less than 5 pages.")

        page = pdf_report[4]

        # Usar get_page_coordinates para obtener las coordenadas
        COORDINATES = get_page_coordinates(4)

        if not COORDINATES:
            logging.warning("No coordinates defined for page 5")
            return {}

        # Extraer texto de columnas completas
        codigo_evaluacion = extract_text_from_area(
            page, COORDINATES['codigo_evaluacion']).strip()
        columna_enero_text = extract_text_from_area(
            page, COORDINATES['columna_enero'])
        columna_julio_text = extract_text_from_area(
            page, COORDINATES['columna_julio'])

        # Función para procesar una línea y convertir a float
        def convert_line_to_float(line: str) -> Optional[float]:
            """Convierte una línea de texto a float, manejando casos especiales."""
            if not line or line.strip() == '':
                return None

            cleaned_line = line.strip().replace(',', '.')

            # Casos especiales
            if cleaned_line in ['-', '-0']:
                return 0.0

            try:
                return float(cleaned_line)
            except ValueError:
                return None

        # Función para procesar una columna completa según el patrón identificado
        def process_column_data(column_text: str, is_enero: bool = False) -> List[Optional[float]]:
            """
            Procesa el texto de una columna completa según el patrón de 19 valores.

            Para Julio: tomar los últimos 10 valores
            Para Enero: tomar los últimos 10 valores, luego intercambiar penúltimo con antepenúltimo
            """
            if not column_text or column_text.strip() == '':
                return [None] * 10

            # Dividir por líneas y limpiar
            lines = [line.strip()
                     for line in column_text.splitlines() if line.strip()]

            # Convertir todas las líneas a float
            all_values = []
            for line in lines:
                value = convert_line_to_float(line)
                all_values.append(value)

            # Verificar que tenemos 19 valores (o al menos 10)
            if len(all_values) < 10:
                logging.warning(
                    f"Expected at least 10 values, got {len(all_values)}")
                # Rellenar con None si faltan valores
                while len(all_values) < 10:
                    all_values.append(None)

            # Tomar los últimos 10 valores
            last_10_values = all_values[-10:] if len(all_values) >= 10 else all_values + [
                None] * (10 - len(all_values))

            # Para enero: intercambiar penúltimo (índice -2) con antepenúltimo (índice -3)
            if is_enero and len(last_10_values) >= 3:
                # Intercambiar posiciones: penúltimo ↔ antepenúltimo
                last_10_values[-2], last_10_values[-3] = last_10_values[-3], last_10_values[-2]

            return last_10_values

        # Procesar columnas
        valores_enero = process_column_data(columna_enero_text, is_enero=True)
        valores_julio = process_column_data(columna_julio_text, is_enero=False)

        # Lista de parámetros energéticos en el orden esperado
        field_names = [
            'q_recuperado_kwh',
            'q_puentes_termicos_kwh',
            'q_contra_terreno_kwh',
            'q_piso_ventilado_kwh',
            'q_ventanas_kwh',
            'q_muros_kwh',
            'q_techo_kwh',
            'q_infiltraciones_kwh',
            'q_ventilacion_kwh',
            'q_sol_kwh'
        ]

        # Preparar estructura de datos para DataFrame (2 filas: Enero y Julio)
        data_list['codigo_evaluacion'] = [codigo_evaluacion, codigo_evaluacion]
        data_list['mes'] = ['Enero', 'Julio']

        # Asignar valores a cada parámetro energético
        for i, field_name in enumerate(field_names):
            enero_val = valores_enero[i] if i < len(valores_enero) else None
            julio_val = valores_julio[i] if i < len(valores_julio) else None
            data_list[field_name] = [enero_val, julio_val]

        return data_list

    except Exception as e:
        logging.error(f"Error accessing Page 5 dictionary: {e}", exc_info=True)
        return {}


def get_informe_cev_v2_pagina5_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """
    Extract data from page 5 into a Pandas DataFrame.
    Returns a DataFrame with 2 rows (Enero, Julio) and 10 energy parameter columns.
    Uses get_page_coordinates for consistency.
    """
    data_dict_of_lists = get_informe_cev_v2_pagina5_as_dict(pdf_report)
    if not data_dict_of_lists:
        return pd.DataFrame()
    try:
        # Crear DataFrame directamente desde el diccionario de listas
        df = pd.DataFrame(data_dict_of_lists)

        # Eliminar columna duplicada de codigo_evaluacion si existe
        if "codigo_evaluacion" in df.columns:
            df = df.drop(columns=["codigo_evaluacion"])

        # Reordenar columnas: mes primero, luego los parámetros energéticos
        if 'mes' in df.columns:
            energy_cols = [col for col in df.columns if col != 'mes']
            df = df[['mes'] + energy_cols]

        return df

    except Exception as e:
        logging.error(
            f"Failed to convert page 5 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()

# ------------------------------------------------------------------------------------------------------------
#  Pagina 6
# ------------------------------------------------------------------------------------------------------------


def get_informe_cev_v2_pagina6_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 6 of an informe_CEV_v2 PDF report and return it as a dictionary.
    Page 6 contains hourly temperature data for 4 months extracted using OCR.
    Uses get_page_coordinates for consistency.
    """
    data_list: Dict[str, List[Any]] = {}
    try:
        if not isinstance(pdf_report, fitz.Document):
            raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 6:
            raise ValueError("PDF has less than 6 pages.")

        page = pdf_report[5]

        # Usar get_page_coordinates para obtener las coordenadas
        COORDINATES = get_page_coordinates(5)

        if not COORDINATES:
            logging.warning("No coordinates defined for page 6")
            return {}

        # Extraer código de evaluación usando método estándar
        codigo_evaluacion = extract_text_from_area(
            page, COORDINATES['codigo_evaluacion']).strip()

        # Meses a procesar
        meses = ['enero', 'abril', 'julio', 'octubre']

        # Inicializar estructura de datos
        data_list['codigo_evaluacion'] = [codigo_evaluacion] * 24
        data_list['hora'] = list(range(1, 25))  # Horas 1 a 24

        # Procesar cada mes
        for mes in meses:
            if mes in COORDINATES:
                # Convertir coordenadas a fitz.Rect para OCR
                x1, y1, x2, y2 = COORDINATES[mes]

                # Normalizar coordenadas usando la misma lógica que extract_text_from_area
                REPORT_WIDTH = 215.9  # mm
                REPORT_HEIGHT = 330.0  # mm

                page_rect = page.rect
                page_width = page_rect.width
                page_height = page_rect.height

                # Normalizar coordenadas
                rx1, ry1 = normalize_coordinates(
                    x1, y1, REPORT_WIDTH, REPORT_HEIGHT, page_width, page_height)
                rx2, ry2 = normalize_coordinates(
                    x2, y2, REPORT_WIDTH, REPORT_HEIGHT, page_width, page_height)

                crop_box = fitz.Rect(rx1, ry1, rx2, ry2)

                # Extraer texto usando OCR
                raw_text = _get_text_from_image_area(page, crop_box)

                # Procesar texto y obtener temperaturas
                temp_data = _parse_hourly_temps_from_text_to_dict(
                    raw_text, mes)

                # Agregar datos al diccionario principal
                data_list[f't_ext_{mes}'] = temp_data[f't_ext_{mes}']
                data_list[f't_int_{mes}'] = temp_data[f't_int_{mes}']

                logging.info(
                    f"Procesado {mes}: {len(temp_data[f't_ext_{mes}'])} valores exteriores, {len(temp_data[f't_int_{mes}'])} valores interiores")
            else:
                # Si no hay coordenadas para el mes, rellenar con None
                data_list[f't_ext_{mes}'] = [None] * 24
                data_list[f't_int_{mes}'] = [None] * 24
                logging.warning(f"No se encontraron coordenadas para {mes}")

        # Validar que todas las listas tengan 24 elementos
        for key, lst in data_list.items():
            if len(lst) != 24:
                logging.warning(
                    f"Length mismatch for {key} (Page 6): expected 24, got {len(lst)}. Padding/truncating.")
                if len(lst) < 24:
                    lst.extend([None] * (24 - len(lst)))
                else:
                    data_list[key] = lst[:24]

        return data_list

    except Exception as e:
        logging.error(f"Error accessing Page 6 dictionary: {e}", exc_info=True)
        return {}


def get_informe_cev_v2_pagina6_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """
    Extract hourly temperature data from page 6 into a Pandas DataFrame.
    Returns a DataFrame with 24 rows (hours) and temperature columns for 4 months.
    Uses get_page_coordinates for consistency.
    """
    data_dict_of_lists = get_informe_cev_v2_pagina6_as_dict(pdf_report)
    if not data_dict_of_lists:
        return pd.DataFrame()
    try:
        # Crear DataFrame directamente desde el diccionario de listas
        df = pd.DataFrame(data_dict_of_lists)

        # Eliminar columna duplicada de codigo_evaluacion si existe
        if "codigo_evaluacion" in df.columns:
            df = df.drop(columns=["codigo_evaluacion"])

        # Reordenar columnas: hora primero, luego las temperaturas por mes
        if 'hora' in df.columns:
            temp_cols = [col for col in df.columns if col != 'hora']
            # Ordenar columnas de temperatura por mes
            ordered_temp_cols = []
            for mes in ['enero', 'abril', 'julio', 'octubre']:
                for temp_type in ['t_ext', 't_int']:
                    col_name = f'{temp_type}_{mes}'
                    if col_name in temp_cols:
                        ordered_temp_cols.append(col_name)

            df = df[['hora'] + ordered_temp_cols]

        return df

    except Exception as e:
        logging.error(
            f"Failed to convert page 6 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()


# ------------------------------------------------------------------------------------------------------------
#  Pagina 7
# ------------------------------------------------------------------------------------------------------------

def get_informe_cev_v2_pagina7_as_dict(pdf_report: fitz.Document) -> Dict[str, Any]:
    """
    Extract data from page 7 of an informe_CEV_v2 PDF report and return it as a dictionary.

    Args:
        pdf_report (fitz.Document): The PyMuPDF document object.

    Returns:
        Dict[str, Any]: A dictionary containing field names as keys and extracted text/data as values.
    """
    result: Dict[str, Any] = {}
    try:
        if not isinstance(pdf_report, fitz.Document):
            raise TypeError("Input must be a fitz.Document object.")
        if len(pdf_report) < 7:
            raise ValueError("PDF has less than 7 pages.")

        page = pdf_report[6]

        # Usar get_page_coordinates para obtener las coordenadas
        COORDINATES = get_page_coordinates(6)

        fields: Dict[str, str] = {k: extract_text_from_area(
            page, v) for k, v in COORDINATES.items()}

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
        logging.error(
            f"Error processing Page 7 dictionary: {e}", exc_info=True)
        return {}


def get_informe_cev_v2_pagina7_as_dataframe(pdf_report: fitz.Document) -> pd.DataFrame:
    """
    Extract evaluation details from page 7 into a Pandas DataFrame.

    Args:
        pdf_report (fitz.Document): The PyMuPDF document object.

    Returns:
        pd.DataFrame: A one-row DataFrame containing the extracted fields.
    """
    data_dict = get_informe_cev_v2_pagina7_as_dict(pdf_report)
    if not data_dict:
        return pd.DataFrame()
    try:
        df = pd.DataFrame.from_dict(data_dict, orient='index').T
        if "codigo_evaluacion" in df.columns:
            df = df.drop(columns=["codigo_evaluacion"])
        return df
    except Exception as e:
        logging.error(
            f"Failed to convert page 7 dict to DataFrame: {e}", exc_info=True)
        return pd.DataFrame()

# ------------------------------------------------------------------------------------------------------------
#  Función de ejemplo de uso
# ------------------------------------------------------------------------------------------------------------


def demo_scraping_workflow_with_contours(pdf_path: str) -> None:
    """
    Ejemplo completo de cómo usar todas las funciones del archivo scraping_functions.py
    """
    try:
        # Abrir PDF
        # pdf_path = "informe_cev_v2.pdf"
        pdf_doc = fitz.open(pdf_path)

        logging.info("=== EXTRACCIÓN DE DATOS ===")

        # Extraer datos de cada página
        logging.info("Extrayendo datos de página 1...")
        page1_data = get_informe_cev_v2_pagina1_as_dict(pdf_doc)
        page1_df = get_informe_cev_v2_pagina1_as_dataframe(pdf_doc)

        logging.info("Extrayendo datos de página 2...")
        page2_data = get_informe_cev_v2_pagina2_as_dict(pdf_doc)
        page2_df = get_informe_cev_v2_pagina2_as_dataframe(pdf_doc)

        logging.info("Extrayendo datos de página 3 (consumos)...")
        page3_consumos_data = get_informe_cev_v2_pagina3_consumos_as_dict(
            pdf_doc)
        page3_consumos_df = get_informe_cev_v2_pagina3_consumos_as_dataframe(
            pdf_doc)

        logging.info("Extrayendo datos de página 3 (envolvente)...")
        page3_envolvente_data = get_informe_cev_v2_pagina3_envolvente_as_dict(
            pdf_doc)
        page3_envolvente_df = get_informe_cev_v2_pagina3_envolvente_as_dataframe(
            pdf_doc)

        logging.info("Extrayendo datos de página 4...")
        page4_data = get_informe_cev_v2_pagina4_as_dict(pdf_doc)
        page4_df = get_informe_cev_v2_pagina4_as_dataframe(pdf_doc)

        logging.info("Extrayendo datos de página 5...")
        page5_data = get_informe_cev_v2_pagina5_as_dict(pdf_doc)
        page5_df = get_informe_cev_v2_pagina5_as_dataframe(pdf_doc)

        logging.info("Extrayendo datos de página 6...")
        page6_data = get_informe_cev_v2_pagina6_as_dict(pdf_doc)
        page6_df = get_informe_cev_v2_pagina6_as_dataframe(pdf_doc)

        logging.info("Extrayendo datos de página 7...")
        page7_data = get_informe_cev_v2_pagina7_as_dict(pdf_doc)
        page7_df = get_informe_cev_v2_pagina7_as_dataframe(pdf_doc)

        logging.info("\n=== VISUALIZACIÓN DE RECTÁNGULOS ===")

        # Dibujar rectángulos en todas las páginas
        logging.info("Dibujando rectángulos en todas las páginas...")
        draw_all_pages_rectangles(
            pdf_doc, "informe_completo_con_rectangulos.pdf")

        # Mostrar resumen de datos extraídos
        logging.info(f"\n=== RESUMEN ===")
        logging.info(
            f"Código de evaluación: {page1_data.get('codigo_evaluacion', 'N/A')}")
        logging.info(f"Región: {page1_data.get('region', 'N/A')}")
        logging.info(f"Comuna: {page1_data.get('comuna', 'N/A')}")
        logging.info(f"Tipo de vivienda: {page1_data.get('tipo_vivienda', 'N/A')}")
        logging.info(
            f"Superficie útil: {page1_data.get('superficie_interior_util_m2', 'N/A')} m²")
        logging.info(
            f"Porcentaje de ahorro: {page1_data.get('porcentaje_ahorro', 'N/A')}%")
        logging.info(
            f"Letra eficiencia: {page1_data.get('letra_eficiencia_energetica_dem', 'N/A')}")

        pdf_doc.close()
        logging.info("\nProcesamiento completado exitosamente!")

    except Exception as e:
        logging.error(f"Error en el ejemplo: {e}")
        logging.error(f"Error in example_usage_complete: {e}", exc_info=True)
