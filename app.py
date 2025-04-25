# --- app.py (Updated - Fix download button trigger issue) ---
import streamlit as st
import fitz # PyMuPDF
import pandas as pd
from typing import List, Tuple, Dict, Any, Optional
from io import BytesIO
import logging
import time
import re # Import regex for sanitizing sheet names
import os # Import os to check for sample file existence

# Import using wildcard as requested
# Make sure scraping_functions.py matches the version the user uploaded
from scraping_functions import *

# Configure logging for the app
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Constants for Examples ---
REPORT_EXAMPLES_FOLDER = "report_examples"
SAMPLE_PDF_PRECAL_NAME = "9_8_1_13e7982f-3470-5a8d-ba76-0a04e3ee0b7b.pdf"
SAMPLE_PDF_CAL_NAME = "10_192_2_b5ffa84b-d3b4-51c6-afa0-bc3eff628bfa.pdf"
SAMPLE_PDF_PRECAL_PATH = os.path.join(REPORT_EXAMPLES_FOLDER, SAMPLE_PDF_PRECAL_NAME)
SAMPLE_PDF_CAL_PATH = os.path.join(REPORT_EXAMPLES_FOLDER, SAMPLE_PDF_CAL_NAME)


# --- Renaming Dictionaries ---
RENAME_MAP_P1 = { "tipo_evaluacion": "Tipo de Evaluación", "codigo_evaluacion": "Código de Evaluación", "region": "Región", "comuna": "Comuna", "direccion": "Dirección", "rol_vivienda_proyecto": "Rol de Vivienda o Proyecto", "tipo_vivienda": "Tipo de Vivienda", "superficie_interior_util_m2": "Superficie Interior Útil (m²)", "porcentaje_ahorro": "Porcentaje de Ahorro (%)", "letra_eficiencia_energetica_dem": "Letra de Eficiencia Energética", "demanda_calefaccion_kwh_m2_ano": "Demanda Calefacción (kWh/m²/año)", "demanda_enfriamiento_kwh_m2_ano": "Demanda Enfriamiento (kWh/m²/año)", "demanda_total_kwh_m2_ano": "Demanda Total (kWh/m²/año)", "emitida_el": "Emitida el" }
RENAME_MAP_P2 = { 'region': 'Región', 'comuna': 'Comuna', 'direccion': 'Dirección', 'rol_vivienda': 'Rol Vivienda', 'tipo_vivienda': 'Tipo Vivienda', 'zona_termica': 'Zona Térmica', 'superficie_interior_util_m2': 'Superficie Interior Útil (m²)', 'solicitado_por': 'Solicitado Por', 'evaluado_por': 'Evaluado Por', 'codigo_evaluacion': 'Código Evaluación', 'demanda_calefaccion_kwh_m2_ano': 'Demanda Calefacción Promedio (kWh/m²/año)', 'demanda_enfriamiento_kwh_m2_ano': 'Demanda Enfriamiento Promedio (kWh/m²/año)', 'demanda_total_kwh_m2_ano': 'Demanda Total Promedio (kWh/m²/año)', 'demanda_total_bis_kwh_m2_ano': 'Demanda Total Vivienda Eval. (kWh/m²/año)', 'demanda_total_referencia_kwh_m2_ano': 'Demanda Total Referencia (kWh/m²/año)', 'porcentaje_ahorro': 'Porcentaje Ahorro (%)', 'muro_principal_descripcion': 'Muro Principal: Descripción', 'muro_principal_exigencia_W_m2_K': 'Muro Principal: Exigencia (W/m²K)', 'muro_secundario_descripcion': 'Muro Secundario: Descripción', 'muro_secundario_exigencia_W_m2_K': 'Muro Secundario: Exigencia (W/m²K)', 'piso_principal_descripcion': 'Piso Principal: Descripción', 'piso_principal_exigencia_W_m2_K': 'Piso Principal: Exigencia (W/m²K)', 'puerta_principal_descripcion': 'Puerta Principal: Descripción', 'puerta_principal_exigencia': 'Puerta Principal: Exigencia', 'techo_principal_descripcion': 'Techo Principal: Descripción', 'techo_principal_exigencia_W_m2_K': 'Techo Principal: Exigencia (W/m²K)', 'techo_secundario_descripcion': 'Techo Secundario: Descripción', 'techo_secundario_exigencia_W_m2_K': 'Techo Secundario: Exigencia (W/m²K)', 'superficie_vidriada_principal_descripcion': 'Sup. Vidriada Principal: Descripción', 'superficie_vidriada_principal_exigencia': 'Sup. Vidriada Principal: Exigencia', 'superficie_vidriada_secundaria_descripcion': 'Sup. Vidriada Secundaria: Descripción', 'superficie_vidriada_secundaria_exigencia': 'Sup. Vidriada Secundaria: Exigencia', 'ventilacion_rah_descripcion': 'Ventilación (RAH): Descripción', 'ventilacion_rah_exigencia': 'Ventilación (RAH): Exigencia', 'infiltraciones_rah_descripcion': 'Infiltraciones (RAH): Descripción', 'infiltraciones_rah_exigencia': 'Infiltraciones (RAH): Exigencia' }
RENAME_MAP_P3_CONSUMOS = { 'codigo_evaluacion': 'Código Evaluación', 'agua_caliente_sanitaria_kwh_m2': 'ACS (kWh/m²)', 'agua_caliente_sanitaria_perc': 'ACS (%)', 'iluminacion_kwh_m2': 'Iluminación (kWh/m²)', 'iluminacion_per': 'Iluminación (%)', 'calefaccion_kwh_m2': 'Calefacción (kWh/m²)', 'calefaccion_kwh_per': 'Calefacción (%)', 'energia_renovable_no_convencional_kwh_m2': 'ERNC (kWh/m²)', 'energia_renovable_no_convencional_per': 'ERNC (%)', 'consumo_total_kwh_m2': 'Consumo Total (kWh/m²)', 'emisiones_kgco2_m2_ano': 'Emisiones (kgCO₂e/m²/año)', 'calefaccion_descripcion_proy': 'Calefacción Proy.: Desc.', 'calefaccion_consumo_proy_kwh': 'Calefacción Proy. (kWh)', 'calefaccion_consumo_proy_per': 'Calefacción Proy. (%)', 'iluminacion_descripcion_proy': 'Iluminación Proy.: Desc.', 'iluminacion_consumo_proy_kwh': 'Iluminación Proy. (kWh)', 'iluminacion_consumo_proy_per': 'Iluminación Proy. (%)', 'agua_caliente_sanitaria_descripcion_proy': 'ACS Proy.: Desc.', 'agua_caliente_sanitaria_consumo_proy_kwh': 'ACS Proy. (kWh)', 'agua_caliente_sanitaria_consumo_proy_per': 'ACS Proy. (%)', 'energia_renovable_no_convencional_descripcion_proy': 'ERNC Proy.: Desc.', 'energia_renovable_no_convencional_consumo_proy_kwh': 'ERNC Proy. (kWh)', 'energia_renovable_no_convencional_consumo_proy_per': 'ERNC Proy. (%)', 'consumo_total_requerido_proy_kwh': 'Consumo Total Proy. (kWh)', 'calefaccion_descripcion_ref': 'Calefacción Ref.: Desc.', 'calefaccion_consumo_ref_kwh': 'Calefacción Ref. (kWh)', 'calefaccion_consumo_ref_per': 'Calefacción Ref. (%)', 'iluminacion_descripcion_ref': 'Iluminación Ref.: Desc.', 'iluminacion_consumo_ref_kwh': 'Iluminación Ref. (kWh)', 'iluminacion_consumo_ref_per': 'Iluminación Ref. (%)', 'agua_caliente_sanitaria_descripcion_ref': 'ACS Ref.: Desc.', 'agua_caliente_sanitaria_consumo_ref_kwh': 'ACS Ref. (kWh)', 'agua_caliente_sanitaria_consumo_ref_per': 'ACS Ref. (%)', 'energia_renovable_no_convencional_descripcion_ref': 'ERNC Ref.: Desc.', 'energia_renovable_no_convencional_consumo_ref_kwh': 'ERNC Ref. (kWh)', 'energia_renovable_no_convencional_consumo_ref_per': 'ERNC Ref. (%)', 'consumo_total_requerido_ref_kwh': 'Consumo Total Ref. (kWh)', 'consumo_ep_calefaccion_kwh': 'EP Consumo Calef. (kWh)', 'consumo_ep_agua_caliente_sanitaria_kwh': 'EP Consumo ACS (kWh)', 'consumo_ep_iluminacion_kwh': 'EP Consumo Ilum. (kWh)', 'consumo_ep_ventiladores_kwh': 'EP Consumo Vent. (kWh)', 'generacion_ep_fotovoltaicos_kwh': 'EP Gen. FV (kWh)', 'aporte_fotovoltaicos_consumos_basicos_kwh': 'EP Aporte FV (kWh)', 'diferencia_fotovoltaica_para_consumo_kwh': 'EP Dif. FV (kWh)', 'aporte_solar_termica_consumos_basicos_kwh': 'EP Aporte Solar T. (kWh)', 'aporte_solar_termica_agua_caliente_sanitaria_kwh': 'EP Aporte Solar T. ACS (kWh)', 'total_consumo_ep_antes_fotovoltaica_kwh': 'EP Total Antes FV (kWh)', 'aporte_fotovoltaicos_consumos_basicos_kwh_bis': 'EP Aporte FV Bis (kWh)', 'consumos_basicos_a_suplir_kwh': 'EP Consumos a Suplir (kWh)', 'consumo_total_ep_obj_kwh': 'Consumo Total EP Obj (kWh)', 'consumo_total_ep_ref_kwh': 'Consumo Total EP Ref (kWh)', 'coeficiente_energetico_c': 'Coeficiente Energético (C)' }
RENAME_MAP_P3_ENVOLVENTE = { 'codigo_evaluacion': 'Código Evaluación', 'orientacion': 'Orientación', 'elementos_opacos_area_m2': 'Opacos: Área (m²)', 'elementos_opacos_U_W_m2_K': 'Opacos: U (W/m²K)', 'elementos_traslucidos_area_m2': 'Traslúcidos: Área (m²)', 'elementos_traslucidos_U_W_m2_K': 'Traslúcidos: U (W/m²K)', 'P01_W_K': 'PT P01 (W/K)', 'P02_W_K': 'PT P02 (W/K)', 'P03_W_K': 'PT P03 (W/K)', 'P04_W_K': 'PT P04 (W/K)', 'P05_W_K': 'PT P05 (W/K)', 'UA_phiL': 'Ht (UA + φL) (W/K)' }
RENAME_MAP_P4 = { 'codigo_evaluacion': 'Código Evaluación', 'mes': 'Mes', 'demanda_calef_viv_eval_kwh': 'Dem. Calef. Eval. (kWh)', 'demanda_calef_viv_ref_kwh': 'Dem. Calef. Ref. (kWh)', 'demanda_enfri_viv_eval_kwh': 'Dem. Enfri. Eval. (kWh)', 'demanda_enfri_viv_ref_kwh': 'Dem. Enfri. Ref. (kWh)', 'sobrecalentamiento_viv_eval_hr': 'Sobrecalent. Eval. (hr)', 'sobrecalentamiento_viv_ref_hr': 'Sobrecalent. Ref. (hr)', 'sobreenfriamiento_viv_eval_hr': 'Sobreenfri. Eval. (hr)', 'sobreenfriamiento_viv_ref_hr': 'Sobreenfri. Ref. (hr)' }
RENAME_MAP_P5_P6 = { 'codigo_evaluacion': 'Código Evaluación', 'content_note': 'Nota' }
RENAME_MAP_P7 = { 'codigo_evaluacion': 'Código Evaluación', 'mandante_nombre': 'Mandante: Nombre', 'mandante_rut': 'Mandante: RUT', 'evaluador_nombre': 'Evaluador: Nombre', 'evaluador_rut': 'Evaluador: RUT', 'evaluador_rol_minvu': 'Evaluador: Rol MINVU' }

# List of rename maps in the order of extracted_dfs
ALL_RENAME_MAPS = [
    RENAME_MAP_P1, RENAME_MAP_P2, RENAME_MAP_P3_CONSUMOS, RENAME_MAP_P3_ENVOLVENTE,
    RENAME_MAP_P4, RENAME_MAP_P5_P6, RENAME_MAP_P5_P6, RENAME_MAP_P7
]

# --- Initialize Session State ---
if 'uploaded_file_bytes' not in st.session_state: st.session_state.uploaded_file_bytes = None
if 'extracted_data' not in st.session_state: st.session_state.extracted_data = None
if 'processing_done' not in st.session_state: st.session_state.processing_done = False
if 'file_name' not in st.session_state: st.session_state.file_name = None
if 'last_uploaded_file_id' not in st.session_state: st.session_state.last_uploaded_file_id = None

# --- Validation Function ---
def is_valid_cev_v2_pdf(pdf_doc: fitz.Document) -> bool:
    if not pdf_doc or len(pdf_doc) < 7: return False
    try:
        text_upper = pdf_doc[0].get_text("text").upper()
        valid = "PRECALIFICACIÓN ENERGÉTICA" in text_upper or "CALIFICACIÓN ENERGÉTICA" in text_upper
        if valid: logging.info("Validation successful.")
        else: logging.warning("Validation failed: Keywords not found.")
        return valid
    except Exception as e:
        logging.error(f"Validation error: {e}", exc_info=True)
        return False

# --- Display Functions ---
def display_dataframe_with_title(title: str, data: pd.DataFrame, transpose: bool = False, rename_map: Optional[Dict[str, str]] = None):
    st.header(title)
    if data is None or data.empty: st.warning(f"No data available for this section."); return
    data_to_display = data.copy()
    if rename_map:
        cols_to_rename = {k: v for k, v in rename_map.items() if k in data_to_display.columns}
        data_to_display = data_to_display.rename(columns=cols_to_rename)

    is_placeholder = 'content_note' in data.columns
    placeholder_note_col = rename_map.get('content_note', 'content_note') if rename_map else 'content_note'

    if is_placeholder and placeholder_note_col in data_to_display.columns:
         note = data_to_display[placeholder_note_col].iloc[0]
         if pd.notna(note): st.info(note)
         if transpose:
              cols_to_drop = [placeholder_note_col]; display_data = data_to_display.drop(columns=cols_to_drop, errors='ignore').T
              if not display_data.empty:
                  display_data.columns = [""]; dynamic_column_config = {col: st.column_config.Column(width=None) for col in display_data.columns}
                  st.dataframe(display_data, column_config=dynamic_column_config)
         return

    if transpose: display_data = data_to_display.T; display_data.columns = [""]
    else: display_data = data_to_display
    dynamic_column_config = {col: st.column_config.Column(width=None) for col in display_data.columns}
    st.dataframe(display_data, column_config=dynamic_column_config)


# --- PDF Processing (Accepts container for progress) ---
def process_pdf(pdf_document: fitz.Document, filename: str, progress_container) -> Tuple[List[pd.DataFrame], List[str]]:
    extracted_data_frames: List[pd.DataFrame] = []
    base_names: List[str] = []
    processing_steps = [
        (get_informe_cev_v2_pagina1_as_dataframe, "Página 1"), (get_informe_cev_v2_pagina2_as_dataframe, "Página 2"),
        (get_informe_cev_v2_pagina3_consumos_as_dataframe, "Página 3 - Consumos"), (get_informe_cev_v2_pagina3_envolvente_as_dataframe, "Página 3 - Envolvente"),
        (get_informe_cev_v2_pagina4_as_dataframe, "Página 4"), (get_informe_cev_v2_pagina5_as_dataframe, "Página 5"),
        (get_informe_cev_v2_pagina6_as_dataframe, "Página 6"), (get_informe_cev_v2_pagina7_as_dataframe, "Página 7"),
    ]
    progress_bar = progress_container.progress(0)
    status_text = progress_container.empty()
    total_steps = len(processing_steps)
    all_successful = True
    for i, (func, base_name) in enumerate(processing_steps):
        status_text.text(f"Procesando: {base_name}...")
        try:
            df = func(pdf_document)
            extracted_data_frames.append(df); base_names.append(base_name)
            if df.empty: logging.warning(f"Processed {base_name} but resulting DataFrame is empty.")
            else: logging.info(f"Processed {base_name}")
        except Exception as e:
            logging.error(f"Error processing {base_name}: {e}", exc_info=True)
            progress_container.warning(f"Error al extraer datos de '{base_name}'.")
            extracted_data_frames.append(pd.DataFrame()); base_names.append(base_name)
            all_successful = False
        progress_bar.progress((i + 1) / total_steps)

    if all_successful: status_text.success("Procesamiento completado."); time.sleep(2)
    else: status_text.warning("Procesamiento completado con errores."); time.sleep(3)
    status_text.empty()
    return extracted_data_frames, base_names

# --- Function to reset state ---
def reset_state():
    st.session_state.uploaded_file_bytes = None; st.session_state.extracted_data = None
    st.session_state.processing_done = False; st.session_state.file_name = None
    st.session_state.last_uploaded_file_id = None # Reset this too

# --- Function to create Excel File ---
def create_multisheet_excel(dataframes_list: List[pd.DataFrame], sheet_names: List[str], rename_maps: List[Optional[Dict[str, str]]]) -> BytesIO:
    # ... (function remains the same) ...
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        if len(dataframes_list) != len(sheet_names) or len(dataframes_list) != len(rename_maps): logging.error("Mismatch for Excel generation."); return excel_buffer
        for i, df_original in enumerate(dataframes_list):
            if df_original is None or df_original.empty: logging.info(f"Skipping empty df for sheet '{sheet_names[i]}'"); continue
            df_to_write = df_original.copy()
            if 'content_note' in df_to_write.columns: df_to_write = df_to_write.drop(columns=['content_note'], errors='ignore')
            rename_map = rename_maps[i]
            if rename_map:
                current_rename_map = {k:v for k,v in rename_map.items() if k != 'content_note'}; cols_to_rename = {k: v for k, v in current_rename_map.items() if k in df_to_write.columns}
                df_to_write = df_to_write.rename(columns=cols_to_rename)
            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', sheet_names[i])[:31]
            try: df_to_write.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            except Exception as e: logging.error(f"Error writing sheet '{safe_sheet_name}': {e}", exc_info=True)
    excel_buffer.seek(0)
    return excel_buffer

# --- Helper Function to Trigger Processing ---
def trigger_processing(file_bytes, filename, placeholder):
    """Stores file details and runs the processing pipeline."""
    # Set state for the file being processed
    st.session_state.uploaded_file_bytes = file_bytes
    st.session_state.file_name = filename
    st.session_state.last_uploaded_file_id = f"{filename}_{time.time()}"
    st.session_state.processing_done = False # Mark as not done before starting
    st.session_state.extracted_data = None   # Clear previous data

    try:
        with fitz.open(stream=file_bytes, filetype="pdf") as pdf_doc:
            placeholder.info("Validando archivo...")
            time.sleep(0.5)
            if not is_valid_cev_v2_pdf(pdf_doc):
                st.error(f"Archivo '{filename}' inválido o no soportado.")
                reset_state(); placeholder.empty()
                return # Stop processing
            else:
                placeholder.info(f"Archivo válido. Procesando '{filename}'...")
                extracted_dfs, base_names = process_pdf(pdf_doc, filename, placeholder)
                st.session_state.extracted_data = extracted_dfs
                st.session_state.processing_done = True # Mark as done only on success/partial success
                # Message handled inside process_pdf
    except Exception as e:
        error_message=str(e).lower(); user_msg=f"Error inesperado procesando {filename}: {e}"
        if any(err in error_message for err in ["cannot open", "damaged", "format error", "no objects found"]): user_msg=f"Error al leer PDF {filename}. Podría estar dañado/protegido."
        log_msg=f"Error processing file {filename}: {e}"; logging.error(log_msg, exc_info=True); st.error(user_msg); reset_state(); placeholder.empty()


# --- Main Application ---
def main():
    st.set_page_config(layout="centered")
    st.title("Informe CEV v2 (PDF scraper)")

    # --- Download Button (Conditional) ---
    if st.session_state.processing_done and st.session_state.extracted_data:
        excel_sheet_names = ["Pagina1", "Pagina2", "Pagina3_Consumos", "Pagina3_Envolvente", "Pagina4", "Pagina5", "Pagina6", "Pagina7"]
        if len(ALL_RENAME_MAPS) == len(st.session_state.extracted_data):
             excel_data = create_multisheet_excel(st.session_state.extracted_data, excel_sheet_names, ALL_RENAME_MAPS)
             download_filename = f"{st.session_state.file_name.replace('.pdf', '')}_Extracted_Data.xlsx" if st.session_state.file_name else "Extracted_Data.xlsx"
             st.download_button(label="Descargar Informe CEV (Excel)", data=excel_data, file_name=download_filename, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key="download_excel_all")
        else: st.warning("No se pudo preparar el archivo Excel (error interno de mapeo).")

    # --- Define Tabs ---
    tab_titles = ["Subir Archivo PDF", "Página 1", "Página 2", "Página 3", "Página 4", "Página 5", "Página 6", "Página 7"]
    tab_upload, tab_p1, tab_p2, tab_p3, tab_p4, tab_p5, tab_p6, tab_p7 = st.tabs(tab_titles)

    # --- Upload Tab Content ---
    with tab_upload:
        st.header("Cargar Archivo PDF")
        st.info("Cargue un archivo PDF en formato 'Informe CEV v2' o pruebe con uno de los ejemplos.")

        # --- File Uploader ---
        uploaded_file_widget = st.file_uploader("Seleccione su archivo PDF", accept_multiple_files=False, type="pdf", key="pdf_uploader", help="Suba un archivo PDF en formato 'Informe CEV v2'.")

        # --- Processing Placeholder ---
        # Define the placeholder HERE, before examples and processing logic
        processing_placeholder = st.empty()

        # --- Trigger processing for USER UPLOAD ---
        if uploaded_file_widget is not None:
            current_file_id = uploaded_file_widget.file_id
            # Check if it's a genuinely new file upload compared to the last processed ID
            if current_file_id != st.session_state.last_uploaded_file_id:
                 trigger_processing(uploaded_file_widget.getvalue(), uploaded_file_widget.name, processing_placeholder)
                 # No rerun here, let state change handle it

        st.markdown("---")
        st.subheader("O prueba con un ejemplo:")

        # --- Examples Section ---
        col1, col2 = st.columns(2)

        # Pre-read sample bytes for download buttons to avoid re-reading on every run
        sample1_bytes = None
        sample2_bytes = None
        try:
            if os.path.exists(SAMPLE_PDF_PRECAL_PATH):
                with open(SAMPLE_PDF_PRECAL_PATH, "rb") as f:
                    sample1_bytes = f.read()
        except Exception as e:
            logging.error(f"Failed to pre-read sample 1: {e}")
        try:
            if os.path.exists(SAMPLE_PDF_CAL_PATH):
                with open(SAMPLE_PDF_CAL_PATH, "rb") as f:
                    sample2_bytes = f.read()
        except Exception as e:
            logging.error(f"Failed to pre-read sample 2: {e}")

        # Example 1: Precalificación
        with col1:
            st.markdown("**Ejemplo Precalificación**")
            if sample1_bytes: # Only show buttons if file was successfully read
                if st.button("Cargar Ejemplo Precalificación", type="primary", key="load_precal", help=f"Procesa el archivo {SAMPLE_PDF_PRECAL_NAME}"):
                    trigger_processing(sample1_bytes, SAMPLE_PDF_PRECAL_NAME, processing_placeholder)
                st.download_button(label="Descargar PDF Ejemplo (Precal.)", data=sample1_bytes, file_name=SAMPLE_PDF_PRECAL_NAME, mime="application/pdf", key="download_precal")
            else:
                st.caption(f"Archivo '{SAMPLE_PDF_PRECAL_NAME}' no encontrado en '{REPORT_EXAMPLES_FOLDER}/'.")

        # Example 2: Calificación
        with col2:
            st.markdown("**Ejemplo Calificación**")
            if sample2_bytes: # Only show buttons if file was successfully read
                if st.button("Cargar Ejemplo Calificación", type="primary", key="load_cal", help=f"Procesa el archivo {SAMPLE_PDF_CAL_NAME}"):
                    trigger_processing(sample2_bytes, SAMPLE_PDF_CAL_NAME, processing_placeholder)
                st.download_button(label="Descargar PDF Ejemplo (Calif.)", data=sample2_bytes, file_name=SAMPLE_PDF_CAL_NAME, mime="application/pdf", key="download_cal")
            else:
                 st.caption(f"Archivo '{SAMPLE_PDF_CAL_NAME}' no encontrado en '{REPORT_EXAMPLES_FOLDER}/'.")

        # Display final status message AFTER buttons and upload logic
        # This message shows the state AFTER the current run's actions
        st.markdown("---") # Separator before final status/prompt
        if st.session_state.processing_done and st.session_state.file_name:
            processing_placeholder.success(f"Resultados listos para '{st.session_state.file_name}'.")
        elif st.session_state.file_name and not st.session_state.processing_done:
             # Error message was already shown by trigger_processing
             pass # Avoid showing redundant warnings here
        elif not st.session_state.file_name:
             processing_placeholder.empty() # Clear placeholder if no file selected/processed
             st.info("Por favor, seleccione un archivo PDF o cargue un ejemplo.")


        # --- Links Section ---
        st.markdown("---"); st.subheader("Recursos Adicionales")
        st.markdown("""
        Para más información sobre la Calificación Energética de Viviendas en Chile o para buscar informes, visite:
        * [Portal Oficial CEV](https://www.calificacionenergetica.cl/)
        * [Buscador Público de Viviendas Calificadas](https://calificacionenergeticaweb.minvu.cl/Publico/BusquedaVivienda.aspx)
        """, unsafe_allow_html=True)

    # --- Data Tab Content ---
    data_tabs = [tab_p1, tab_p2, tab_p3, tab_p4, tab_p5, tab_p6, tab_p7]
    data_structure = {
         0: ("Información General", 0, True, RENAME_MAP_P1),
         1: ("Información Gral. / Demanda / Diseño Arq.", 1, True, RENAME_MAP_P2),
         2: [("Consumo Energético Estimado", 2, True, RENAME_MAP_P3_CONSUMOS),
             ("Resumen Envolvente", 3, False, RENAME_MAP_P3_ENVOLVENTE)],
         3: ("Demanda / Sobrecalentamiento / Sobreenfriamiento Mensual", 4, False, RENAME_MAP_P4),
         4: ("Página 5 (Gráficos)", 5, True, RENAME_MAP_P5_P6),
         5: ("Página 6 (Gráficos)", 6, True, RENAME_MAP_P5_P6),
         6: ("Antecedentes de la Evaluación", 7, True, RENAME_MAP_P7)
    }

    for i, tab in enumerate(data_tabs):
        with tab:
            if st.session_state.processing_done and st.session_state.extracted_data and len(st.session_state.extracted_data) == 8:
                if st.session_state.file_name: st.caption(f"Mostrando resultados para: {st.session_state.file_name}")
                structure_info = data_structure[i]
                if isinstance(structure_info, list): # Handle combined Page 3
                    part1 = structure_info[0]; part2 = structure_info[1]
                    display_dataframe_with_title(part1[0], st.session_state.extracted_data[part1[1]], transpose=part1[2], rename_map=part1[3])
                    st.markdown("---")
                    display_dataframe_with_title(part2[0], st.session_state.extracted_data[part2[1]], transpose=part2[2], rename_map=part2[3])
                else: # Handle single page tabs
                    title, df_index, transpose_flag, rename_map_dict = structure_info
                    current_rename_map = rename_map_dict if rename_map_dict else None
                    display_dataframe_with_title(title, st.session_state.extracted_data[df_index], transpose=transpose_flag, rename_map=current_rename_map)
            else:
                 st.info("Suba o cargue un archivo PDF válido en la pestaña 'Subir Archivo PDF' para ver los datos.")

if __name__ == "__main__":
    main()