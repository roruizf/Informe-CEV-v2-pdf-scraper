import streamlit as st
import fitz # PyMuPDF
import pandas as pd
from typing import List, Tuple, Dict, Any, Optional
from io import BytesIO
import logging
import time
import re # Import regex for sanitizing sheet names

# Import using wildcard as requested
from scraping_functions import *

# Configure logging for the app
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

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
    if not pdf_doc or len(pdf_doc) != 7: return False
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
    """Displays a dataframe with title, optional rename, optional transpose."""
    st.header(title)
    if data is None or data.empty:
        st.warning(f"No data available for this section.")
        return

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

    if transpose:
        display_data = data_to_display.T; display_data.columns = [""]
    else:
        display_data = data_to_display

    dynamic_column_config = {col: st.column_config.Column(width=None) for col in display_data.columns}
    st.dataframe(display_data, column_config=dynamic_column_config)


# --- PDF Processing ---
def process_pdf(pdf_document: fitz.Document, filename: str) -> Tuple[List[pd.DataFrame], List[str]]:
    # ... (processing logic remains the same) ...
    extracted_data_frames: List[pd.DataFrame] = []
    base_names: List[str] = []
    processing_steps = [
        (get_informe_cev_v2_pagina1_as_dataframe, "Página 1"),
        (get_informe_cev_v2_pagina2_as_dataframe, "Página 2"),
        (get_informe_cev_v2_pagina3_consumos_as_dataframe, "Página 3 - Consumos"),
        (get_informe_cev_v2_pagina3_envolvente_as_dataframe, "Página 3 - Envolvente"),
        (get_informe_cev_v2_pagina4_as_dataframe, "Página 4"),
        (get_informe_cev_v2_pagina5_as_dataframe, "Página 5"),
        (get_informe_cev_v2_pagina6_as_dataframe, "Página 6"),
        (get_informe_cev_v2_pagina7_as_dataframe, "Página 7"),
    ]
    progress_container = st.container()
    progress_bar = progress_container.progress(0)
    status_text = progress_container.empty()
    total_steps = len(processing_steps)
    for i, (func, base_name) in enumerate(processing_steps):
        status_text.text(f"Procesando: {base_name}...")
        try: df = func(pdf_document); extracted_data_frames.append(df); base_names.append(base_name); logging.info(f"Processed {base_name}")
        except Exception as e: logging.error(f"Error processing {base_name}: {e}", exc_info=True); progress_container.warning(f"Error extracting '{base_name}'."); extracted_data_frames.append(pd.DataFrame()); base_names.append(base_name)
        progress_bar.progress((i + 1) / total_steps)
    status_text.success("Procesamiento completado.")
    return extracted_data_frames, base_names

# --- Function to reset state ---
def reset_state():
    st.session_state.uploaded_file_bytes = None
    st.session_state.extracted_data = None
    st.session_state.processing_done = False
    st.session_state.file_name = None
    st.session_state.last_uploaded_file_id = None

# --- Function to create Excel File ---
def create_multisheet_excel(dataframes_list: List[pd.DataFrame], sheet_names: List[str], rename_maps: List[Optional[Dict[str, str]]]) -> BytesIO:
    """Creates a multi-sheet Excel file from a list of dataframes."""
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        if len(dataframes_list) != len(sheet_names) or len(dataframes_list) != len(rename_maps):
             logging.error("Mismatch between dataframes, sheet names, or rename maps for Excel generation.")
             # Write an error sheet? Or just return empty buffer? For now, return empty.
             return excel_buffer # Consider adding an error sheet

        for i, df_original in enumerate(dataframes_list):
            if df_original is None or df_original.empty:
                logging.info(f"Skipping empty dataframe for sheet '{sheet_names[i]}'")
                continue # Skip empty dataframes

            df_to_write = df_original.copy()
            rename_map = rename_maps[i]

            # Apply renaming if map exists
            if rename_map:
                cols_to_rename = {k: v for k, v in rename_map.items() if k in df_to_write.columns}
                df_to_write = df_to_write.rename(columns=cols_to_rename)

            # Sanitize sheet name (max 31 chars, no invalid chars like / ? * [ ] :)
            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', sheet_names[i])[:31]

            try:
                df_to_write.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            except Exception as e:
                 logging.error(f"Error writing sheet '{safe_sheet_name}': {e}", exc_info=True)
                 # Optionally write a placeholder sheet indicating error

    excel_buffer.seek(0)
    return excel_buffer

# --- Main Application ---
def main():
    st.set_page_config(layout="centered")
    st.title("Informe CEV v2 (PDF scraper)")

    # --- Download Button (Conditional) ---
    if st.session_state.processing_done and st.session_state.extracted_data:
        excel_sheet_names = [
            "Pagina1", "Pagina2", "Pagina3_Consumos", "Pagina3_Envolvente",
            "Pagina4", "Pagina5", "Pagina6", "Pagina7"
        ]
        # Ensure we have the correct number of rename maps
        if len(ALL_RENAME_MAPS) == len(st.session_state.extracted_data):
             excel_data = create_multisheet_excel(
                 st.session_state.extracted_data,
                 excel_sheet_names,
                 ALL_RENAME_MAPS
             )
             download_filename = f"{st.session_state.file_name.replace('.pdf', '')}_Extracted_Data.xlsx" if st.session_state.file_name else "Extracted_Data.xlsx"

             st.download_button(
                 label="Descargar Informe CEV (Excel)",
                 data=excel_data,
                 file_name=download_filename,
                 mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                 key="download_excel_all"
             )
        else:
            st.warning("No se pudo preparar el archivo Excel (error interno de mapeo).")


    # --- Define Tabs ---
    tab_titles = ["Subir Archivo PDF", "Página 1", "Página 2", "Página 3", "Página 4", "Página 5", "Página 6", "Página 7"]
    tab_upload, tab_p1, tab_p2, tab_p3, tab_p4, tab_p5, tab_p6, tab_p7 = st.tabs(tab_titles)

    # --- Upload Tab Content ---
    with tab_upload:
        st.header("Cargar Archivo PDF")
        uploaded_file_widget = st.file_uploader("Seleccione un archivo PDF (informe_cev_v2)", accept_multiple_files=False, type="pdf", key="pdf_uploader", help="Suba un archivo PDF en formato 'Informe CEV v2'.")

        if uploaded_file_widget is not None:
            current_file_id = uploaded_file_widget.file_id
            if current_file_id != st.session_state.last_uploaded_file_id:
                st.info(f"Nuevo archivo detectado: '{uploaded_file_widget.name}'. Procesando...")
                st.session_state.uploaded_file_bytes = uploaded_file_widget.getvalue(); st.session_state.file_name = uploaded_file_widget.name; st.session_state.last_uploaded_file_id = current_file_id; st.session_state.processing_done = False; st.session_state.extracted_data = None
                processing_placeholder = st.empty()
                try:
                    with fitz.open(stream=st.session_state.uploaded_file_bytes, filetype="pdf") as pdf_doc:
                        processing_placeholder.info("Validando archivo...")
                        if not is_valid_cev_v2_pdf(pdf_doc):
                            st.error(f"Archivo '{st.session_state.file_name}' inválido o no soportado."); reset_state(); processing_placeholder.empty()
                        else:
                            processing_placeholder.info(f"Archivo válido. Procesando '{st.session_state.file_name}'...")
                            extracted_dfs, base_names = process_pdf(pdf_doc, st.session_state.file_name)
                            st.session_state.extracted_data = extracted_dfs; st.session_state.processing_done = True; processing_placeholder.empty(); st.success(f"Procesamiento de '{st.session_state.file_name}' completado.")
                            # Use st.rerun() to refresh the interface cleanly after processing
                            st.rerun()
                except Exception as e:
                    error_message=str(e).lower(); user_msg=f"Error inesperado procesando {st.session_state.file_name}: {e}"
                    if any(err in error_message for err in ["cannot open", "damaged", "format error", "no objects found"]): user_msg=f"Error al leer PDF {st.session_state.file_name}. Podría estar dañado/protegido."
                    log_msg=f"Error processing file {st.session_state.file_name}: {e}"; logging.error(log_msg, exc_info=True); st.error(user_msg); reset_state(); processing_placeholder.empty()
            elif st.session_state.processing_done: st.success(f"Archivo '{st.session_state.file_name}' procesado correctamente.")
            elif st.session_state.file_name: st.warning(f"Archivo '{st.session_state.file_name}' cargado, pero hubo error.")
        else:
            if st.session_state.last_uploaded_file_id is not None: reset_state()
            if not st.session_state.processing_done: st.info("Por favor, seleccione un archivo PDF.")
       
        # --- Links Section ---
        st.markdown("---"); st.subheader("Recursos Adicionales")
        st.markdown("""
        Para más información sobre la Calificación Energética de Viviendas en Chile o para buscar informes, visite:
        * [Portal Oficial CEV](https://www.calificacionenergetica.cl/)
        * [Buscador Público de Viviendas Calificadas](https://calificacionenergeticaweb.minvu.cl/Publico/BusquedaVivienda.aspx)
        """, unsafe_allow_html=True)

    # --- Data Tab Content ---
    data_tabs = [tab_p1, tab_p2, tab_p3, tab_p4, tab_p5, tab_p6, tab_p7]
    # Structure: Tab Index -> (Title | [(Title, DF Index, Transpose Flag, Rename Map), ...], DF Index, Transpose Flag, Rename Map)
    data_structure = {
         0: ("Información General", 0, True, RENAME_MAP_P1),                      # P1
         1: ("Información Gral. / Demanda energética / Diseño de Arquitectura", 1, True, RENAME_MAP_P2),                     # P2
         2: [("Consumo energético estimado ARQUITECTURA + EQUIPOS + TIPO DE ENERGÍA", 2, True, RENAME_MAP_P3_CONSUMOS),             # P3 (Part 1)
             ("Resumen Envolvente", 3, False, RENAME_MAP_P3_ENVOLVENTE)],       # P3 (Part 2)
         3: ("Demanda, sobrecalentamiento y sobreenfriamiento mensual", 4, False, RENAME_MAP_P4),              # P4
         4: ("Página 5 (No disponible)", 5, True, RENAME_MAP_P5_P6),                # P5 (Placeholder) - Transpose=True
         5: ("Página 6 (No disponible)", 6, True, RENAME_MAP_P5_P6),                # P6 (Placeholder) - Transpose=True
         6: ("Antecedentes de la evaluación", 7, True, RENAME_MAP_P7)                      # P7
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
                    # Check if rename_map_dict is None before passing, although it shouldn't be based on data_structure
                    current_rename_map = rename_map_dict if rename_map_dict else None
                    display_dataframe_with_title(title, st.session_state.extracted_data[df_index], transpose=transpose_flag, rename_map=current_rename_map)
            else:
                 st.info("Suba un archivo PDF válido en la pestaña 'Subir Archivo PDF' para ver los datos.")

if __name__ == "__main__":
    main()