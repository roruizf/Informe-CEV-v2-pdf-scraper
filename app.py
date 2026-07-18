import streamlit as st
import fitz
import pandas as pd
from typing import List, Tuple, Dict, Optional
from io import BytesIO
import logging
import os

from scraping_functions import (
    get_informe_cev_v2_pagina1_as_dataframe,
    get_informe_cev_v2_pagina2_as_dataframe,
    get_informe_cev_v2_pagina3_consumos_as_dataframe,
    get_informe_cev_v2_pagina3_envolvente_as_dataframe,
    get_informe_cev_v2_pagina4_as_dataframe,
    get_informe_cev_v2_pagina5_as_dataframe,
    get_informe_cev_v2_pagina6_as_dataframe,
    get_informe_cev_v2_pagina7_as_dataframe,
)

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

REPORT_EXAMPLES_FOLDER = "report_examples"

RENAME_MAP_P1 = {"tipo_evaluacion": "Tipo de Evaluación", "codigo_evaluacion": "Código de Evaluación", "region": "Región", "comuna": "Comuna", "direccion": "Dirección", "rol_vivienda_proyecto": "Rol de Vivienda o Proyecto", "tipo_vivienda": "Tipo de Vivienda", "superficie_interior_util_m2":
                 "Superficie Interior Útil (m²)", "porcentaje_ahorro": "Porcentaje de Ahorro (%)", "letra_eficiencia_energetica_dem": "Letra de Eficiencia Energética", "demanda_calefaccion_kwh_m2_ano": "Demanda Calefacción (kWh/m²/año)", "demanda_enfriamiento_kwh_m2_ano": "Demanda Enfriamiento (kWh/m²/año)", "demanda_total_kwh_m2_ano": "Demanda Total (kWh/m²/año)", "emitida_el": "Emitida el"}
RENAME_MAP_P2 = {'region': 'Región', 'comuna': 'Comuna', 'direccion': 'Dirección', 'rol_vivienda': 'Rol Vivienda', 'tipo_vivienda': 'Tipo Vivienda', 'zona_termica': 'Zona Térmica', 'superficie_interior_util_m2': 'Superficie Interior Útil (m²)', 'solicitado_por': 'Solicitado Por', 'evaluado_por': 'Evaluado Por', 'codigo_evaluacion': 'Código Evaluación', 'demanda_calefaccion_kwh_m2_ano': 'Demanda Calefacción Promedio (kWh/m²/año)', 'demanda_enfriamiento_kwh_m2_ano': 'Demanda Enfriamiento Promedio (kWh/m²/año)', 'demanda_total_kwh_m2_ano': 'Demanda Total Promedio (kWh/m²/año)', 'demanda_total_bis_kwh_m2_ano': 'Demanda Total Vivienda Eval. (kWh/m²/año)', 'demanda_total_referencia_kwh_m2_ano': 'Demanda Total Referencia (kWh/m²/año)', 'porcentaje_ahorro': 'Porcentaje Ahorro (%)', 'muro_principal_descripcion': 'Muro Principal: Descripción', 'muro_principal_exigencia_W_m2_K': 'Muro Principal: Exigencia (W/m²K)', 'muro_secundario_descripcion': 'Muro Secundario: Descripción', 'muro_secundario_exigencia_W_m2_K': 'Muro Secundario: Exigencia (W/m²K)', 'piso_principal_descripcion': 'Piso Principal: Descripción',
                 'piso_principal_exigencia_W_m2_K': 'Piso Principal: Exigencia (W/m²K)', 'puerta_principal_descripcion': 'Puerta Principal: Descripción', 'puerta_principal_exigencia': 'Puerta Principal: Exigencia', 'techo_principal_descripcion': 'Techo Principal: Descripción', 'techo_principal_exigencia_W_m2_K': 'Techo Principal: Exigencia (W/m²K)', 'techo_secundario_descripcion': 'Techo Secundario: Descripción', 'techo_secundario_exigencia_W_m2_K': 'Techo Secundario: Exigencia (W/m²K)', 'superficie_vidriada_principal_descripcion': 'Sup. Vidriada Principal: Descripción', 'superficie_vidriada_principal_exigencia': 'Sup. Vidriada Principal: Exigencia', 'superficie_vidriada_secundaria_descripcion': 'Sup. Vidriada Secundaria: Descripción', 'superficie_vidriada_secundaria_exigencia': 'Sup. Vidriada Secundaria: Exigencia', 'ventilacion_rah_descripcion': 'Ventilación (RAH): Descripción', 'ventilacion_rah_exigencia': 'Ventilación (RAH): Exigencia', 'infiltraciones_rah_descripcion': 'Infiltraciones (RAH): Descripción', 'infiltraciones_rah_exigencia': 'Infiltraciones (RAH): Exigencia'}
RENAME_MAP_P3_CONSUMOS = {'codigo_evaluacion': 'Código Evaluación', 'agua_caliente_sanitaria_kwh_m2': 'ACS (kWh/m²)', 'agua_caliente_sanitaria_per': 'ACS (%)', 'iluminacion_kwh_m2': 'Iluminación (kWh/m²)', 'iluminacion_per': 'Iluminación (%)', 'calefaccion_kwh_m2': 'Calefacción (kWh/m²)', 'calefaccion_kwh_per': 'Calefacción (%)', 'energia_renovable_no_convencional_kwh_m2': 'ERNC (kWh/m²)', 'energia_renovable_no_convencional_per': 'ERNC (%)', 'consumo_total_kwh_m2': 'Consumo Total (kWh/m²)', 'emisiones_kgco2_m2_ano': 'Emisiones (kgCO₂e/m²/año)', 'calefaccion_descripcion_proy': 'Calefacción Proy.: Desc.', 'calefaccion_consumo_proy_kwh': 'Calefacción Proy. (kWh)', 'calefaccion_consumo_proy_per': 'Calefacción Proy. (%)', 'iluminacion_descripcion_proy': 'Iluminación Proy.: Desc.', 'iluminacion_consumo_proy_kwh': 'Iluminación Proy. (kWh)', 'iluminacion_consumo_proy_per': 'Iluminación Proy. (%)', 'agua_caliente_sanitaria_descripcion_proy': 'ACS Proy.: Desc.', 'agua_caliente_sanitaria_consumo_proy_kwh': 'ACS Proy. (kWh)', 'agua_caliente_sanitaria_consumo_proy_per': 'ACS Proy. (%)', 'energia_renovable_no_convencional_descripcion_proy': 'ERNC Proy.: Desc.', 'energia_renovable_no_convencional_consumo_proy_kwh': 'ERNC Proy. (kWh)', 'energia_renovable_no_convencional_consumo_proy_per': 'ERNC Proy. (%)', 'consumo_total_requerido_proy_kwh': 'Consumo Total Proy. (kWh)', 'calefaccion_descripcion_ref': 'Calefacción Ref.: Desc.', 'calefaccion_consumo_ref_kwh': 'Calefacción Ref. (kWh)', 'calefaccion_consumo_ref_per': 'Calefacción Ref. (%)', 'iluminacion_descripcion_ref': 'Iluminación Ref.: Desc.',
                          'iluminacion_consumo_ref_kwh': 'Iluminación Ref. (kWh)', 'iluminacion_consumo_ref_per': 'Iluminación Ref. (%)', 'agua_caliente_sanitaria_descripcion_ref': 'ACS Ref.: Desc.', 'agua_caliente_sanitaria_consumo_ref_kwh': 'ACS Ref. (kWh)', 'agua_caliente_sanitaria_consumo_ref_per': 'ACS Ref. (%)', 'energia_renovable_no_convencional_descripcion_ref': 'ERNC Ref.: Desc.', 'energia_renovable_no_convencional_consumo_ref_kwh': 'ERNC Ref. (kWh)', 'energia_renovable_no_convencional_consumo_ref_per': 'ERNC Ref. (%)', 'consumo_total_requerido_ref_kwh': 'Consumo Total Ref. (kWh)', 'consumo_ep_calefaccion_kwh': 'EP Consumo Calef. (kWh)', 'consumo_ep_agua_caliente_sanitaria_kwh': 'EP Consumo ACS (kWh)', 'consumo_ep_iluminacion_kwh': 'EP Consumo Ilum. (kWh)', 'consumo_ep_ventiladores_kwh': 'EP Consumo Vent. (kWh)', 'generacion_ep_fotovoltaicos_kwh': 'EP Gen. FV (kWh)', 'aporte_fotovoltaicos_consumos_basicos_kwh': 'EP Aporte FV (kWh)', 'diferencia_fotovoltaica_para_consumo_kwh': 'EP Dif. FV (kWh)', 'aporte_solar_termica_calefaccion_kwh': 'EP Aporte Solar T. (kWh)', 'aporte_solar_termica_agua_caliente_sanitaria_kwh': 'EP Aporte Solar T. ACS (kWh)', 'total_consumo_ep_antes_fotovoltaica_kwh': 'EP Total Antes FV (kWh)', 'aporte_fotovoltaicos_consumos_basicos_kwh_bis': 'EP Aporte FV Bis (kWh)', 'consumos_basicos_a_suplir_kwh': 'EP Consumos a Suplir (kWh)', 'consumo_total_ep_obj_kwh': 'Consumo Total EP Obj (kWh)', 'consumo_total_ep_ref_kwh': 'Consumo Total EP Ref (kWh)', 'coeficiente_energetico_c': 'Coeficiente Energético (C)'}
RENAME_MAP_P3_ENVOLVENTE = {'codigo_evaluacion': 'Código Evaluación', 'orientacion': 'Orientación', 'elementos_opacos_area_m2': 'Opacos: Área (m²)', 'elementos_opacos_U_W_m2_K': 'Opacos: U (W/m²K)', 'elementos_traslucidos_area_m2': 'Traslúcidos: Área (m²)',
                            'elementos_traslucidos_U_W_m2_K': 'Traslúcidos: U (W/m²K)', 'P01_W_K': 'PT P01 (W/K)', 'P02_W_K': 'PT P02 (W/K)', 'P03_W_K': 'PT P03 (W/K)', 'P04_W_K': 'PT P04 (W/K)', 'P05_W_K': 'PT P05 (W/K)', 'UA_phiL': 'Ht (UA + φL) (W/K)'}
RENAME_MAP_P4 = {'codigo_evaluacion': 'Código Evaluación', 'mes': 'Mes', 'demanda_calef_viv_eval_kwh': 'Dem. Calef. Eval. (kWh)', 'demanda_calef_viv_ref_kwh': 'Dem. Calef. Ref. (kWh)', 'demanda_enfri_viv_eval_kwh': 'Dem. Enfri. Eval. (kWh)', 'demanda_enfri_viv_ref_kwh': 'Dem. Enfri. Ref. (kWh)',
                 'sobrecalentamiento_viv_eval_hr': 'Sobrecalent. Eval. (hr)', 'sobrecalentamiento_viv_ref_hr': 'Sobrecalent. Ref. (hr)', 'sobreenfriamiento_viv_eval_hr': 'Sobreenfri. Eval. (hr)', 'sobreenfriamiento_viv_ref_hr': 'Sobreenfri. Ref. (hr)'}
RENAME_MAP_P5 = {
    'mes': 'Mes',
    'q_recuperado_kwh': 'Calor Recuperado (kWh)',
    'q_puentes_termicos_kwh': 'Puentes Térmicos (kWh)',
    'q_contra_terreno_kwh': 'Contra Terreno (kWh)',
    'q_piso_ventilado_kwh': 'Piso Ventilado (kWh)',
    'q_ventanas_kwh': 'Ventanas (kWh)',
    'q_muros_kwh': 'Muros (kWh)',
    'q_techo_kwh': 'Techo (kWh)',
    'q_infiltraciones_kwh': 'Infiltraciones (kWh)',
    'q_ventilacion_kwh': 'Ventilación (kWh)',
    'q_sol_kwh': 'Ganancia Solar (kWh)',
}
RENAME_MAP_P6 = {'hora': 'Hora'}
for _mes in ['enero', 'abril', 'julio', 'octubre']:
    RENAME_MAP_P6[f't_ext_{_mes}'] = f'T° Exterior - {_mes.capitalize()}'
    RENAME_MAP_P6[f't_int_{_mes}'] = f'T° Interior - {_mes.capitalize()}'
    RENAME_MAP_P6[f't_conf_{_mes}'] = f'T° Confort - {_mes.capitalize()}'
    RENAME_MAP_P6[f'conf_ext_{_mes}'] = f'Conf. Ext. - {_mes.capitalize()}'
    RENAME_MAP_P6[f'conf_int_{_mes}'] = f'Conf. Int. - {_mes.capitalize()}'
    RENAME_MAP_P6[f'conf_conf_{_mes}'] = f'Conf. Conf. - {_mes.capitalize()}'
    RENAME_MAP_P6[f'n_profiles_{_mes}'] = f'Perfiles - {_mes.capitalize()}'
RENAME_MAP_P7 = {'codigo_evaluacion': 'Código Evaluación', 'mandante_nombre': 'Mandante: Nombre', 'mandante_rut': 'Mandante: RUT',
                 'evaluador_nombre': 'Evaluador: Nombre', 'evaluador_rut': 'Evaluador: RUT', 'evaluador_rol_minvu': 'Evaluador: Rol MINVU'}

EFICIENCIA_COLORS = {
    'A+': '#006837', 'A': '#1B7837', 'B': '#5AAE61',
    'C': '#A6D96A', 'D': '#FEE08B', 'E': '#FDAE61',
    'F': '#F46D43', 'G': '#D73027', None: '#CCCCCC'
}

if 'extracted_dfs' not in st.session_state:
    st.session_state.extracted_dfs = None
if 'file_bytes' not in st.session_state:
    st.session_state.file_bytes = None
if 'filename' not in st.session_state:
    st.session_state.filename = None


def is_valid_cev_v2_pdf(pdf_doc: fitz.Document) -> bool:
    if not pdf_doc or len(pdf_doc) != 7:
        return False
    try:
        text_upper = pdf_doc[0].get_text("text").upper()
        return "PRECALIFICACIÓN ENERGÉTICA" in text_upper or "CALIFICACIÓN ENERGÉTICA" in text_upper
    except Exception as e:
        logging.error(f"Validation error: {e}", exc_info=True)
        return False


def display_dataframe_with_title(title: str, data: pd.DataFrame, transpose: bool = False, rename_map: Optional[Dict[str, str]] = None, hide_cols: Optional[List[str]] = None):
    st.header(title)
    if data is None or data.empty:
        st.warning("No data available for this section.")
        return

    data_to_display = data.copy()
    if rename_map:
        cols_to_rename = {k: v for k, v in rename_map.items()
                          if k in data_to_display.columns}
        data_to_display = data_to_display.rename(columns=cols_to_rename)

    if hide_cols:
        data_to_display = data_to_display.drop(
            columns=[c for c in hide_cols if c in data_to_display.columns], errors='ignore')

    is_placeholder = 'content_note' in data.columns
    if is_placeholder and 'content_note' in data_to_display.columns:
        note = data_to_display['content_note'].iloc[0]
        if pd.notna(note):
            st.info(note)
        data_to_display = data_to_display.drop(columns=['content_note'], errors='ignore')
        if transpose and not data_to_display.empty and not all(data_to_display.isna().all()):
            display_data = data_to_display.T
            if len(display_data.columns) == 1:
                display_data.columns = [""]
            st.dataframe(display_data, use_container_width=True)
        return

    if transpose:
        display_data = data_to_display.T
        if len(display_data.columns) == 1:
            display_data.columns = ["Valor"]
        else:
            display_data.columns = [f"Valor {i+1}" for i in range(len(display_data.columns))]
    else:
        display_data = data_to_display

    display_data.index = display_data.index.map(str)
    display_data.columns = display_data.columns.map(str)

    for col in display_data.columns:
        if display_data[col].dtype == 'object':
            display_data[col] = display_data[col].map(lambda x: "" if pd.isna(x) else str(x))

    st.dataframe(display_data, use_container_width=True)


def process_pdf(pdf_document: fitz.Document, filename: str) -> List[pd.DataFrame]:
    extracted_data_frames: List[pd.DataFrame] = []
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
        try:
            df = func(pdf_document)
            extracted_data_frames.append(df)
            logging.info(f"Processed {base_name}")
        except Exception as e:
            logging.error(f"Error processing {base_name}: {e}", exc_info=True)
            progress_container.warning(f"Error extracting '{base_name}'.")
            extracted_data_frames.append(pd.DataFrame())
        progress_bar.progress((i + 1) / total_steps)
    status_text.success("Procesamiento completado.")
    return extracted_data_frames


def local_css(file_name):
    if os.path.exists(file_name):
        with open(file_name) as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)


def display_metric_card(label, value, delta=None):
    st.metric(label=label, value=value, delta=delta)


def eficiencia_badge_html(letra):
    color = EFICIENCIA_COLORS.get(letra, '#CCCCCC')
    return f'<span style="background:{color};color:white;font-weight:bold;padding:2px 10px;border-radius:12px;font-size:0.9em">{letra}</span>'


def plot_energy_flows(df_p5):
    if df_p5.empty or 'mes' not in df_p5.columns:
        st.warning("Datos de flujo no disponibles")
        return
    plot_df = df_p5.set_index('mes')
    flow_cols = [c for c in plot_df.columns if c.startswith('q_') and c.endswith('_kwh')]
    if not flow_cols:
        return
    labels = {
        'q_recuperado_kwh': 'Recuperado', 'q_puentes_termicos_kwh': 'PT',
        'q_contra_terreno_kwh': 'Terreno', 'q_piso_ventilado_kwh': 'Piso Ven.',
        'q_ventanas_kwh': 'Ventanas', 'q_muros_kwh': 'Muros',
        'q_techo_kwh': 'Techo', 'q_infiltraciones_kwh': 'Infilt.',
        'q_ventilacion_kwh': 'Ventil.', 'q_sol_kwh': 'Solar',
    }
    plot_df = plot_df[flow_cols].rename(columns=labels)
    st.bar_chart(plot_df, height=350)


def ocr_confidence_summary(df_p6):
    if df_p6.empty:
        return {}
    result = {}
    for mes in ['enero', 'abril', 'julio', 'octubre']:
        c_ext = df_p6.get(f'conf_ext_{mes}', pd.Series([-1]*24))
        c_int = df_p6.get(f'conf_int_{mes}', pd.Series([-1]*24))
        confs = pd.concat([c_ext, c_int]).dropna()
        confs = confs[confs >= 0]
        avg = confs.mean() if not confs.empty else -1
        n_prof = df_p6[f'n_profiles_{mes}'].iloc[0] if f'n_profiles_{mes}' in df_p6.columns else 2
        if avg >= 0.7:
            status = "✅ Alta"
        elif avg >= 0.45:
            status = "⚠️ Media"
        elif avg >= 0:
            status = "❌ Baja"
        else:
            status = "⬜ Sin datos"
        result[mes] = (avg, status, int(n_prof) if pd.notna(n_prof) else 2)
    return result


def create_single_excel(dfs: List[pd.DataFrame], source_filename: str) -> BytesIO:
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        sheet_info = [
            ("P1_General", 0, RENAME_MAP_P1),
            ("P2_Arquitectura", 1, RENAME_MAP_P2),
            ("P3_Consumos", 2, RENAME_MAP_P3_CONSUMOS),
            ("P3_Envolvente", 3, RENAME_MAP_P3_ENVOLVENTE),
            ("P4_Demanda", 4, RENAME_MAP_P4),
            ("P5_Flujos", 5, RENAME_MAP_P5),
            ("P6_Temperaturas", 6, RENAME_MAP_P6),
            ("P7_Responsables", 7, RENAME_MAP_P7)
        ]
        for sheet_name, idx, ren_map in sheet_info:
            if idx < len(dfs) and not dfs[idx].empty:
                df_temp = dfs[idx].copy()
                if ren_map:
                    cols_to_ren = {k: v for k, v in ren_map.items() if k in df_temp.columns}
                    df_temp = df_temp.rename(columns=cols_to_ren)
                df_temp.to_excel(writer, sheet_name=sheet_name, index=False)
    excel_buffer.seek(0)
    return excel_buffer


def main():
    st.set_page_config(layout="wide", page_title="R2F | CEV Scraper Dashboard")
    local_css("assets/style.css")

    st.title("🏡 Dashboard Informe CEV v2")
    st.markdown("---")

    with st.sidebar:
        st.markdown("<h2 style='color:#1B7837;'>🏡 R2F</h2>", unsafe_allow_html=True)
        st.header("⚙️ Configuración")

        uploaded_file = st.file_uploader(
            "Cargar Informe PDF",
            type="pdf",
            help="Arrastra un archivo PDF de Informe CEV v2."
        )

        if uploaded_file is not None:
            file_bytes = uploaded_file.getvalue()
            name = uploaded_file.name

            if st.session_state.filename != name or st.session_state.extracted_dfs is None:
                st.info("Procesando archivo...")
                try:
                    with fitz.open(stream=file_bytes, filetype="pdf") as pdf_doc:
                        if is_valid_cev_v2_pdf(pdf_doc):
                            dfs = process_pdf(pdf_doc, name)
                            st.session_state.extracted_dfs = dfs
                            st.session_state.file_bytes = file_bytes
                            st.session_state.filename = name
                        else:
                            st.error(f"'{name}' no es un Informe CEV v2 válido.")
                except Exception as e:
                    st.error(f"Error en {name}: {e}")

        if st.session_state.extracted_dfs is not None:
            st.divider()
            st.subheader("📄 Informe Cargado")
            st.caption(st.session_state.filename)

            st.divider()
            excel_data = create_single_excel(
                st.session_state.extracted_dfs,
                st.session_state.filename
            )
            st.download_button(
                "📥 Descargar Excel",
                data=excel_data,
                file_name=f"{st.session_state.filename.replace('.pdf', '')}_CEV.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type="primary"
            )

            if st.button("🗑️ Limpiar y cargar otro"):
                st.session_state.extracted_dfs = None
                st.session_state.file_bytes = None
                st.session_state.filename = None

    if st.session_state.extracted_dfs is None:
        st.info("👋 Bienvenida/o. Carga un archivo PDF en el panel lateral para comenzar.")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### ¿No tienes archivos?")
            st.write("Usa un Informe CEV v2 (PRECALIFICACIÓN ENERGÉTICA o CALIFICACIÓN ENERGÉTICA).")
        return

    dfs = st.session_state.extracted_dfs
    filename = st.session_state.filename

    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 Resumen de Proyecto",
        "⚡ Balance Energético",
        "🏗️ Detalles de Envolvente",
        "🌡️ Análisis Climático"
    ])

    with tab1:
        st.subheader(f"Informe: {filename}")
        p1 = dfs[0]
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            letra_val = p1.get('letra_eficiencia_energetica_dem', ['-'])[0]
            letra_color = EFICIENCIA_COLORS.get(letra_val, '#CCCCCC')
            st.markdown(
                f'<div style="background:{letra_color};color:white;padding:10px;border-radius:10px;text-align:center">'
                f'<div style="font-size:0.8em">Eficiencia</div>'
                f'<div style="font-size:2em;font-weight:bold">{letra_val}</div></div>',
                unsafe_allow_html=True)
        with col_m2:
            display_metric_card("Ahorro Estimado", f"{p1.get('porcentaje_ahorro', [0])[0]}%")
        with col_m3:
            display_metric_card("Superficie Útil", f"{p1.get('superficie_interior_util_m2', [0])[0]} m²")
        with col_m4:
            st.caption("Código Evaluación")
            st.write(f"**{p1.get('codigo_evaluacion', ['N/A'])[0]}**")

        st.divider()
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            display_dataframe_with_title("📄 Antecedentes", dfs[0], transpose=True, rename_map=RENAME_MAP_P1)
        with col_d2:
            display_dataframe_with_title("📍 Ubicación y Proyecto", dfs[1], transpose=True, rename_map=RENAME_MAP_P2)
            st.divider()
            display_dataframe_with_title("👤 Responsables", dfs[7], transpose=True, rename_map=RENAME_MAP_P7)

    with tab2:
        st.subheader("Análisis de Demanda y Consumo")

        st.markdown("##### 🔹 Consumos Requeridos (kWh/m²)")
        consumo_cols = ['agua_caliente_sanitaria_kwh_m2', 'iluminacion_kwh_m2',
                        'calefaccion_kwh_m2', 'energia_renovable_no_convencional_kwh_m2',
                        'consumo_total_kwh_m2', 'emisiones_kgco2_m2_ano']
        if not dfs[2].empty and any(c in dfs[2].columns for c in consumo_cols):
            req_df = dfs[2][[c for c in consumo_cols if c in dfs[2].columns]].T
            req_df.columns = ['Valor']
            req_df.index = [RENAME_MAP_P3_CONSUMOS.get(c, c) for c in req_df.index]
            st.dataframe(req_df, use_container_width=True)

        with st.expander("🔋 Consumos Proyectados y Referencia (kWh)"):
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                st.markdown("**Proyectado**")
                proy_cols = [c for c in dfs[2].columns if c.endswith('_proy_kwh') or c.endswith('_proy_per')]
                if proy_cols:
                    proy_df = dfs[2][proy_cols].T
                    proy_df.columns = ['Valor']
                    proy_df.index = [RENAME_MAP_P3_CONSUMOS.get(c, c) for c in proy_df.index]
                    st.dataframe(proy_df, use_container_width=True)
            with col_p2:
                st.markdown("**Referencia**")
                ref_cols = [c for c in dfs[2].columns if c.endswith('_ref_kwh') or c.endswith('_ref_per')]
                if ref_cols:
                    ref_df = dfs[2][ref_cols].T
                    ref_df.columns = ['Valor']
                    ref_df.index = [RENAME_MAP_P3_CONSUMOS.get(c, c) for c in ref_df.index]
                    st.dataframe(ref_df, use_container_width=True)

        with st.expander("☀️ Generación FV y Solar Térmica"):
            fv_cols = [c for c in dfs[2].columns if 'fotovoltaico' in c or 'solar_termica' in c
                      or 'consumo_ep_' in c or 'coeficiente' in c]
            if fv_cols:
                fv_df = dfs[2][fv_cols].T
                fv_df.columns = ['Valor']
                fv_df.index = [RENAME_MAP_P3_CONSUMOS.get(c, c) for c in fv_df.index]
                st.dataframe(fv_df, use_container_width=True)

        st.divider()
        st.write("#### 📊 Demanda Energética Mensual")
        if not dfs[4].empty:
            plot_df = dfs[4].copy()
            if 'mes_id' in plot_df.columns:
                 plot_df = plot_df.set_index('mes_id')
            cols_to_plot = [
                'demanda_calef_viv_eval_kwh',
                'demanda_calef_viv_ref_kwh',
                'demanda_enfri_viv_eval_kwh',
                'demanda_enfri_viv_ref_kwh'
            ]
            existing_cols = [c for c in cols_to_plot if c in plot_df.columns]
            if existing_cols:
                st.bar_chart(plot_df[existing_cols])

        if not dfs[4].empty:
            plot_df = dfs[4].copy()
            if 'mes_id' in plot_df.columns:
                 plot_df = plot_df.set_index('mes_id')
            oh_cols = [c for c in ['sobrecalentamiento_viv_eval_hr', 'sobrecalentamiento_viv_ref_hr',
                                   'sobreenfriamiento_viv_eval_hr', 'sobreenfriamiento_viv_ref_hr']
                      if c in plot_df.columns]
            if oh_cols:
                st.write("#### 🌡️ Horas de Sobrecalentamiento / Sobreenfriamiento")
                st.bar_chart(plot_df[oh_cols])

        display_dataframe_with_title("📄 Detalle Mensual", dfs[4], transpose=False, rename_map=RENAME_MAP_P4)

    with tab3:
        st.subheader("Especificaciones Técnicas de la Envolvente")

        exigencia_cols = {k: v for k, v in RENAME_MAP_P2.items() if 'exigencia' in k.lower() or 'Exigencia' in v}
        if not dfs[1].empty and exigencia_cols:
            st.markdown("##### ✅ Cumplimiento de Exigencias Térmicas")
            exig_df = dfs[1][[c for c in exigencia_cols.keys() if c in dfs[1].columns]].T
            if not exig_df.empty:
                exig_df.columns = ['Valor (W/m²K)']
                cumplimiento = []
                for val in exig_df['Valor (W/m²K)']:
                    try:
                        fval = float(val)
                        cumplimiento.append("✅ Cumple" if fval <= 3.0 else "⚠️ Revisar")
                    except (ValueError, TypeError):
                        cumplimiento.append("—")
                exig_df['Estado'] = cumplimiento
                idx_names = {k: v for k, v in RENAME_MAP_P2.items()}
                exig_df.index = [idx_names.get(c, c) for c in exig_df.index]
                st.dataframe(exig_df, use_container_width=True)

        display_dataframe_with_title("🧱 Elementos Constructivos por Orientación", dfs[3], transpose=False, rename_map=RENAME_MAP_P3_ENVOLVENTE)

        if not dfs[0].empty and 'letra_eficiencia_energetica_dem' in dfs[0].columns:
            letra = dfs[0]['letra_eficiencia_energetica_dem'].iloc[0]
            st.markdown(f"**Letra de Eficiencia:** {eficiencia_badge_html(letra)}", unsafe_allow_html=True)

    with tab4:
        st.subheader("Variables Ambientales y Flujos")

        st.markdown("##### 🌪️ Flujos Energéticos: Enero vs Julio")
        if not dfs[5].empty:
            col_chart, col_table = st.columns([3, 1])
            with col_chart:
                plot_energy_flows(dfs[5])
            with col_table:
                display_df = dfs[5].copy()
                if RENAME_MAP_P5:
                    cols_to_ren = {k: v for k, v in RENAME_MAP_P5.items() if k in display_df.columns}
                    display_df = display_df.rename(columns=cols_to_ren)
                st.dataframe(display_df, use_container_width=True, hide_index=True)
        st.divider()

        st.markdown("##### 🌡️ Comportamiento Térmico Horario")
        if not dfs[6].empty:
            conf_summary = ocr_confidence_summary(dfs[6])
            badge_cols = st.columns(4)
            for idx, (mes, info) in enumerate(conf_summary.items()):
                with badge_cols[idx]:
                    avg_conf, status, n_prof = info
                    st.metric(
                        label=f"{mes.capitalize()} ({n_prof} perfiles)",
                        value=status,
                        delta=f"conf: {avg_conf:.0%}" if avg_conf >= 0 else None,
                    )

            temp_cols = [c for c in dfs[6].columns if c.startswith('t_') and not c.startswith('t_conf_')]
            if 'hora' in dfs[6].columns and temp_cols:
                chart_data = dfs[6][['hora'] + temp_cols].set_index('hora')
                chart_cols = {}
                for mes in ['enero', 'abril', 'julio', 'octubre']:
                    chart_cols[f't_ext_{mes}'] = f'Ext - {mes.capitalize()}'
                    chart_cols[f't_int_{mes}'] = f'Int - {mes.capitalize()}'
                chart_data = chart_data.rename(columns=chart_cols)
                st.line_chart(chart_data)

            confort_cols = [c for c in dfs[6].columns if c.startswith('t_conf_')]
            if confort_cols and any(dfs[6][c].notna().any() for c in confort_cols):
                st.caption("📌 Líneas punteadas = Temperatura de Confort (disponible en informes con 3 perfiles)")

            st.caption("Temperaturas promedio por hora para meses críticos. Badges indican confianza del OCR.")
        else:
            st.warning("Datos de temperatura no disponibles (error de OCR).")


if __name__ == "__main__":
    main()
