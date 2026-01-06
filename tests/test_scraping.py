import pytest
from unittest.mock import MagicMock, patch
import pandas as pd
import fitz
from scraping_functions import (
    get_informe_cev_v2_pagina1_as_dataframe,
    get_informe_cev_v2_pagina1_as_dict,
    get_informe_cev_v2_pagina2_as_dict,
    get_informe_cev_v2_pagina2_as_dataframe,
    get_informe_cev_v2_pagina3_consumos_as_dict,
    get_informe_cev_v2_pagina3_consumos_as_dataframe,
    get_informe_cev_v2_pagina3_envolvente_as_dict,
    get_informe_cev_v2_pagina3_envolvente_as_dataframe,
    get_informe_cev_v2_pagina4_as_dict,
    get_informe_cev_v2_pagina4_as_dataframe,
    get_informe_cev_v2_pagina5_as_dict,
    get_informe_cev_v2_pagina5_as_dataframe,
    get_informe_cev_v2_pagina6_as_dict,
    get_informe_cev_v2_pagina6_as_dataframe,
    get_informe_cev_v2_pagina7_as_dict,
    get_informe_cev_v2_pagina7_as_dataframe
)

@patch('scraping_functions.extract_text_from_area')
def test_get_informe_cev_v2_pagina1_as_dict(mock_extract, mock_pdf_document):
    # Setup mock returns for various fields in Page 1
    # This is a simplified mock to verify the glue logic
    def side_effect(page, rect):
        # Return different strings based on the field name (extracted from the mock call if possible)
        # or just generic test data
        return "Test Data"

    mock_extract.side_effect = side_effect
    
    result = get_informe_cev_v2_pagina1_as_dict(mock_pdf_document)
    
    assert isinstance(result, dict)
    assert result['region'] == "Test Data"
    assert result['comuna'] == "Test Data"

@patch('scraping_functions.get_informe_cev_v2_pagina1_as_dict')
def test_get_informe_cev_v2_pagina1_as_dataframe(mock_get_dict, mock_pdf_document):
    # Mock the dict return to test DF conversion
    mock_get_dict.return_value = {
        'codigo_evaluacion': '12345',
        'region': 'Metropolitana',
        'comuna': 'Santiago'
    }
    
    df = get_informe_cev_v2_pagina1_as_dataframe(mock_pdf_document)
    
    assert isinstance(df, pd.DataFrame)
    assert df.iloc[0]['region'] == 'Metropolitana'
    assert 'codigo_evaluacion' not in df.columns # Should be dropped/processed

@patch('scraping_functions.extract_text_from_area')
def test_get_informe_cev_v2_pagina2_as_dict(mock_extract, mock_pdf_document):
    mock_extract.return_value = "Page 2 Data"
    result = get_informe_cev_v2_pagina2_as_dict(mock_pdf_document)
    assert isinstance(result, dict)
    assert 'region' in result

@patch('scraping_functions.get_informe_cev_v2_pagina3_consumos_as_dict')
def test_get_informe_cev_v2_pagina3_consumos_as_dataframe(mock_get_dict, mock_pdf_document):
    mock_get_dict.return_value = {'agua_caliente_sanitaria_kwh_m2': 100.0, 'codigo_evaluacion': '123'}
    df = get_informe_cev_v2_pagina3_consumos_as_dataframe(mock_pdf_document)
    assert isinstance(df, pd.DataFrame)
    assert 'codigo_evaluacion' not in df.columns

@patch('scraping_functions.get_informe_cev_v2_pagina3_envolvente_as_dict')
def test_get_informe_cev_v2_pagina3_envolvente_as_dataframe(mock_get_dict, mock_pdf_document):
    mock_get_dict.return_value = {'orientacion': ['N', 'S'], 'codigo_evaluacion': ['123', '123']}
    df = get_informe_cev_v2_pagina3_envolvente_as_dataframe(mock_pdf_document)
    assert isinstance(df, pd.DataFrame)
    assert len(df) == 2
    assert 'codigo_evaluacion' not in df.columns

@patch('scraping_functions.get_informe_cev_v2_pagina4_as_dict')
def test_get_informe_cev_v2_pagina4_as_dataframe(mock_get_dict, mock_pdf_document):
    mock_get_dict.return_value = {'mes_id': [1], 'demanda': [50.0], 'codigo_evaluacion': ['123']}
    df = get_informe_cev_v2_pagina4_as_dataframe(mock_pdf_document)
    assert isinstance(df, pd.DataFrame)
    assert 'mes' in df.columns # Should be mapped
    assert 'codigo_evaluacion' not in df.columns

@patch('scraping_functions.get_informe_cev_v2_pagina5_as_dict')
def test_get_informe_cev_v2_pagina5_as_dataframe(mock_get_dict, mock_pdf_document):
    mock_get_dict.return_value = {'mes': ['Enero'], 'q_sol_kwh': [10.0], 'codigo_evaluacion': ['123']}
    df = get_informe_cev_v2_pagina5_as_dataframe(mock_pdf_document)
    assert isinstance(df, pd.DataFrame)
    assert 'codigo_evaluacion' not in df.columns

@patch('scraping_functions.get_informe_cev_v2_pagina6_as_dict')
def test_get_informe_cev_v2_pagina6_as_dataframe(mock_get_dict, mock_pdf_document):
    mock_get_dict.return_value = {'hora': [1], 't_ext_enero': [20.0], 'codigo_evaluacion': ['123']}
    df = get_informe_cev_v2_pagina6_as_dataframe(mock_pdf_document)
    assert isinstance(df, pd.DataFrame)
    assert 'codigo_evaluacion' not in df.columns

@patch('scraping_functions.get_informe_cev_v2_pagina7_as_dict')
def test_get_informe_cev_v2_pagina7_as_dataframe(mock_get_dict, mock_pdf_document):
    mock_get_dict.return_value = {'evaluador': 'Juan', 'codigo_evaluacion': '123'}
    df = get_informe_cev_v2_pagina7_as_dataframe(mock_pdf_document)
    assert isinstance(df, pd.DataFrame)
    assert 'codigo_evaluacion' not in df.columns
