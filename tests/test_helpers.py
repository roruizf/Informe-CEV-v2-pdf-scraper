import pytest
import pandas as pd
from scraping_functions import (
    safe_float_convert,
    extraer_numeros_de_texto,
    extraer_numeros_de_elemento,
    _from_procentaje_ahorro_to_letra,
    normalize_coordinates
)

def test_safe_float_convert():
    # Test standard cases
    assert safe_float_convert("1.234,56") == 1234.56
    assert safe_float_convert("100,0") == 100.0
    assert safe_float_convert("75.5") == 75.5
    
    # Test multi-line cases (new functionality)
    multi_line = "ignore this\n123,45\nand this"
    assert safe_float_convert(multi_line) == 123.45
    
    # Test empty/None cases
    assert safe_float_convert("") is None
    assert safe_float_convert(None) is None
    assert safe_float_convert("", default=0.0) == 0.0
    
    # Test invalid cases
    assert safe_float_convert("not a number") is None
    assert safe_float_convert("12.34.56") is None # Multiple dots without thousand separator logic
    assert safe_float_convert("   123,45   ") == 123.45 # Extra whitespace
    assert safe_float_convert("0,0") == 0.0
    assert safe_float_convert("-5,5") == -5.5
    assert safe_float_convert("1.000,50") == 1000.50 # Standard Chilean format
    assert safe_float_convert("1,000.50") == 1000.50 # US format

def test_extraer_numeros_de_texto():
    text = "The values are 123.45 and 67,89 and 1.000 and 100."
    # 123.45 -> decimal point (2 digits after)
    # 67,89 -> decimal comma (normalized to point)
    # 1.000 -> thousand separator (3 digits after)
    # 100 -> integer
    assert extraer_numeros_de_texto(text) == ["123.45", "67.89", "1000", "100"]

def test_extraer_numeros_de_elemento():
    # 12.34 -> decimal point (not 3 digits after)
    assert extraer_numeros_de_elemento("12.34") == ["12.34"]
    # 1.234 -> thousand separator (3 digits after)
    assert extraer_numeros_de_elemento("1.234") == ["1234"]
    assert extraer_numeros_de_elemento("ABC 123 DEF") == ["123"]

def test_from_procentaje_ahorro_to_letra():
    assert _from_procentaje_ahorro_to_letra(0.86) == "A+"
    assert _from_procentaje_ahorro_to_letra(0.75) == "A"
    assert _from_procentaje_ahorro_to_letra(0.60) == "B"
    assert _from_procentaje_ahorro_to_letra(0.45) == "C"
    assert _from_procentaje_ahorro_to_letra(0.30) == "D"
    assert _from_procentaje_ahorro_to_letra(0.0) == "E"
    assert _from_procentaje_ahorro_to_letra(-0.2) == "F"
    assert _from_procentaje_ahorro_to_letra(-0.5) == "G"
    assert _from_procentaje_ahorro_to_letra(None) is None

def test_normalize_coordinates():
    # Test typical normalization (scaling from 215.9x279.4 to 612x792)
    x, y = normalize_coordinates(107.95, 139.7, 215.9, 279.4, 612, 792)
    assert x == pytest.approx(306.0)
    assert y == pytest.approx(396.0)
    
    # Test zero division handling
    x, y = normalize_coordinates(10, 10, 0, 0, 100, 100)
    assert x == 0.0
    assert y == 0.0
