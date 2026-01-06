import pytest
from unittest.mock import MagicMock
import fitz

@pytest.fixture
def mock_pdf_document():
    """Creates a mock fitz.Document with mocked pages."""
    mock_doc = MagicMock(spec=fitz.Document)
    mock_doc.__len__.return_value = 7
    
    # Mock pages
    pages = []
    for i in range(7):
        mock_page = MagicMock(spec=fitz.Page)
        mock_page.number = i
        mock_page.rect = MagicMock()
        mock_page.rect.width = 612
        mock_page.rect.height = 792
        pages.append(mock_page)
    
    mock_doc.__getitem__.side_effect = lambda i: pages[i]
    return mock_doc
