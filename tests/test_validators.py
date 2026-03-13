import pytest
from openpyxl import Workbook
from src.validators import ExcelValidators


@pytest.fixture
def mock_workbook():
  """Create temporary Excel workbook for tests."""
  wb = Workbook()
  ws = wb.active
  return wb, ws


def test_check_columns_all_present(mock_workbook):
  """Test: all mandatory columns are present."""
  wb, ws = mock_workbook
  ws.append(["Column 1", "Column 3"])

  required = ["Column 1", "Column 2", "Column 3", "Column 4"]
  errors = ExcelValidators.check_columns(ws, required)

  assert any("MISSING" in err for err in errors)
  assert "Column 2" in str(errors)


def test_check_columns_extra(mock_workbook):
  """Test: detect any extra columns."""
  wb, ws = mock_workbook
  ws.append(["Column 1", "Column 2", "Extra Column"])

  required = ["Column 1", "Column 2"]
  errors = ExcelValidators.check_columns(ws, required)

  assert any("EXTRA" in err for err in errors)
  assert "Extra Column" in str(errors)
  
