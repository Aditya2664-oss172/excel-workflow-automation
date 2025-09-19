import pandas as pd
import pytest
from excel_automation import clean_data, validate_with_sql, load_data

# Sample test DataFrame
@pytest.fixture
def sample_df():
    return pd.DataFrame({
        "Name": ["Alice", "Bob", "Alice", None],
        "Amount": [100, -50, None, 200],
        "Category": ["A", None, "B", "C"]
    })

def test_load_data_excel(tmp_path):
    # Create a temporary Excel file
    df = pd.DataFrame({"Col1": [1, 2], "Col2": ["A", "B"]})
    temp_file = tmp_path / "test.xlsx"
    df.to_excel(temp_file, index=False)

    loaded_df = load_data(str(temp_file))
    assert not loaded_df.empty
    assert list(loaded_df.columns) == ["Col1", "Col2"]

def test_clean_data(sample_df):
    cleaned_df = clean_data(sample_df)

    # Check duplicates removed
    assert len(cleaned_df) < len(sample_df)

    # Check missing numeric values filled
    assert cleaned_df["Amount"].isnull().sum() == 0

    # Check missing text values filled
    assert "Unknown" in cleaned_df["Category"].values or \
           "Unknown" in cleaned_df["Name"].values

def test_validate_with_sql(sample_df):
    cleaned_df = clean_data(sample_df)
    validated_df = validate_with_sql(cleaned_df)

    # Ensure no negative amounts
    assert (validated_df["Amount"] >= 0).all()

    # Ensure columns remain the same
    assert list(validated_df.columns) == list(cleaned_df.columns)
