import polars as pl
import pytest

from export_utils import ensure_excel_output_path
from operations import op_enrich


def test_op_enrich_blank_fill_casts_joined_numeric_values_to_text():
    df_a = pl.DataFrame({"ID": [1, 2]})
    df_b = pl.DataFrame({"Key": [2], "Age": [30]})

    result, _ = op_enrich(
        df_a=df_a,
        df_b=df_b,
        en_a_col="ID",
        en_b_col="Key",
        en_cols={"Age": True},
        en_mode="Toutes les lignes de A (LEFT JOIN)",
        en_suffix=" (B)",
        en_skip_empty=False,
        en_fill="Chaîne vide",
        en_fill_val="",
    )

    assert result["Age"].to_list() == ["", "30"]


def test_op_enrich_custom_fill_preserves_numeric_dtype_when_possible():
    df_a = pl.DataFrame({"ID": [1, 2]})
    df_b = pl.DataFrame({"Key": [2], "Age": [30]})

    result, _ = op_enrich(
        df_a=df_a,
        df_b=df_b,
        en_a_col="ID",
        en_b_col="Key",
        en_cols={"Age": True},
        en_mode="Toutes les lignes de A (LEFT JOIN)",
        en_suffix=" (B)",
        en_skip_empty=False,
        en_fill="Valeur personnalisée",
        en_fill_val="0",
    )

    assert result.schema["Age"] == pl.Int64
    assert result["Age"].to_list() == [0, 30]


def test_op_enrich_custom_fill_falls_back_to_text_for_non_numeric_value():
    df_a = pl.DataFrame({"ID": [1, 2]})
    df_b = pl.DataFrame({"Key": [2], "Age": [30]})

    result, _ = op_enrich(
        df_a=df_a,
        df_b=df_b,
        en_a_col="ID",
        en_b_col="Key",
        en_cols={"Age": True},
        en_mode="Toutes les lignes de A (LEFT JOIN)",
        en_suffix=" (B)",
        en_skip_empty=False,
        en_fill="Valeur personnalisée",
        en_fill_val="N/A",
    )

    assert result["Age"].to_list() == ["N/A", "30"]


@pytest.mark.parametrize(
    ("raw_path", "expected"),
    [
        ("report", "report.xlsx"),
        ("report.tmp", "report.xlsx"),
        ("report.xlsx", "report.xlsx"),
        (" folder\\summary.tmp ", "folder\\summary.xlsx"),
    ],
)
def test_ensure_excel_output_path(raw_path, expected):
    assert ensure_excel_output_path(raw_path) == expected
