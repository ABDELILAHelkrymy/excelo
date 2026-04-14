import pytest
import polars as pl
from operations import (
    op_minus_or_intersect,
    op_extend,
    op_enrich,
    op_extract,
    op_combine
)

@pytest.fixture
def df_a():
    return pl.DataFrame({
        "ID": ["1", "2", "3", "4"],
        "Name": ["Alice", "Bob", "Charlie", "David"]
    })

@pytest.fixture
def df_b():
    return pl.DataFrame({
        "Key": ["3", "4", "5", "6"],
        "Age": ["30", "40", "50", "60"]
    })

def test_op_minus(df_a, df_b):
    # Subtract B from A using ID/Key
    res, nf = op_minus_or_intersect(
        minus=True, df_a=df_a, df_b=df_b,
        mi_mode="Standard (1 A × 1 B)",
        mi_a_col="ID", mi_b_col="Key",
        mi_multi_files={}, mi_multi_key="", mi_nf=False,
        mi_a_col_vars={}, mi_b_col_vars={},
        mi_suffix="", mi_principal="A"
    )
    # Expected: IDs 1 and 2 are left
    assert len(res) == 2
    assert "1" in res["ID"].to_list()
    assert "2" in res["ID"].to_list()
    assert nf is None

def test_op_intersect(df_a, df_b):
    # Intersect A and B
    res, nf = op_minus_or_intersect(
        minus=False, df_a=df_a, df_b=df_b,
        mi_mode="Standard (1 A × 1 B)",
        mi_a_col="ID", mi_b_col="Key",
        mi_multi_files={}, mi_multi_key="", mi_nf=True,
        mi_a_col_vars={"ID": True, "Name": True},
        mi_b_col_vars={"Age": True},
        mi_suffix=" (B)", mi_principal="A"
    )
    # Expected: IDs 3 and 4
    assert len(res) == 2
    assert "3" in res["ID"].to_list()
    assert "Age" in res.columns
    # NF should contain IDs 5 and 6
    assert nf is not None
    assert len(nf) == 2
    assert "5" in nf["Key"].to_list()

def test_op_extend(df_a, df_b):
    res = op_extend(df_a, df_b, ext_dedup=False, ext_key="")
    assert len(res) == 8
    assert "Name" in res.columns
    assert "Age" in res.columns

def test_op_enrich(df_a, df_b):
    res, en_found = op_enrich(
        df_a=df_a, df_b=df_b,
        en_a_col="ID", en_b_col="Key",
        en_cols={"Age": True}, en_mode="LEFT (Garder tout A)",
        en_suffix=" (B)", en_skip_empty=False,
        en_fill="Chaîne vide", en_fill_val=""
    )
    assert len(res) == 4
    assert "Age" in res.columns
    assert list(res.filter(pl.col("ID") == "1")["Age"])[0] == ""
    assert list(res.filter(pl.col("ID") == "3")["Age"])[0] == "30"

def test_op_extract(df_a):
    res = op_extract(df_a, [
        {"col": "Name", "cond": "contient", "val": "li", "case": False, "logic": ""}
    ])
    assert len(res) == 2
    names = res["Name"].to_list()
    assert "Alice" in names
    assert "Charlie" in names

def test_op_combine(df_a):
    cb_files = {
        "file1": {"Sheet1": df_a},
        "file2": {"Sheet1": df_a}
    }
    res = op_combine(cb_files, cb_dedup=True, cb_dedup_key="ID", cb_cols={"ID": True, "Name": True})
    assert len(res) == 4  # Deduped to original 4 rows
