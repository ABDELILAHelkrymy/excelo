# Plan: Generic Excel/CSV Filter GUI App

The current hardcoded `filter.py` will be replaced by a **CustomTkinter desktop app** (`app.py`) that lets the user load any Excel/CSV files, inspect their columns, then apply one of four operations — all without touching code. The original `filter.py` is kept as-is for reference.

---

## Architecture

Single file `app.py` with three logical layers:
- **UI layer** — CustomTkinter widgets, layout, events
- **Data layer** — pandas operations (load, preview, run operation)
- **State** — simple attributes on the App class tracking loaded DataFrames, selected columns, and last result

Dependencies to install: `customtkinter`, `pandas`, `openpyxl`, `CTkTable` (for scrollable previews)

---

## Screen Layout

Three stacked sections inside a scrollable main window:

**Section 1 — File Loader (split left/right)**
- Left (File A — primary): `[Browse]` button → shows filename, populates "Key Column A" dropdown, shows a 5-row preview table
- Right (File B — secondary, optional): same controls — hidden/greyed for extract-only operations

**Section 2 — Operation Panel**
- Segmented button (radio-style): `Minus (A−B)` | `Extend (A∪B)` | `Intersect (A∩B)` | `Extract (Filter A)`
- Sub-panel that swaps based on selection:
  - **Minus / Intersect**: "Column from A" dropdown + "Column from B" dropdown (cross-file key mapping like current script)
  - **Extend**: "Deduplicate" checkbox + optional key column
  - **Extract**: "Column" dropdown + "Condition" dropdown (equals / not equals / contains / starts with / gt / lt / is empty) + value entry box + case-sensitive toggle

**Section 3 — Result & Export**
- Stats bar: `File A: X rows | File B: Y rows | Result: Z rows`
- Scrollable result preview table (first 100 rows, all columns)
- "Columns to keep" multi-select (default: all)
- Output path picker + `[Save as Excel]` / `[Save as CSV]` buttons

**Extras (toolbar / footer)**
- Light/Dark mode toggle (top-right)
- `[Save Profile]` / `[Load Profile]` buttons → persist current file paths + operation settings to a JSON file in the same directory
- `[Clear All]` button

---

## Operations Detail

| Operation | Logic | pandas call |
|---|---|---|
| **Minus** A − B | Rows in A whose key is NOT in B's key | `df_a[~df_a[col_a].isin(set_b)]` |
| **Extend** A ∪ B | Stack A + B vertically, optionally dedup | `pd.concat([a, b]).drop_duplicates(subset=key)` |
| **Intersect** A ∩ B | Rows in A whose key IS in B's key | `df_a[df_a[col_a].isin(set_b)]` |
| **Extract** | Filter A by a value condition on one column | `df_a[df_a[col].str.contains(val)]` etc. |

All key comparisons normalize via `.astype(str).str.strip()` (preserving current script behavior).

---

## Steps

1. Create `app.py` at the workspace root

2. **Imports & config**: `customtkinter`, `pandas`, `tkinter.filedialog`, `json`, `pathlib`; set default appearance mode and color theme

3. **`load_file(path)` helper**: detects `.xlsx`/`.xls` vs `.csv`, returns a DataFrame; shows a `CTkMessagebox` on error

4. **File panel widget** (reusable component class `FilePanel`): encapsulates Browse button, filename label, column dropdown, and preview table; used twice for File A and File B

5. **Operation panel**: `CTkSegmentedButton` drives which sub-frame is `.pack()`ed — each sub-frame is pre-built and swapped in/out on selection change

6. **`run_operation()` method**: reads current UI state → runs the correct pandas logic → stores result DataFrame → updates stats bar and preview table

7. **Preview table**: embed a `ttk.Treeview` (scrollable, works inside CTk windows) with a vertical scrollbar; refresh via `populate_table(df, max_rows=100)`

8. **Column selector**: `CTkScrollableFrame` with `CTkCheckBox` per column — populated after result is ready

9. **Export buttons**: call `df[selected_cols].to_excel(path)` or `.to_csv(path, index=False)` based on chosen format

10. **Profile save/load**: serialize `{file_a, file_b, col_a, col_b, operation, condition, value}` to `filter_profile.json` using `json.dump`; reload and replay the UI state

11. **Dark/Light toggle**: `customtkinter.set_appearance_mode(mode)` bound to a `CTkSwitch`

---

## Proposed Extra Features (nice-to-have, can be added after core)

- **Deduplication tool**: standalone tab to remove duplicates from a single file by chosen column
- **Column rename / mapping table**: when extending two files with different column names, show a mapping UI
- **Batch mode**: apply the same operation to a folder of files (glob pattern)
- **History log**: small log panel at the bottom showing timestamps and row-count outcomes of past runs in the session

---

## Verification

```
pip install customtkinter openpyxl CTkTable
python app.py
```

- Load `data1/liste-global.xlsx` as File A, `data1/liste-payées.xlsx` as File B
- Select **Minus**, map `رقم بطاقة التعريف الوطنية` → `N_Identifiant`
- Run → result row count should match current `liste-non-payées.xlsx`
- Save output and confirm it matches the existing file

---

## Decisions

- Chose **CustomTkinter** over Tkinter for modern look; `ttk.Treeview` embedded for the table widget since CTkTable has no built-in scroll for large datasets
- Single `app.py` file (no sub-packages) — appropriate for this script's scale
- Profile JSON saved to the workspace root alongside `app.py`
