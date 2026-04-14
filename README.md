# Excel / CSV Filter Tool

A generic, interactive desktop application built with Python and `customtkinter` to process, filter, and modify Excel and CSV files.

## Features

The tool allows you to perform several operations on tabular data efficiently:
- **Soustraire (A − B)**: Return rows of A whose key is NOT present in B.
- **Fusionner (A ∪ B)**: Stack A and B vertically. Option to remove duplicates.
- **Intersecter (A ∩ B)**: Return rows of A whose key IS present in B.
- **Enrichir (A ← B)**: Compare A and B on a key. Keep all rows of A and add chosen columns from B (LEFT JOIN).
- **Combiner (Multi-fichiers)**: Stack several Excel/CSV files, choose columns to keep, and remove duplicates.
- **Extraire (Filtrer A)**: Filter rows of A based on a condition on a chosen column.
- **Scinder (Onglets)**: Split each sheet of a multi-sheet Excel file into separate files.
- **Scinder (Colonnes)**: Load a single file, choose columns to keep, and select one or more columns to split by. This generates one file per unique value combination.
- **Scinder (Lignes)**: Load a single file and break it down into smaller files containing the specified number of rows.

## Installation

1. Make sure you have [Python 3.8+](https://www.python.org/) installed.
2. Clone this repository or download the source code.
3. (Optional but recommended) Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   
   # On Windows:
   .venv\Scripts\activate
   
   # On macOS/Linux:
   source .venv/bin/activate
   ```
4. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

You can launch the tool by double-clicking the `run.bat` (or `run.vbs` for a silent start) file on Windows, or by running the script from your terminal:

```bash
python app.py
```

Select your operation, load your files (`A` and/or `B`), choose your configuration settings on the user-friendly interface, and launch the operation.
