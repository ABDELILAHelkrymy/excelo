import pathlib


def ensure_excel_output_path(path: str) -> str:
    """Normalize export paths so every save targets an .xlsx file."""
    normalized = path.strip()
    if not normalized:
        return ""

    output_path = pathlib.Path(normalized)
    if output_path.suffix.lower() == ".xlsx":
        return str(output_path)
    if output_path.suffix:
        output_path = output_path.with_suffix(".xlsx")
    else:
        output_path = output_path.with_name(f"{output_path.name}.xlsx")
    return str(output_path)
