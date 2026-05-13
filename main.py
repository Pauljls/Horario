import json
from pathlib import Path
from typing import Any

import pandas as pd

# Fila (base 0) donde están los encabezados reales en el Excel
_HEADER_ROW = 2
# Columna con basura a la izquierda que se elimina
_DROP_COL = 0

# Palabras clave para filtrar filas de recreo/almuerzo
_SKIP_PATTERNS = r"RECREO|ALMUERZO|R\s+E\s+C|A\s+L\s+M"


def _clean_sheet(df: pd.DataFrame) -> list[dict[str, Any]]:
    # Usar la fila _HEADER_ROW como encabezados
    headers = df.iloc[_HEADER_ROW].tolist()
    df = df.iloc[_HEADER_ROW + 1 :].copy()
    df.columns = headers

    # Eliminar primera columna (siempre vacía en este formato)
    first_col = df.columns[_DROP_COL]
    df = df.drop(columns=[first_col])

    # Normalizar nombres de columnas: minúsculas, sin espacios extra
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Filtrar filas de RECREO, ALMUERZO y filas completamente vacías
    hora_col = df.columns[0]
    mask_skip = df[hora_col].astype(str).str.contains(
        _SKIP_PATTERNS, case=False, na=False, regex=True
    )
    df = df[~mask_skip].dropna(how="all").reset_index(drop=True)

    # Convertir a object para que where() pueda insertar None
    df = df.astype(object).where(pd.notna(df), None)

    return df.to_dict(orient="records")


def read_schedule(
    file_path: str, sheet_name: int | str | None = None
) -> dict[str, list[dict[str, Any]]] | list[dict[str, Any]]:
    """Lee un archivo Excel y devuelve los horarios limpios.

    Si `sheet_name` es None, lee todas las hojas y devuelve un diccionario con
    el nombre de cada hoja como clave y su arreglo de filas como valor.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {file_path}")

    if sheet_name is None:
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=str)
        return {name: _clean_sheet(df) for name, df in sheets.items()}

    df = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=str)
    return _clean_sheet(df)


def parse_sheet_name(value: str | None) -> int | str | None:
    if value is None:
        return None
    if value.isdigit():
        return int(value)
    return value


if __name__ == "__main__":
    archivo = Path(__file__).parent / "horarios.xlsx"
    resultado = read_schedule(str(archivo))
    print(json.dumps(resultado, indent=2, ensure_ascii=False))
