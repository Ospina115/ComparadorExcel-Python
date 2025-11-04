#!/usr/bin/env python3
"""
Script para comparar archivos Excel entre dos carpetas.

Funcionamiento básico:
- Lee las rutas desde un archivo .env (FOLDER_A, FOLDER_B, OUTPUT_DIR)
- Empareja archivos entre las dos carpetas por nombre (matching exacto o fuzzy)
- Para cada par de archivos, lee la primera hoja y busca filas nuevas, eliminadas y modificadas
- Guarda un archivo Excel con hojas: added, removed, modified en OUTPUT_DIR

Notas/Asunciones:
- Si existe una columna de clave (id, ID, Codigo, etc.) se usa como llave para detectar modificaciones.
- Si no hay clave, se compara por fila completa (hash de la fila) para detectar añadidos/eliminados; las modificaciones no se pueden detectar sin llave.

Requisitos: pandas, python-dotenv, openpyxl
"""
from pathlib import Path
import os
import difflib
import hashlib
import sys
from dotenv import load_dotenv
import pandas as pd


def row_hash(series: pd.Series) -> str:
    """Genera un hash para una fila (serie) para comparar igualdad de contenido."""
    # Convertir valores en string y normalizar None/NaN
    vals = ['' if pd.isna(v) else str(v).strip() for v in series.values]
    s = '|'.join(vals)
    return hashlib.md5(s.encode('utf-8')).hexdigest()


def find_key_column(df: pd.DataFrame):
    """Intentar inferir una columna que actúe como clave primaria."""
    candidates = ['id', 'ID', 'Id', 'codigo', 'Código', 'codigo_municipio', 'codigo_munic', 'codigoMunicipio', 'cod']
    for c in df.columns:
        if c in candidates:
            return c
    # también preferir columnas únicas
    for c in df.columns:
        if df[c].is_unique:
            return c
    return None


def pair_files(files_a, files_b, threshold=0.9):
    """Empareja archivos de A con B usando nombre base y fuzzy matching.
    Retorna lista de tuplas (path_a, path_b, score)
    """
    pairs = []
    stems_b = {f.stem: f for f in files_b}
    for fa in files_a:
        a_stem = fa.stem
        # match exacto (case-insensitive)
        for bstem, fb in stems_b.items():
            if a_stem.lower() == bstem.lower():
                pairs.append((fa, fb, 1.0))
                break
        else:
            # fuzzy
            candidates = list(stems_b.keys())
            match = difflib.get_close_matches(a_stem, candidates, n=1, cutoff=threshold)
            if match:
                pairs.append((fa, stems_b[match[0]], difflib.SequenceMatcher(None, a_stem, match[0]).ratio()))
    return pairs


def compare_dataframes(df_a: pd.DataFrame, df_b: pd.DataFrame):
    """Compara dos DataFrames y retorna added, modified (DataFrames).

    - Si hay columna clave: detecta added (en B no en A) y modified (mismo key, valores distintos),
      devolviendo para modified las filas desde df_b (datos más recientes).
    - Si no hay clave: detecta added por hash de fila; modified queda vacío (no se puede inferir).
    """
    key = find_key_column(df_a)
    if key and key in df_b.columns:
        # usar key
        a_idx = df_a.set_index(key)
        b_idx = df_b.set_index(key)
        added_keys = b_idx.index.difference(a_idx.index)
        common = a_idx.index.intersection(b_idx.index)

        added = b_idx.loc[added_keys].reset_index() if len(added_keys) > 0 else pd.DataFrame()

        modified_keys = []
        for k in common:
            row_a = a_idx.loc[k]
            row_b = b_idx.loc[k]
            if not row_a.fillna('').equals(row_b.fillna('')):
                modified_keys.append(k)

        modified = b_idx.loc[modified_keys].reset_index() if modified_keys else pd.DataFrame()
        return added, modified
    else:
        # comparar por hash de la fila (cuando no hay clave)
        ha = df_a.apply(row_hash, axis=1)
        hb = df_b.apply(row_hash, axis=1)
        set_a = set(ha)
        set_b = set(hb)
        added_hashes = set_b - set_a
        added = df_b[[h in added_hashes for h in hb.values]]
        modified = pd.DataFrame()  # no es posible detectar modificaciones sin clave
        return added.reset_index(drop=True), modified


def main():
    load_dotenv()
    FOLDER_A = os.getenv('FOLDER_A')
    FOLDER_B = os.getenv('FOLDER_B')
    OUTPUT_DIR = os.getenv('OUTPUT_DIR', 'comparisons_output')
    THRESHOLD = float(os.getenv('FUZZY_THRESHOLD', '0.9'))

    if not FOLDER_A or not FOLDER_B:
        print('Por favor configure FOLDER_A y FOLDER_B en el archivo .env')
        sys.exit(1)

    pA = Path(FOLDER_A)
    pB = Path(FOLDER_B)
    out = Path(OUTPUT_DIR)
    out.mkdir(parents=True, exist_ok=True)

    exts = {'.xlsx', '.xls', '.xlsm'}
    files_a = [f for f in pA.iterdir() if f.suffix in exts and f.is_file()]
    files_b = [f for f in pB.iterdir() if f.suffix in exts and f.is_file()]

    print(f'Archivos en A: {len(files_a)}, en B: {len(files_b)}')

    pairs = pair_files(files_a, files_b, threshold=THRESHOLD)
    print(f'Encontrados {len(pairs)} pares potenciales (threshold={THRESHOLD})')

    for fa, fb, score in pairs:
        print(f'Comparando: {fa.name}  <->  {fb.name}  (score={score:.2f})')
        try:
            # leer primera hoja de cada archivo
            dfa = pd.read_excel(fa, sheet_name=0)
            dfb = pd.read_excel(fb, sheet_name=0)
        except Exception as e:
            print(f'Error leyendo uno de los archivos: {e}')
            continue

        added, modified = compare_dataframes(dfa, dfb)

        # Si no hay rows añadidas ni modificadas, no generar archivo
        if (added.empty if isinstance(added, pd.DataFrame) else True) and (modified.empty if isinstance(modified, pd.DataFrame) else True):
            print('No se encontraron cambios (ni añadidos ni modificados). No se genera archivo.')
            continue

        # guardar resultado en un excel con las hojas necesarias (solo added y/o modified)
        out_file = out / f"{fa.stem}_vs_{fb.stem}_comparacion.xlsx"
        with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
            if not added.empty:
                added.to_excel(writer, sheet_name='added', index=False)
            if not modified.empty:
                modified.to_excel(writer, sheet_name='modified', index=False)

        print(f'Resultado guardado en: {out_file}')


if __name__ == '__main__':
    main()
