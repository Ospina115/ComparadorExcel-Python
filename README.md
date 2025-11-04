# Comparador de archivos Excel

Este proyecto contiene un script en Python para comparar archivos Excel entre dos carpetas.

Pasos rápidos:

1. Instalar dependencias (recomendado crear un entorno virtual):

```powershell
python -m venv .venv; .\.venv\Scripts\Activate.ps1; pip install -r requirements.txt
```

2. Editar el archivo `.env` y poner rutas absolutas para `FOLDER_A`, `FOLDER_B` y `OUTPUT_DIR`.

3. Ejecutar el script:

```powershell
python compare_excels.py
```

Qué hace el script:

- Empareja archivos en ambas carpetas por nombre (coincidencia exacta o fuzzy).
- Para cada par, compara la primera hoja.
- Si encuentra una columna clave (por ejemplo `id`, `ID`, `codigo`) la usa para detectar filas modificadas.
- Genera un archivo Excel en `OUTPUT_DIR` con hojas `added`, `removed`, `modified`.

Notas y recomendaciones:

- Si los nombres de archivo no coinciden exactamente, ajustar `FUZZY_THRESHOLD` en `.env`.
- Si tus archivos contienen varias hojas y quieres comparar otra hoja, modifica el script para seleccionar la hoja deseada.
