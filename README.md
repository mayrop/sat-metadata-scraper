# SAT CFDI Catálogos

Scripts para descargar, procesar y publicar los catálogos oficiales del SAT en HuggingFace.

## Estructura

```
scripts/
  catalogos/
    scrape.py                   # Descarga los XLS del SAT
    extract.py                  # Extrae los catálogos a CSV
    generate_hf.py              # Genera el folder para HuggingFace
    generate_uso_regimen_table.py
  listado_69b/
    scrape.py                   # Descarga los CSVs del SAT (69-B y 69-B Bis)
    merge.py                    # Limpia, transforma y combina en un solo CSV
    generate_hf.py              # Genera el folder para HuggingFace
hf/
  csv/                          # CSVs intermedios
  xls/                          # XLS descargados del SAT
  dataset/
    catalogos/                  # → mayrop/sat-catalogos
    listado-69b/                # → mayrop/sat-listado-69b
output/                         # Manifests y archivos crudos descargados
```

---

## Dataset: sat-catalogos

Catálogos del Anexo 20 (factura electrónica, retenciones) y complementos.

HuggingFace: `mayrop/sat-catalogos`

```bash
# 1. Descargar los XLS del SAT
uv run scripts/catalogos/scrape.py

# 2. Extraer los catálogos a CSV
uv run scripts/catalogos/extract.py

# 3. Generar el folder del dataset para HuggingFace
uv run scripts/catalogos/generate_hf.py

# 4. Subir a HuggingFace
hf upload-large-folder mayrop/sat-catalogos hf/dataset/catalogos/ --repo-type=dataset
```

---

## Dataset: sat-listado-69b

Listados de contribuyentes publicados bajo los Artículos 69-B y 69-B Bis del CFF.

HuggingFace: `mayrop/sat-listado-69b`

```bash
# 1. Descargar los archivos crudos del SAT
uv run scripts/listado_69b/scrape.py

# 2. Limpiar, transformar y combinar en un solo CSV
uv run scripts/listado_69b/merge.py

# 3. Generar el folder del dataset para HuggingFace
uv run scripts/listado_69b/generate_hf.py

# 4. Subir a HuggingFace
hf upload-large-folder mayrop/sat-listado-69b hf/dataset/listado-69b/ --repo-type=dataset
```
