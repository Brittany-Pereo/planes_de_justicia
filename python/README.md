# Planes de Justicia — Streamlit (Python)

Port a Python/Streamlit de la app Shiny (`R/`) de este mismo repositorio.
Consulta la productividad por Plan de Justicia (CLUES) del IMSS Bienestar,
muestra el avance contra la meta anual y permite descargar los datos en
Excel o un informe en PowerPoint.

## Correr localmente

```bash
cd python
pip install -r requirements.txt
streamlit run app.py
```

## Estructura

- `app.py` — app de Streamlit (UI, gráficas, descargas).
- `data_io.py` — carga y consulta de los datos fuente.
- `crear_pptx.py` — generador del informe PowerPoint (port de `R/utils_crear_pptx.R`).
- `data/` — Excel fuente (histórico, metas/catálogo de CLUES, procedimientos-personas).
- `.streamlit/config.toml` — tema visual.

## Actualizar los datos

Los tres archivos en `data/` son la fuente de verdad:

- `cubos_completos_plan_justicia.xlsx` — histórico de productividad.
- `metas_planes_justicia.xlsx` — catálogo de CLUES y metas anuales.
- `procedimientos_personas_plan_justicia.xlsx` — procedimientos y personas.

Para actualizarlos: reemplaza el archivo correspondiente en `data/` y haz
`git push`. Streamlit Community Cloud vuelve a desplegar automáticamente con
cada push a la rama conectada — no hace falta ningún paso manual adicional.

## Desplegar en Streamlit Community Cloud

En [share.streamlit.io](https://share.streamlit.io):

1. **New app** → selecciona este repositorio y la rama.
2. **Main file path**: `python/app.py`
3. Deploy.

Streamlit detecta `python/requirements.txt` automáticamente porque está en
la misma carpeta que `app.py`.
