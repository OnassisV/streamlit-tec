# App TEC Norvial

Aplicacion Streamlit para:
- cargar un Excel o CSV
- mapear columnas de placa y tiempos
- activar o desactivar pasos del pipeline
- aplicar reglas manuales desde una tabla editable
- exportar `base_limpia`, `casos_eliminados`, `casos_pendientes` y reportes de trazabilidad

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app_tec_norvial_streamlit.py
```

## Publicar gratis en Streamlit Community Cloud

1. Sube esta carpeta a un repositorio publico en GitHub.
2. Entra a Streamlit Community Cloud.
3. Crea una nueva app apuntando al archivo `app_tec_norvial_streamlit.py`.
4. Usa Python por defecto y deja `requirements.txt` como dependencia.

## Persistencia

La app funciona sin base de datos.

Hoy el modo recomendado para una publicacion gratis es:
- `APP_STORAGE_MODE = "none"`

Si luego la corres en tu propia PC o en un servidor tuyo, puedes activar historial local:
- `APP_STORAGE_MODE = "sqlite"`
- `APP_SQLITE_PATH = "data/app_runs.db"`

La capa de persistencia ya esta aislada en `app_storage.py`, asi que mas adelante podemos cambiar a PostgreSQL o MySQL sin rehacer la app.

## Archivos principales

- `app_tec_norvial_streamlit.py`: interfaz y pipeline principal.
- `app_storage.py`: persistencia opcional.
- `requirements.txt`: dependencias para despliegue.
# streamlit-tec
