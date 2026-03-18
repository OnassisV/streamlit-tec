# Despliegue en GitHub y Streamlit Community Cloud

## Recomendacion principal

Sube **solo esta carpeta `streamlit_app`** como si fuera tu repositorio.

Eso te conviene porque:
- el repo queda limpio
- no subes notebooks ni resultados pesados que no hacen falta
- la ruta del archivo principal queda simple
- en Streamlit Community Cloud el archivo de entrada sera directamente `app_tec_norvial_streamlit.py`

## Cambio a macOS

Si, puedes continuar sin problemas en macOS.

Esta app:
- usa Python puro
- no depende de rutas Windows
- no necesita Excel instalado
- trabaja con rutas relativas

Mi recomendacion es que en macOS uses **Python 3.12**, porque al **17 de marzo de 2026** Streamlit Community Cloud usa por defecto Python 3.12 y permite elegir la version en el despliegue.  
Fuentes oficiales:
- https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy
- https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies

## Que copiar a la Mac

Copia solo esta carpeta:

- `streamlit_app/`

No necesitas llevar:
- `__pycache__/`
- resultados exportados
- la carpeta raiz vacia `.streamlit`
- la carpeta `data` si esta vacia

## Probar localmente en macOS

Abre Terminal dentro de la carpeta `streamlit_app` y ejecuta:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app_tec_norvial_streamlit.py
```

Si tu Mac no tiene Python instalado:

```bash
brew install python@3.12
```

Si alguna dependencia falla al compilar, instala primero las herramientas de linea de comandos:

```bash
xcode-select --install
```

Referencia oficial:
- https://docs.streamlit.io/get-started/installation/command-line

## Crear el repo en GitHub

Dentro de `streamlit_app`:

```bash
git init
git add .
git commit -m "Initial Streamlit TEC app"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/TU_REPO.git
git push -u origin main
```

## Desplegar gratis en Streamlit Community Cloud

1. Entra a https://share.streamlit.io/
2. Conecta tu cuenta de GitHub.
3. Haz clic en `Create app`.
4. Selecciona:
   - repositorio: tu repo
   - branch: `main`
   - file path: `app_tec_norvial_streamlit.py`
5. En `Advanced settings` elige Python `3.12`.
6. En `Secrets` usa por ahora:

```toml
APP_STORAGE_MODE = "none"
```

7. Despliega la app.

Fuentes oficiales:
- https://docs.streamlit.io/deploy/streamlit-community-cloud/get-started/connect-your-github-account
- https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy
- https://docs.streamlit.io/deploy/streamlit-community-cloud/share-your-app

## Nota sobre persistencia

Para Streamlit Community Cloud gratis, te recomiendo dejar:

```toml
APP_STORAGE_MODE = "none"
```

No te recomiendo usar SQLite dentro de Community Cloud para historial permanente, porque ese entorno no debe asumirse como almacenamiento persistente. Si mas adelante quieres guardar historial real, ahi si conviene conectar una base de datos en hosting.
