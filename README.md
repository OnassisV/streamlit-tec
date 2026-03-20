# App TEC Norvial

Aplicacion Streamlit para:
- mostrar una portada modular del aplicativo
- cargar un Excel o CSV
- mapear columnas de placa y tiempos
- activar o desactivar pasos del pipeline
- aplicar reglas manuales desde una tabla editable
- exportar `base_limpia`, `casos_eliminados`, `casos_pendientes` y reportes de trazabilidad

## Modulos visibles

- `Inicio`: portada principal del aplicativo.
- `TEC`: modulo operativo actual.
- `Relevamientos`: placeholder visual listo para evolucionar.
- `Auditorias`: placeholder visual listo para evolucionar.
- `Satisfaccion`: placeholder visual listo para evolucionar.
- `Flujogramas`: placeholder visual listo para evolucionar.

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

## Persistencia y seguridad

La app ahora soporta tres modos:

- `APP_STORAGE_MODE = "none"`: sin base de datos.
- `APP_STORAGE_MODE = "sqlite"`: historial local simple.
- `APP_STORAGE_MODE = "mysql"`: historial, login, roles y gestion de usuarios.

Si quieres login y control de usuarios, usa MySQL.

### Variables para MySQL

Configura estas variables en `.streamlit/secrets.toml` o como variables de entorno:

```toml
APP_STORAGE_MODE = "mysql"
APP_MYSQL_HOST = "127.0.0.1"
APP_MYSQL_PORT = 3306
APP_MYSQL_DATABASE = "tec_norvial"
APP_MYSQL_USER = "tu_usuario"
APP_MYSQL_PASSWORD = "tu_password"
APP_MYSQL_CHARSET = "utf8mb4"
APP_MYSQL_CONNECT_TIMEOUT = 10
APP_MYSQL_SSL_DISABLED = false
APP_MYSQL_SSL_VERIFY_CERT = false
APP_MYSQL_SSL_VERIFY_IDENTITY = false
```

La app crea automaticamente las tablas `app_runs`, `app_roles` y `app_users`.

Si tu proveedor de MySQL en la nube exige SSL, puedes completar tambien:

```toml
APP_MYSQL_SSL_CA = "ruta_o_certificado_ca"
APP_MYSQL_SSL_CERT = "ruta_cert_cliente"
APP_MYSQL_SSL_KEY = "ruta_key_cliente"
```

### Login y contrasenas

Las contrasenas no se guardan cifradas de forma reversible. Se almacenan con hash seguro usando `PBKDF2-HMAC-SHA256` y salt aleatorio por usuario.

### Roles incluidos

- `super_admin`: puede cambiar configuraciones generales, crear usuarios, editar usuarios, resetear contrasenas y definir vigencia.
- `admin_operaciones`: puede habilitar o deshabilitar usuarios y definir fechas de vigencia, pero no crear usuarios.
- `analista`: puede procesar informacion usando la configuracion general definida.

### Vigencia de usuarios

Cada usuario puede quedar:

- habilitado o deshabilitado manualmente
- activo desde una fecha
- activo hasta una fecha

Si la fecha aun no empieza o ya vencio, el login se bloquea.

### Datos obligatorios de usuario

Cada usuario ahora requiere:

- `username`
- `full_name`
- `email`
- `phone_number`
- `password`
- `role_key`

## Archivos principales

- `app_tec_norvial_streamlit.py`: interfaz y pipeline principal.
- `app_storage.py`: persistencia opcional.
- `app_auth.py`: autenticacion, hashing de contrasenas y sesion.
- `requirements.txt`: dependencias para despliegue.
# streamlit-tec
