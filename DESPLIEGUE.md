# Desplegar en internet (gratis) con OneDrive como base de datos

## 1. Azure — ya en marcha
Sigue los pasos que te di en el chat. Al terminar tendrás 3 valores:
`MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`.

## 2. Generar el refresh token (una vez, en tu PC)
En una terminal, dentro de esta carpeta (`webapp`):

```bash
set MS_TENANT_ID=xxxx
set MS_CLIENT_ID=xxxx
set MS_CLIENT_SECRET=xxxx
python bootstrap_auth.py
```

Se abrirá el navegador, inicias sesión con tu cuenta y aceptas los permisos.
Al final te imprime un `MS_REFRESH_TOKEN` — guárdalo, lo necesitas en el paso 4.

## 3. Subir el código a GitHub
```bash
git remote add origin https://github.com/<tu-usuario>/parque-ecografos.git
git branch -M main
git commit -m "Parque de Ecografos"
git push -u origin main
```
(Crea antes el repositorio vacío en github.com — puede ser privado.)

## 4. Crear el servicio en Render.com
1. Regístrate en render.com (gratis) y conecta tu cuenta de GitHub.
2. "New +" → "Web Service" → elige el repo `parque-ecografos`.
3. Build command: `pip install -r requirements.txt` — Start command: se coge del `Procfile` automáticamente.
4. En "Environment", añade estas variables:

| Variable | Valor |
|---|---|
| MS_TENANT_ID | (del paso 1) |
| MS_CLIENT_ID | (del paso 1) |
| MS_CLIENT_SECRET | (del paso 1) |
| MS_REFRESH_TOKEN | (del paso 2) |
| APP_PASSWORD | la contraseña que quieras para entrar a la app |
| SECRET_KEY | `trKXv26lgL3ZyPUpMkzX4l4wSILT2N0LUJG77U4tIpE` |

5. Deploy. Cuando termine, Render te da una URL tipo `https://parque-ecografos.onrender.com` — esa es la que usas desde cualquier sitio, con la contraseña que pusiste en `APP_PASSWORD`.

Nota: el plan gratuito de Render "duerme" el servicio tras ~15 min sin uso; la
primera carga tras estar dormido tarda unos 30-50s en despertar, luego va normal.
