# Desplegar en internet (gratis) con OneDrive como base de datos

## 1. Azure — ✅ hecho
App registrada, permisos `Files.ReadWrite` + `User.Read` concedidos.

## 2. Refresh token — ✅ hecho
Ya generado y probado (lectura + escritura funcionando contra tu OneDrive
personal, carpeta `PARQUE ECOS/ECOS DATOS.xlsx`). Los 4 valores
(`MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `MS_REFRESH_TOKEN`) te
los paso por chat cuando vayamos a rellenar Render — nunca se guardan en
este repositorio.

Si alguna vez hay que regenerarlo (p. ej. caduca por inactividad de +90
días), repite esto en esta carpeta:
```bash
set MS_TENANT_ID=xxxx
set MS_CLIENT_ID=xxxx
set MS_CLIENT_SECRET=xxxx
python bootstrap_auth.py
```

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
4. En "Environment", añade estas variables (los valores te los paso por chat):

| Variable | Valor |
|---|---|
| MS_TENANT_ID | (te lo paso por chat) |
| MS_CLIENT_ID | (te lo paso por chat) |
| MS_CLIENT_SECRET | (te lo paso por chat) |
| MS_REFRESH_TOKEN | (te lo paso por chat) |
| APP_PASSWORD | la contraseña que quieras para entrar a la app |
| SECRET_KEY | `trKXv26lgL3ZyPUpMkzX4l4wSILT2N0LUJG77U4tIpE` |

5. Deploy. Cuando termine, Render te da una URL tipo `https://parque-ecografos.onrender.com` — esa es la que usas desde cualquier sitio, con la contraseña que pusiste en `APP_PASSWORD`.

Nota: el plan gratuito de Render "duerme" el servicio tras ~15 min sin uso; la
primera carga tras estar dormido tarda unos 30-50s en despertar, luego va normal.

## Nota sobre el archivo Excel
El Excel "en vivo" que edita la app ahora vive en tu OneDrive **personal**,
en `PARQUE ECOS/ECOS DATOS.xlsx` (no en la carpeta de equipo `CCM Servicio -
Documentos`, que requeriría un permiso de administrador que no estaba
disponible). El original en la carpeta de equipo se quedó tal cual, como
copia congelada — si luego se resuelve el permiso de admin, se puede migrar
la app a leer/escribir ahí en vez de tu OneDrive personal.

## Fotos, historial y averías
Estas funciones usan hojas nuevas del propio Excel (`META`, `HISTORIAL`,
`AVERIAS`) que la app crea sola la primera vez que hacen falta — no hay que
tocar el Excel a mano. Las fotos se guardan como archivos sueltos en tu
OneDrive, en `PARQUE ECOS/fotos/<número de serie>/`, reutilizando las mismas
credenciales de Microsoft Graph que ya tienes configuradas. No hace falta
ninguna variable de entorno nueva para esto.

## 5. Alertas automáticas por email (revisión atrasada)
La app comprueba, cuando se le pide, qué equipos tienen la revisión (PM)
atrasada según su propio campo "Periodicidad de revisión" (editable por
equipo; si se deja en blanco se asume 12 meses), y manda un email SOLO la
primera vez que un equipo pasa a estar atrasado (no todos los días).

### 5.1 Variables de entorno nuevas en Render
Añade estas variables en Render (Settings → Environment), junto a las que ya
tenías:

| Variable | Valor |
|---|---|
| ALERT_TOKEN | una contraseña larga inventada por ti (protege el endpoint de alertas) |
| SMTP_HOST | `smtp.gmail.com` si usas Gmail |
| SMTP_PORT | `587` |
| SMTP_USER | tu cuenta de correo (ej. `tucuenta@gmail.com`) |
| SMTP_PASS | una "contraseña de aplicación" de esa cuenta (no tu contraseña normal) |
| ALERT_EMAIL_TO | a qué email quieres que lleguen los avisos (puede ser el mismo SMTP_USER) |

Si usas Gmail: activa la verificación en 2 pasos en tu cuenta de Google y
genera una "contraseña de aplicación" en myaccount.google.com/apppasswords —
esa es la que va en `SMTP_PASS`, nunca tu contraseña normal de Gmail.

### 5.2 Programar el chequeo diario (GitHub Actions)
Este repositorio ya incluye `.github/workflows/check-alertas.yml`, que llama
todos los días a `/api/alertas/check`. Solo falta darle el token:

1. En GitHub, entra en el repositorio → **Settings** → **Secrets and
   variables** → **Actions**.
2. "New repository secret" → nombre `ALERT_TOKEN`, valor: el mismo que
   pusiste en Render.
3. Listo. Puedes probarlo a mano en la pestaña **Actions** → el workflow
   "Chequeo diario de alertas de revisión" → **Run workflow**.

Para probar sin que se envíe ningún email de verdad ni se marque nada como
"ya avisado", puedes visitar en el navegador (con sesión iniciada no hace
falta, este endpoint usa su propio token):
`https://parque-ecografos.onrender.com/api/alertas/check?token=TU_TOKEN&dry_run=1`
