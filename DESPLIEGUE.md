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
