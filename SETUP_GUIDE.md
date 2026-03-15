# 🏗️ BELFOR Equipment Log — Setup Guide
## De cero a link en ~10 minutos, sin IT

---

## Paso 1 — Crea cuenta en GitHub (gratis)
1. Ve a **github.com**
2. Click "Sign up" → email, password, username
3. Confirma tu email

---

## Paso 2 — Sube el código a GitHub
1. Haz login en github.com
2. Click **"New repository"** (botón verde)
3. Nombre: `belfor-equipment-log`
4. Selecciona **Private** (solo tú lo ves)
5. Click **"Create repository"**
6. En la página del repo, click **"uploading an existing file"**
7. Arrastra estos 2 archivos:
   - `app.py`
   - `requirements.txt`
8. Click **"Commit changes"**

---

## Paso 3 — Consigue tu API Key de Anthropic
1. Ve a **console.anthropic.com**
2. Login con tu cuenta (o crea una)
3. Click **"API Keys"** → **"Create Key"**
4. Copia la key (empieza con `sk-ant-...`)
5. **Guárdala** — solo se muestra una vez

> 💡 El costo es muy bajo: ~$0.01-0.05 por chat procesado

---

## Paso 4 — Deploy en Streamlit Cloud (gratis)
1. Ve a **share.streamlit.io**
2. Login con tu cuenta de GitHub
3. Click **"New app"**
4. Repository: `belfor-equipment-log`
5. Branch: `main`
6. Main file: `app.py`
7. Click **"Advanced settings"**
8. En **Secrets**, pega esto (con TU key real):
   ```
   ANTHROPIC_API_KEY = "sk-ant-TU-KEY-AQUI"
   ```
9. Click **"Deploy"**

---

## Paso 5 — ¡Listo!
Streamlit te da un link como:
```
https://belfor-equipment-log.streamlit.app
```

Guárdalo en tus favoritos. Cada vez que quieras generar un log:
1. Abre el link
2. Sube el .zip de WhatsApp
3. Escribe el nombre del proyecto
4. Click "Procesar" → descarga el Excel

---

## ¿Cómo actualizar la app en el futuro?
Si quieres cambiar algo, edita `app.py` directamente en GitHub
y Streamlit se actualiza solo en ~1 minuto.

---

## Soporte
¿Algo no funciona? Pregúntale a Claude con el mensaje de error
y te ayuda a resolverlo en segundos.
