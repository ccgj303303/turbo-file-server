# Turbo File Server — Deploy en Railway
## Tiempo estimado: 10 minutos

---

## PASO 1 — Instalar GitHub Desktop (si no lo tienes)
https://desktop.github.com/ — descarga e instala en tu Mac.

---

## PASO 2 — Crear repositorio en GitHub
1. Abre GitHub Desktop
2. File → New Repository
3. Name: `turbo-file-server`
4. Local path: elige una carpeta en tu Mac
5. Click "Create Repository"
6. Copia todos los archivos de esta carpeta al repositorio
7. Commit message: "Initial deploy"
8. Click "Publish repository" → PUBLIC

---

## PASO 3 — Deploy en Railway
1. Ve a https://railway.app y crea cuenta (usa tu GitHub)
2. Click "New Project" → "Deploy from GitHub repo"
3. Selecciona `turbo-file-server`
4. Railway detecta automáticamente Python + Flask
5. Espera 2-3 minutos → verás "✅ Deployed"

---

## PASO 4 — Configurar variable secreta
1. En Railway → tu proyecto → "Variables"
2. Click "New Variable"
3. Name: `TURBO_API_TOKEN`
4. Value: pon cualquier contraseña larga, ejemplo: `TurboSecure2026!Chris`
5. Click "Add"

---

## PASO 5 — Obtener tu URL
1. En Railway → tu proyecto → "Settings" → "Domains"
2. Click "Generate Domain"
3. Copia la URL, ejemplo: `https://turbo-file-server-production.up.railway.app`
4. Prueba: abre esa URL + `/health` en tu browser
5. Debes ver: `{"service": "Turbo File Server", "status": "ok"}`

---

## PASO 6 — Configurar n8n
En n8n → Settings → Environment Variables, agrega:
- `TURBO_SERVER_URL` = tu URL de Railway (sin "/" al final)
- `TURBO_API_TOKEN` = la misma contraseña del Paso 4

---

## PASO 7 — Probar
En Telegram, dile a Turbo:
- "Hazme una presentación del SEC CAP para el board, 6 slides"
- "Crea un Excel con un portfolio de 5 préstamos de ejemplo"

---

## Costo
- Railway Hobby Plan: $5/mes
- Incluye 500 horas de ejecución — más que suficiente para uso personal
