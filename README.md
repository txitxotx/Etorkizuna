# 📊 Portfolio Dashboard — Auto-actualizable desde Excel

Dashboard de inversión que se actualiza automáticamente cada vez que modificas y subes el Excel a GitHub.

---

## 🚀 Cómo funciona

```
Tú editas el Excel  →  Subes a GitHub  →  GitHub Actions ejecuta parse_excel.py
→  Genera data.json  →  Vercel despliega  →  Dashboard actualizado en ~2 min
```

---

## ⚡ Configuración inicial (una sola vez, ~10 minutos)

### Paso 1 — Subir a GitHub

1. Ve a [github.com](https://github.com) → **New repository**
2. Nombre: `portfolio-dashboard` (o el que quieras), **privado** recomendado
3. Sube todos estos archivos (arrastra la carpeta o usa GitHub Desktop):
   ```
   portfolio_cuadro_mandos.xlsx    ← tu Excel
   parse_excel.py
   requirements.txt
   vercel.json
   public/
     index.html
     data.json
   .github/
     workflows/
       update-dashboard.yml
   ```

> **Con GitHub Desktop** (más fácil): descarga [desktop.github.com](https://desktop.github.com),
> arrastra la carpeta del proyecto y haz "Publish repository".

### Paso 2 — Conectar Vercel

1. Ve a [vercel.com](https://vercel.com) → **Add New Project**
2. Importa tu repositorio de GitHub
3. Configuración:
   - **Framework Preset**: `Other`
   - **Output Directory**: `public`
   - **Build Command**: *(dejar vacío)*
4. Haz clic en **Deploy** ✅

Tu dashboard ya está online con la URL que te da Vercel.

### Paso 3 — Verificar que GitHub Actions funciona

1. En tu repositorio de GitHub → pestaña **Actions**
2. Deberías ver el workflow `📊 Actualizar Dashboard`
3. Si hay un tick verde ✅ todo funciona

---

## 📝 Flujo de trabajo diario

### Actualizar precios (lo más habitual)

1. Abre `portfolio_cuadro_mandos.xlsx`
2. Ve a la hoja **📋 ACTIVOS**
3. Actualiza la columna **G — Precio Hoy** (en amarillo) con los precios actuales
4. Guarda el archivo
5. Sube el Excel a GitHub (GitHub Desktop → Commit → Push, o arrastrando el archivo)
6. Espera ~2 minutos → tu dashboard en Vercel se actualiza solo

### Cambiar parámetros (tasa libre de riesgo, rentabilidad esperada, etc.)

1. Ve a la hoja **⚙️ INPUTS**
2. Modifica los valores en **azul** (tasa libre de riesgo, primas, pesos objetivo...)
3. Guarda y sube igual que antes

### Añadir o eliminar activos

1. En **📋 ACTIVOS**, añade o elimina filas manteniendo el formato
2. `parse_excel.py` detecta automáticamente filas 5–29 con datos
3. Sube → GitHub Actions regenera → dashboard actualizado

---

## 🖥️ Uso local (sin internet)

Si quieres ver el dashboard en tu ordenador sin publicarlo:

```bash
# Instalar dependencias (solo la primera vez)
pip install openpyxl

# Generar data.json desde el Excel
python parse_excel.py

# Abrir el dashboard
# macOS:
open public/index.html
# Windows:
start public/index.html
# Linux:
xdg-open public/index.html
```

---

## 📁 Estructura del proyecto

```
portfolio-dashboard/
│
├── portfolio_cuadro_mandos.xlsx   ← TU EXCEL (edita esto)
├── parse_excel.py                 ← lee el Excel, genera data.json
├── requirements.txt               ← dependencias Python
├── vercel.json                    ← configuración del servidor
│
├── public/
│   ├── index.html                 ← el dashboard (no tocar)
│   └── data.json                  ← datos generados automáticamente
│
└── .github/
    └── workflows/
        └── update-dashboard.yml  ← automatización de GitHub
```

---

## ❓ Preguntas frecuentes

**¿Debo subir el Excel cada vez?**
Sí, GitHub necesita detectar el cambio para lanzar el workflow. Basta con guardarlo y hacer push.

**¿Cuánto tarda en actualizarse?**
Normalmente entre 60 y 120 segundos desde que haces push.

**¿Puede mi repositorio ser privado?**
Sí. Vercel puede conectarse a repositorios privados de GitHub.

**¿Qué pasa si el workflow falla?**
Ve a GitHub → Actions → haz clic en el workflow fallido para ver el error. El problema más común es que el nombre de una hoja del Excel no coincide con el esperado.

**¿Cómo añado más activos?**
Simplemente añade filas en **📋 ACTIVOS** antes de la fila 30, manteniendo el mismo formato de columnas. El parser lee filas 5–29.

**¿Puedo cambiar el nombre del Excel?**
Sí, pero actualiza también la línea `paths:` en `.github/workflows/update-dashboard.yml`.

---

## 🔧 Hojas del Excel y lo que lee el script

| Hoja | Qué lee |
|------|---------|
| `📋 ACTIVOS` | Filas 5–29: nombre, categoría, títulos, precio compra, precio hoy, rentabilidades |
| `⚙️ INPUTS` | Tasa libre de riesgo, pesos objetivo, rentabilidades esperadas, volatilidades |
| `🔍 ANÁLISIS` | Escenarios de estrés (filas 26–30) |

---

*Generado automáticamente — Portfolio Dashboard v2*
