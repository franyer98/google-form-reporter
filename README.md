# Google Form Reporter

Trae las respuestas de un formulario de Google y genera un reporte automático en Excel con datos crudos, resumen estadístico y gráficos. Diseñado para correr todos los días sin intervención manual.

## Qué genera

Cada ejecución produce tres archivos:
- **Excel (`.xlsx`)**: hoja "Datos" con las respuestas crudas y hoja "Resumen" con conteos, estadísticas y gráficos de barras nativos de Excel.
- **PDF (`.pdf`)**: reporte listo para compartir, con resumen por pregunta, gráficos y la tabla completa de datos.
- **HTML (`site/index.html`)**: la misma información para verla directamente en el navegador, sin descargar nada. Cuando corre por GitHub Actions, este archivo se publica automáticamente en GitHub Pages (ver sección 5).

## 1. Preparar el formulario y la hoja de cálculo

1. Abre tu Google Form → pestaña **Respuestas** → ícono de Sheets → **Crear hoja de cálculo**.
2. Copia el ID de la hoja desde la URL: `https://docs.google.com/spreadsheets/d/EL_ID/edit`.
3. Anota el nombre de la pestaña de respuestas (por defecto: `Respuestas de formulario 1`).

## 2. Crear la cuenta de servicio (acceso sin login manual)

1. Ve a [Google Cloud Console](https://console.cloud.google.com/) → crea un proyecto.
2. Habilita la **Google Sheets API**.
3. Ve a **Credenciales → Crear credenciales → Cuenta de servicio**.
4. Dentro de la cuenta de servicio, genera una **clave JSON** y descárgala como `credentials.json`.
5. Copia el email de la cuenta de servicio (termina en `...gserviceaccount.com`).
6. Abre tu Google Sheet → **Compartir** → pega ese email con permiso de **Lector**.

## 3. Instalación local

```bash
git clone <tu-repo>
cd google_form_reporter
pip install -r requirements.txt
cp .env.example .env
```

Edita `.env` con tu `SHEET_ID` y `WORKSHEET_NAME`. Coloca `credentials.json` en la raíz del proyecto.

Prueba manualmente:

```bash
python report_generator.py
```

Esto crea `reportes/reporte_YYYY-MM-DD_HHMM.xlsx` y `reportes/reporte_YYYY-MM-DD_HHMM.pdf`.

## 4. Automatizar la ejecución diaria

### Opción A — GitHub Actions (recomendada, no necesitas dejar el PC prendido)

1. Sube este proyecto a un repositorio de GitHub.
2. Ve a **Settings → Secrets and variables → Actions** y crea:
   - `GOOGLE_CREDENTIALS_JSON_B64`: contenido de `credentials.json` codificado en base64 (`base64 -w0 credentials.json` en Linux/Mac, o `certutil -encode credentials.json tmp.b64` en Windows).
   - `SHEET_ID`: el ID de tu hoja.
   - `WORKSHEET_NAME`: el nombre de la pestaña.
3. El workflow en `.github/workflows/daily_report.yml` corre todos los días a las 8:00 am (hora Colombia) y deja el Excel y el PDF descargables en la pestaña **Actions → artifacts** (artefacto llamado "reportes"), conservándolos por 14 días.
4. Puedes ajustar el horario cambiando la línea `cron` (formato UTC).

### Opción B — Cron (Linux/Mac)

```bash
crontab -e
# Agrega esta línea para correr todos los días a las 8:00 am:
0 8 * * * cd /ruta/al/proyecto && /usr/bin/python3 report_generator.py >> log.txt 2>&1
```

### Opción C — Task Scheduler (Windows)

1. Abre **Programador de tareas** → **Crear tarea básica**.
2. Desencadenador: **Diariamente**, elige la hora.
3. Acción: **Iniciar un programa** → selecciona `python.exe` y en argumentos pon la ruta a `report_generator.py`.

## 5. Ver el reporte en una página web (GitHub Pages)

El workflow publica automáticamente `site/index.html` en GitHub Pages en cada corrida, así podés ver el último reporte desde el navegador sin descargar nada.

**Paso único de configuración** (una sola vez, después de subir el proyecto a GitHub):

1. Ve a **Settings → Pages**.
2. En **Build and deployment → Source**, elegí **GitHub Actions**.
3. Corré el workflow una vez (**Actions → Reporte diario del formulario → Run workflow**).
4. La URL de tu página quedará publicada en **Settings → Pages** (formato `https://<usuario>.github.io/<repo>/`) y también aparece en el resumen de cada ejecución del workflow.

> ⚠️ Si el repositorio es público, esa URL también es pública: cualquiera con el link puede ver las respuestas del formulario (no aparece indexada en buscadores ni listada en ningún lado, pero no está protegida por login). Si el formulario recolecta datos sensibles, hacé el repositorio privado — GitHub Pages privado requiere un plan Pro/Team/Enterprise.

## 6. Próximas mejoras posibles

- Enviar el reporte por correo automáticamente (con `smtplib` o SendGrid).
- Subir el Excel a Google Drive en vez de solo guardarlo localmente.
- Guardar histórico en una base de datos para comparar tendencias entre reportes.

## Estructura del proyecto

```
google_form_reporter/
├── report_generator.py
├── requirements.txt
├── .env.example
└── .github/workflows/daily_report.yml
```
