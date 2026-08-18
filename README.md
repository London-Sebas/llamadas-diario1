# Seguidor Llamadas — Automatización Semanal

Script que extrae la hoja `BD_INSTITUCIONALIZADOS_LLAMADAS` de Google Sheets,
filtra la **semana anterior**, y exporta un `.xlsx` como artefacto en GitHub Actions.

---

## ⚙️ Configuración inicial (una sola vez)

### 1. Crear un repositorio en GitHub
Sube todos estos archivos a un repo (puede ser privado).

```
mi-repo/
├── .github/
│   └── workflows/
│       └── reporte_semanal.yml
├── seguidor_llamadas.py
├── requirements.txt
└── README.md
```

---

### 2. Agregar el secret `GOOGLE_CREDENTIALS`

El script necesita las credenciales de tu **Service Account** de Google para leer el Sheet.

1. Ve a tu repo en GitHub → **Settings** → **Secrets and variables** → **Actions**
2. Haz clic en **New repository secret**
3. Nombre: `GOOGLE_CREDENTIALS`
4. Valor: **todo el contenido** del archivo `.json` de tu Service Account  
   (es el archivo que descargaste de Google Cloud Console al crear la Service Account)
5. Guarda.

> ⚠️ Nunca subas el `.json` directamente al repositorio.

---

### 3. Verificar que la Service Account tenga acceso al Google Sheet

En el Google Sheet comparte el archivo con el email de tu Service Account  
(se ve en el `.json` como `"client_email": "nombre@proyecto.iam.gserviceaccount.com"`).  
Permisos necesarios: **Lector**.

---

## 🕐 Cuándo se ejecuta

| Trigger | Detalle |
|---|---|
| **Automático** | Todos los lunes a las **06:00 hora Colombia** (11:00 UTC) |
| **Manual** | GitHub → Actions → *Seguidor Llamadas* → **Run workflow** |

En la ejecución manual puedes opcionalmente ingresar `FECHA_INICIO` y `FECHA_FIN`
en formato `YYYY-MM-DD` para extraer un rango personalizado.

---

## 📥 Descargar el Excel generado

1. GitHub → **Actions**
2. Clic en la ejecución más reciente de *Seguidor Llamadas — Reporte Semanal*
3. Sección **Artifacts** → descargar el `.zip` (contiene el `.xlsx`)

Los artefactos se conservan **30 días**.

---

## 🔧 Ajustes comunes

| Qué cambiar | Dónde |
|---|---|
| Hora de ejecución | `cron` en `.github/workflows/reporte_semanal.yml` |
| Cambiar a quincena automática | Editar lógica de fechas en `seguidor_llamadas.py` |
| Filtrar por perfil profesional | `FILTRAR_POR_PERFIL = True` en `seguidor_llamadas.py` |
| ID del Google Sheet | `FILE_ID_LLAMADAS` en `seguidor_llamadas.py` |
