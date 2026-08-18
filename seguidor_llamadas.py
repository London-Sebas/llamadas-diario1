"""
EXTRACCIÓN BD_INSTITUCIONALIZADOS_LLAMADAS
Filtro automático: lunes → Domingo de la semana en curso
Se ejecuta cada Lunes a las 12:30 a.m (hora Colombia) via GitHub Actions

NUEVO: cruce con la hoja BD_INSTITUCIONALIZADOS (mismo archivo) por
'Numero Documento' para traer ESTADO EN NSSC y SUB ESTADO al reporte.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import sys
import json
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import gspread
from email.mime.image import MIMEImage
from google.oauth2.service_account import Credentials

# ============================================================
# CONFIGURACIÓN AUTOMÁTICA DE FECHAS
# Reporte DIARIO
# Extrae todos los registros del día anterior
# ============================================================

hoy = datetime.today()

dia_anterior = hoy - timedelta(days=1)

FECHA_INICIO = (
    os.environ.get("OVERRIDE_FECHA_INICIO")
    or dia_anterior.strftime("%Y-%m-%d")
)

FECHA_FIN = (
    os.environ.get("OVERRIDE_FECHA_FIN")
    or dia_anterior.strftime("%Y-%m-%d")
)

print("=" * 60)
print("📅 CONFIGURACIÓN DEL REPORTE")
print(f"📅 Hoy            : {hoy.strftime('%Y-%m-%d %H:%M:%S')}")
print(f"📅 Día anterior   : {dia_anterior.strftime('%Y-%m-%d')}")
print(f"📅 FECHA_INICIO   : {FECHA_INICIO}")
print(f"📅 FECHA_FIN      : {FECHA_FIN}")
print("=" * 60)

# -- Parámetros generales ─────────────────────────────────────
FILE_ID_LLAMADAS = '1-oqtyFJ4UIwBuuMQPLKqhg3nucJxjZwkzIaXTcuz-yw'
NOMBRE_HOJA = 'BD_INSTITUCIONALIZADOS_LLAMADAS'
NOMBRE_HOJA_INSTITUCIONALIZADOS = 'BD_INSTITUCIONALIZADOS'  # ← nueva hoja para el cruce
RUTA_SALIDA = "output"
FILTRAR_POR_PERFIL = True
PERFILES_VALIDOS = ['AUXILIAR']

FILTRAR_POR_PROFESIONAL = False                                   # ← nuevo flag
PROFESIONALES_VALIDOS = [                                        # ← nueva lista
    'MARTA LUCÍA CASTAÑEDA ALMANZA',
    'LINDA MARIA NIETO ARRIETA',
    'SEBASTIAN VALENCIA OSORIO',
    'santiago cadavid giraldo',
    'yurany andrea ariza paniagua',
    'catalina maturana martinez',
]

# -- Configuración correo Outlook / Office 365 ────────────────
EMAIL_REMITENTE = os.environ.get("EMAIL_REMITENTE", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
EMAIL_DESTINATARIO = os.environ.get("EMAIL_DESTINATARIO", "")
EMAIL_CC = os.environ.get("EMAIL_CC", "")

SMTP_HOST = "smtp.office365.com"
SMTP_PORT = 587

# -- Nombres de columnas ──────────────────────────────────────
COLUMNA_FECHA = 'MARCA TEMPORAL'
COLUMNA_FECHA_LLAMADA = 'Fecha y hora de la llamada'
COLUMNA_DOCUMENTO = 'Numero Documento'
COLUMNA_PROFESIONAL = 'Profesional que realiza la llamada'
COLUMNA_PERFIL = 'Perfil profesional'
COLUMNA_EFECTIVIDAD = 'Efectividad de la llamada'
COLUMNA_MOTIVO = 'Motivo no efectiva'
COLUMNA_OBSERVACIONES = 'Observaciones profesional'

# -- Columnas a traer desde BD_INSTITUCIONALIZADOS ────────────
COLUMNA_ESTADO_NSSC = 'ESTADO EN NSSC'
COLUMNA_SUB_ESTADO = 'SUB ESTADO'

COLUMNAS_REPORTE = [
    COLUMNA_FECHA,
    COLUMNA_FECHA_LLAMADA,
    COLUMNA_DOCUMENTO,
    COLUMNA_PROFESIONAL,
    COLUMNA_PERFIL,
    COLUMNA_EFECTIVIDAD,
    COLUMNA_MOTIVO,
    COLUMNA_OBSERVACIONES,
]


# ============================================================
# 1. AUTENTICACIÓN Y DESCARGA DESDE GOOGLE SHEETS
# ============================================================
def load_sheet(file_id: str, sheet_name: str) -> pd.DataFrame:
    """Descarga una hoja de Google Sheets usando Service Account."""

    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise EnvironmentError(
            "❌ Variable de entorno GOOGLE_CREDENTIALS no encontrada.\n"
            "   Agrega el contenido del .json como secret en GitHub."
        )

    creds_dict = json.loads(creds_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(file_id)
    worksheet = spreadsheet.worksheet(sheet_name)

    data = worksheet.get_all_records(expected_headers=[])
    df = pd.DataFrame(data)

    print(f"✅ Hoja cargada: {df.shape[0]:,} filas × {df.shape[1]} columnas")
    print(f"📌 Columnas disponibles: {df.columns.tolist()}\n")
    return df


df_llamadas = load_sheet(FILE_ID_LLAMADAS, NOMBRE_HOJA)


# ============================================================
# 2. VALIDAR COLUMNAS
# ============================================================
cols_existentes = [c for c in COLUMNAS_REPORTE if c in df_llamadas.columns]
cols_faltantes = [c for c in COLUMNAS_REPORTE if c not in df_llamadas.columns]

if cols_faltantes:
    print(f"⚠️ Columnas no encontradas: {cols_faltantes}")

df_base = df_llamadas[cols_existentes].copy()


# ============================================================
# 3. CONVERSIÓN Y LIMPIEZA DE FECHA
# ============================================================
fecha_inicio = pd.to_datetime(FECHA_INICIO)
fecha_fin = pd.to_datetime(FECHA_FIN)

df_base[COLUMNA_FECHA] = pd.to_datetime(
    df_base[COLUMNA_FECHA].astype(str).str.strip(),
    errors='coerce',
    dayfirst=True
)

total_filas = len(df_base)
nulas = df_base[COLUMNA_FECHA].isna().sum()
con_fecha = total_filas - nulas

print(f"📊 Total filas descargadas : {total_filas:,}")
print(f"✅ Filas con fecha válida  : {con_fecha:,}")
print(f"⚠️ Filas con fecha nula    : {nulas:,}  ← se excluirán")


# ============================================================
# 4. FILTROS: RANGO DE FECHAS + PERFIL PROFESIONAL
# ============================================================
df_base = df_base.dropna(how='all').reset_index(drop=True)
columnas_clave = [c for c in [COLUMNA_FECHA, COLUMNA_PERFIL, COLUMNA_DOCUMENTO] if c in df_base.columns]
df_base = df_base.dropna(subset=columnas_clave, how='all').reset_index(drop=True)

print(f"\n📋 Filas tras limpiar vacíos: {len(df_base):,}")

fecha_inicio_dt = pd.to_datetime(FECHA_INICIO)
fecha_fin_exclusiva = pd.to_datetime(FECHA_FIN) + timedelta(days=1)

mask_fechas = (
    df_base[COLUMNA_FECHA].notna() &
    (df_base[COLUMNA_FECHA] >= fecha_inicio_dt) &
    (df_base[COLUMNA_FECHA] < fecha_fin_exclusiva)
)

if FILTRAR_POR_PERFIL:
    mask_perfil = (
        df_base[COLUMNA_PERFIL]
        .astype(str).str.strip().str.upper()
        .isin(PERFILES_VALIDOS)
    )
else:
    mask_perfil = pd.Series(True, index=df_base.index)

# ── máscara profesionales ─────────────────────────────────
if FILTRAR_POR_PROFESIONAL:
    mask_profesional = (
        df_base[COLUMNA_PROFESIONAL]
        .astype(str).str.strip().str.upper()
        .isin([p.upper() for p in PROFESIONALES_VALIDOS])
    )
else:
    mask_profesional = pd.Series(True, index=df_base.index)

df_reporte = (
    df_base
    .loc[mask_fechas & mask_perfil & mask_profesional]
    .sort_values(COLUMNA_FECHA)
    .reset_index(drop=True)
)


# ============================================================
# LIMPIEZA DOCUMENTOS
# Evita notación científica en Excel / normaliza para el cruce
# ============================================================
def limpiar_documento(serie: pd.Series) -> pd.Series:
    return (
        serie
        .astype(str)
        .str.replace('.0', '', regex=False)
        .str.strip()
    )

df_reporte[COLUMNA_DOCUMENTO] = limpiar_documento(df_reporte[COLUMNA_DOCUMENTO])


# ============================================================
# 4.1 CRUCE CON BD_INSTITUCIONALIZADOS
# Trae ESTADO EN NSSC y SUB ESTADO por Numero Documento
# ============================================================
df_institucionalizados = load_sheet(FILE_ID_LLAMADAS, NOMBRE_HOJA_INSTITUCIONALIZADOS)

cols_cruce = [COLUMNA_DOCUMENTO, COLUMNA_ESTADO_NSSC, COLUMNA_SUB_ESTADO]
cols_cruce_existentes = [c for c in cols_cruce if c in df_institucionalizados.columns]
cols_cruce_faltantes = [c for c in cols_cruce if c not in df_institucionalizados.columns]

if cols_cruce_faltantes:
    print(f"⚠️ Columnas no encontradas en {NOMBRE_HOJA_INSTITUCIONALIZADOS}: {cols_cruce_faltantes}")

df_institucionalizados = df_institucionalizados[cols_cruce_existentes].copy()

if COLUMNA_DOCUMENTO in df_institucionalizados.columns:
    df_institucionalizados[COLUMNA_DOCUMENTO] = limpiar_documento(df_institucionalizados[COLUMNA_DOCUMENTO])

    duplicados_inst = df_institucionalizados[COLUMNA_DOCUMENTO].duplicated().sum()
    if duplicados_inst:
        print(
            f"⚠️ {duplicados_inst:,} documentos duplicados en {NOMBRE_HOJA_INSTITUCIONALIZADOS}, "
            "se conserva el primer registro para el cruce"
        )
        df_institucionalizados = df_institucionalizados.drop_duplicates(
            subset=COLUMNA_DOCUMENTO, keep='first'
        )

    df_reporte = df_reporte.merge(
        df_institucionalizados,
        on=COLUMNA_DOCUMENTO,
        how='left'
    )

    if COLUMNA_ESTADO_NSSC in df_reporte.columns:
        sin_cruce = df_reporte[COLUMNA_ESTADO_NSSC].isna().sum()
        print(f"🔗 Registros del reporte sin cruce en {NOMBRE_HOJA_INSTITUCIONALIZADOS}: {sin_cruce:,}")
else:
    print(f"⚠️ No se pudo cruzar: '{COLUMNA_DOCUMENTO}' no está en {NOMBRE_HOJA_INSTITUCIONALIZADOS}")


# ============================================================
# 5. RESUMEN EN CONSOLA
# ============================================================
print("=" * 55)
print(f"📋 Rango consultado : {FECHA_INICIO} -> {FECHA_FIN}")
print(f"📊 Total registros  : {len(df_reporte):,}")

if COLUMNA_PERFIL in df_reporte.columns:
    conteo_perfiles = df_reporte[COLUMNA_PERFIL].value_counts().to_dict()
    print(f"👤 Por perfil       : {conteo_perfiles}")
print("=" * 55)

if COLUMNA_PROFESIONAL in df_reporte.columns:
    conteo_prof = df_reporte[COLUMNA_PROFESIONAL].value_counts().to_dict()
    print(f"🧑‍⚕️ Por profesional  : {conteo_prof}")
print("=" * 55)   

if df_reporte.empty:
    print("⚠️ Sin registros en el período. Se generará un archivo vacío.")

# ============================================================
# 5.1 RESUMEN DE REGISTROS POR PROFESIONAL
# ============================================================

if COLUMNA_PROFESIONAL in df_reporte.columns and not df_reporte.empty:

    resumen_profesionales = (
        df_reporte[COLUMNA_PROFESIONAL]
        .astype(str)
        .str.strip()
        .replace("", "SIN PROFESIONAL")
        .value_counts()
        .rename_axis("Profesional")
        .reset_index(name="Registros")
    )

else:

    resumen_profesionales = pd.DataFrame(
        columns=["Profesional", "Registros"]
    )

print("\n📊 REGISTROS POR PROFESIONAL")
print(resumen_profesionales.to_string(index=False))

# ============================================================
# 6. EXPORTAR A EXCEL
# ============================================================
os.makedirs(RUTA_SALIDA, exist_ok=True)

nombre_archivo = f"BD_Llamadas_{FECHA_INICIO}_al_{FECHA_FIN}.xlsx"
ruta_completa = os.path.join(RUTA_SALIDA, nombre_archivo)

with pd.ExcelWriter(ruta_completa, engine='openpyxl') as writer:
    df_reporte.to_excel(writer, index=False, sheet_name='Reporte')

    ws = writer.sheets['Reporte']

    # Buscar columna documento
    col_doc_idx = df_reporte.columns.get_loc(COLUMNA_DOCUMENTO) + 1

    # Formato texto
    for row in range(2, len(df_reporte) + 2):
        ws.cell(row=row, column=col_doc_idx).number_format = '@'
        
print(f"\n✅ Archivo exportado en: {ruta_completa}")

# Exponer variables para pasos posteriores de GitHub Actions
with open(os.environ.get("GITHUB_ENV", "/dev/null"), "a") as f:
    f.write(f"NOMBRE_EXCEL={nombre_archivo}\n")
    f.write(f"RUTA_EXCEL={ruta_completa}\n")


# ============================================================
# 5.2 TABLA HTML PARA EL CORREO
# ============================================================

if not resumen_profesionales.empty:

    filas_tabla = ""

    for _, fila in resumen_profesionales.iterrows():

        profesional = fila["Profesional"]
        cantidad = int(fila["Registros"])

        filas_tabla += f"""
        <tr>
            <td style="
                padding:8px 12px;
                border:1px solid #d9d9d9;
                text-align:left;
            ">
                {profesional}
            </td>

            <td style="
                padding:8px 12px;
                border:1px solid #d9d9d9;
                text-align:center;
                font-weight:bold;
            ">
                {cantidad}
            </td>
        </tr>
        """

    tabla_profesionales_html = f"""
    <table style="
        border-collapse:collapse;
        width:100%;
        max-width:700px;
        font-family:Calibri, Arial, sans-serif;
        font-size:14px;
        margin-top:10px;
        margin-bottom:20px;
    ">

        <thead>
            <tr style="background-color:#f2f2f2;">

                <th style="
                    padding:9px 12px;
                    border:1px solid #d9d9d9;
                    text-align:left;
                ">
                    Profesional
                </th>

                <th style="
                    padding:9px 12px;
                    border:1px solid #d9d9d9;
                    text-align:center;
                ">
                    Registros
                </th>

            </tr>
        </thead>

        <tbody>
            {filas_tabla}
        </tbody>

    </table>
    """

else:

    tabla_profesionales_html = """
    <p>
        <strong>No se encontraron registros para el día reportado.</strong>
    </p>
    """

# ============================================================
# 7. ENVÍO POR CORREO — Outlook / Office 365 (STARTTLS)
# ============================================================
def enviar_correo(
    ruta_adjunto: str,
    nombre_adjunto: str,
    total_registros: int,
    fecha_inicio: str,
    fecha_fin: str,
    tabla_profesionales_html: str,
) -> None:
    """Envía el Excel generado como adjunto vía Outlook SMTP."""

    # Validar credenciales
    if not all([EMAIL_REMITENTE, EMAIL_PASSWORD, EMAIL_DESTINATARIO]):
        print(
            "⚠️ No se enviará correo: faltan EMAIL_REMITENTE, "
            "EMAIL_PASSWORD o EMAIL_DESTINATARIO en los secrets."
        )
        return

    # ── Construir mensaje ─────────────────────────────────────
    msg = MIMEMultipart()
    msg["From"] = EMAIL_REMITENTE
    msg["To"] = EMAIL_DESTINATARIO
    msg["Subject"] = f"📊 Reporte BD Llamadas | {fecha_inicio} al {fecha_fin}"

    if EMAIL_CC:
        msg["Cc"] = EMAIL_CC

    cuerpo_html = f"""
    <html>
    <body style='font-family: Calibri, Arial, sans-serif; color: #333; line-height: 1.6;'>
    
        <p>Buenos días,</p>
    
        <p>
            El presente reporte se genera con el fin de dar cumplimiento
            a la solicitud de información de las llamadas realizadas.
        </p>
    
        <p>
            A continuación, se presenta el resumen de registros ingresados
            durante el día <strong>{fecha_inicio}</strong>:
        </p>
    
        <h3 style="margin-bottom:5px;">
            📊 Registros por profesional
        </h3>
    
        {tabla_profesionales_html}
    
        <p>
            <strong>Total de registros del día: {total_registros:,}</strong>
        </p>
    
        <p>
            Se adjunta el archivo Excel con el detalle de los registros.
            Si se requieren cambios, agregar información o cualquier otra
            modificación, hacérmelo saber para darle respuesta lo más pronto posible.
        </p>
    
        <p>
            Muchas gracias por su atención.
        </p>
    
        <p>
            Buen día. 😊
        </p>
    
        <br>
    
        <p style='font-size:0.85em; color:#888;'>
            Fecha del reporte: {fecha_inicio}
            &nbsp;|&nbsp;
            Registros: {total_registros:,}
            &nbsp;|&nbsp;
            Archivo: {nombre_adjunto}
        </p>
    
        <!-- FIRMA -->
        <img src="cid:firma_digital"
             style="width:300px; max-width:100%;">
    
        <br><br>
    
    </body>
    </html>
    """

    msg.attach(MIMEText(cuerpo_html, "html"))
    # ── Adjuntar firma digital embebida ─────────────────────
    ruta_firma = "Firma NSSC Correos.png"
    
    with open(ruta_firma, "rb") as img:
        firma = MIMEImage(img.read())
    
    firma.add_header("Content-ID", "<firma_digital>")
    firma.add_header("Content-Disposition", "inline", filename="Firma NSSC Correos.png")
    
    msg.attach(firma)
    # ── Adjuntar el Excel ─────────────────────────────────────
    with open(ruta_adjunto, "rb") as f:
        parte = MIMEBase("application", "octet-stream")
        parte.set_payload(f.read())

    encoders.encode_base64(parte)
    part_header = f'attachment; filename="{nombre_adjunto}"'
    parte.add_header("Content-Disposition", part_header)
    msg.attach(parte)

    # ── Enviar vía STARTTLS ───────────────────────────────────
    destinatarios = [EMAIL_DESTINATARIO]
    if EMAIL_CC:
        destinatarios += [addr.strip() for addr in EMAIL_CC.split(",") if addr.strip()]

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
            server.sendmail(EMAIL_REMITENTE, destinatarios, msg.as_string())

        print(f"📧 Correo enviado a: {', '.join(destinatarios)}")

    except smtplib.SMTPAuthenticationError:
        print("❌ Error de autenticación SMTP. Revisa EMAIL_REMITENTE y EMAIL_PASSWORD.")
        sys.exit(1)
    except smtplib.SMTPException as e:
        print(f"❌ Error SMTP al enviar correo: {e}")
        sys.exit(1)


enviar_correo(
    ruta_adjunto=ruta_completa,
    nombre_adjunto=nombre_archivo,
    total_registros=len(df_reporte),
    fecha_inicio=FECHA_INICIO,
    fecha_fin=FECHA_FIN,
    tabla_profesionales_html=tabla_profesionales_html,
)
