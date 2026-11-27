# -*- coding: utf-8 -*-
"""
CRM TECNOLOGÍAS EN SALUD · Validación de precios (Circular 022 + SISMED 2024/2019), seguimiento de PQRS y asistente de IA
Autor: Ivan Santiago Herrera Baquero · Versión 6.1 (generado desde el notebook de Colab)

Uso en Colab:  !python "/content/drive/MyDrive/TECNOLOGIAS EN SALUD/proyecto_validacion_precios.py"   o pegar el contenido en una celda.
Cada bloque "# %%" corresponde a una celda del notebook.
"""
try:
    from IPython.display import display
except ImportError:
    display = print

# %% [markdown]
# # CRM Tecnologías en Salud · Validación de precios de medicamentos y seguimiento de PQRS
#
# [![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/ivanherrerabaquero-dev/validacion-precios-medicamentos/blob/main/notebooks/proyecto_validacion_precios.ipynb)
# [![Portal](https://img.shields.io/badge/Portal-GitHub%20Pages-1f4e79)](https://ivanherrerabaquero-dev.github.io/validacion-precios-medicamentos/)
#
# **Autor:** Ivan Santiago Herrera Baquero · Administrador de Empresas · Gestión de PQRS y análisis de datos en salud
#
# ## Contexto
#
# **CRM TECNOLOGÍAS EN SALUD** (empresa ficticia) es un gestor farmacéutico: distribuye medicamentos a clientes institucionales mediante **licitaciones y convenios** y opera **puntos de dispensación** para pacientes. El área de PQRS y apoyo a licitaciones:
#
# | Función del cargo | Paso del notebook |
# |---|---|
# | Consolidación de datos | 3 (lectura robusta de cualquier archivo) |
# | Validación de ofertas frente al precio máximo (Circular 022) y al SISMED | 4 |
# | Recepción y seguimiento de PQR | 5 |
# | Atención a usuarios internos de licitaciones y convenios | 5 |
# | Metas regulatorias (SIC, Supersalud, INVIMA) | 5B |
# | Generación de informes ejecutivos | 6, 7 y 8B (Excel, HTML y portal web) |
# | Generación de documentos gerenciales | 8 (Word) |
#
# ## Datos
#
# - **Reales:** precios máximos, CUM y medicamentos de la tabla pública de la **Circular 022 de 2026** de la Comisión Nacional de Precios de Medicamentos y Dispositivos Médicos. Casos de vigilancia del mercado publicados por prensa y entidades, **anonimizados** (sin nombres de empresas).
# - **Ficticios:** la empresa, los compradores, las tarifas, las cantidades, los puntos de dispensación y las PQRS.
# - **Sin datos aleatorios:** el notebook **no inventa nada**; solo lee los Excel que dejes en Drive. Si cambias un valor en un Excel y vuelves a ejecutar, cambian los resultados.

# %% [markdown]
# ## Paso 0. Preparar el entorno
#
# **En Colab** conecta tu Google Drive y trabaja sobre la carpeta `TECNOLOGIAS EN SALUD` (subcarpetas `entrada`, `referencia` y `salida`). Todo lo que genera el notebook queda en esa carpeta. Google te pedirá autorizar el acceso a Drive.
#
# Archivos que debes tener en Drive:
#
# | Carpeta | Archivo | Contenido |
# |---|---|---|
# | `entrada` | `Oferta_Red_Altos_del_Condor_nov2026.xlsx` | Oferta del comprador grande |
# | `entrada` | `LP-045-2026 propuesta economica Guayacan.xlsx` | Oferta del comprador mediano |
# | `entrada` | `cotizacion nogal (2).xlsx` | Oferta del comprador pequeño |
# | `entrada` | `base_pqrs_2026.xlsx` | Registro de PQRS del semestre |
# | `referencia` | Tabla de la Circular 022 (`.xlsb`, `.xlsx` o `.csv`) | Precios máximos oficiales |
# | `referencia` | `catalogo_empresa.xlsx` | Portafolio y precio piso (costo) de la empresa |
# | `referencia` | `sismed_precios_institucionales.xlsx` | Precios de venta publicados en el SISMED para medicamentos no regulados |
#
# Los nombres de archivo y de columna pueden cambiar: el notebook reconoce cada archivo por su contenido.

# %%
import os, sys, subprocess, glob

USUARIO_GITHUB = "ivanherrerabaquero-dev"
REPO = "validacion-precios-medicamentos"
CARPETA_DRIVE = "/content/drive/MyDrive/TECNOLOGIAS EN SALUD"   # cámbiala si tu carpeta tiene otro nombre
EN_COLAB = "google.colab" in sys.modules

if EN_COLAB:
    from google.colab import drive
    drive.mount("/content/drive")
    subprocess.run([sys.executable, "-m", "pip", "-q", "install", "python-docx", "openpyxl", "pyxlsb", "xlrd"])
    os.makedirs("/content/proyecto/data", exist_ok=True)
    os.chdir("/content/proyecto")
    enlaces = {"data/entrada": "entrada", "data/referencia": "referencia", "data/salida": "salida",
               "docs": "docs", "reportes": "reportes"}
    for local, en_drive in enlaces.items():
        destino = os.path.join(CARPETA_DRIVE, en_drive)
        os.makedirs(destino, exist_ok=True)
        if not os.path.islink(local):
            os.symlink(destino, local)

for carpeta in ["data/entrada", "data/referencia", "data/salida", "docs/data", "docs/descargas",
                "docs/img", "reportes/graficos"]:
    os.makedirs(carpeta, exist_ok=True)

print("Carpeta de trabajo:", CARPETA_DRIVE if EN_COLAB else os.getcwd())
for c in ["entrada", "referencia", "salida"]:
    print(f"  {c}:", os.listdir(f"data/{c}"))

# %% [markdown]
# ## Paso 1. Funciones de lectura inteligente
#
# Estas funciones permiten que el notebook **no dependa de nombres exactos**:
#
# - **Columnas:** cada columna estándar tiene una lista de sinónimos. El encabezado se normaliza (sin tildes, mayúsculas, sin signos) y se compara por coincidencia exacta, por palabras contenidas y por similitud de texto. Así, `Código CUM`, `cum/ium` o `CUM` se reconocen como **CUM**, y `Valor unitario ofertado` o `Precio` como **TARIFA**.
# - **Encabezado:** se busca en las primeras 40 filas de **cada hoja** la fila que más columnas reconocidas tiene.
# - **Archivos:** se lee cualquier `.xlsx`, `.xlsm`, `.xls`, `.xlsb` o `.csv`, se llame como se llame.

# %%
import re, unicodedata, difflib
import numpy as np
import pandas as pd

def normalizar(t):
    t = unicodedata.normalize("NFKD", str(t)).encode("ascii", "ignore").decode().upper()
    t = re.sub(r"[^A-Z0-9 ]", " ", t)
    return re.sub(r"\s+", " ", t).strip()

def puntaje(columna, sinonimo):
    a, b = normalizar(columna), normalizar(sinonimo)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    ta, tb = set(a.split()), set(b.split())
    if tb <= ta:                                   # todas las palabras del sinónimo están en la columna
        return 0.90 + 0.05 * len(tb) / len(ta)
    return difflib.SequenceMatcher(None, a, b).ratio() * 0.9

def mapear_columnas(columnas, sinonimos, umbral=0.78):
    """Devuelve {columna_estandar: columna_original} asignando cada original una sola vez."""
    candidatos = []
    for est, lista in sinonimos.items():
        for col in columnas:
            if pd.isna(col) or str(col).strip() == "":
                continue
            s = max(puntaje(col, x) for x in lista)
            if s >= umbral:
                candidatos.append((s, est, col))
    mapa, usadas = {}, set()
    for s, est, col in sorted(candidatos, key=lambda x: -x[0]):
        if est not in mapa and col not in usadas:
            mapa[est] = col; usadas.add(col)
    return mapa

EXTENSIONES = (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv")

def leer_hojas(ruta, nrows=None):
    """Devuelve {nombre_hoja: DataFrame sin encabezado} para cualquier formato soportado."""
    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".csv":
        for sep in [",", ";", "\t", "|"]:
            df = pd.read_csv(ruta, header=None, sep=sep, nrows=nrows, dtype=str, encoding_errors="ignore")
            if df.shape[1] > 2:
                return {"csv": df}
        return {"csv": df}
    motor = {".xlsb": "pyxlsb", ".xls": "xlrd"}.get(ext)
    return pd.read_excel(ruta, sheet_name=None, header=None, nrows=nrows, engine=motor)

def encontrar_tabla(ruta, sinonimos, obligatorias, filas_busqueda=40):
    """Busca en todas las hojas la fila de encabezado con más columnas reconocidas."""
    mejor = None
    for hoja, crudo in leer_hojas(ruta, nrows=filas_busqueda).items():
        for i in range(min(filas_busqueda, len(crudo))):
            mapa = mapear_columnas(list(crudo.iloc[i].values), sinonimos)
            if all(o in mapa for o in obligatorias):
                if mejor is None or len(mapa) > len(mejor["mapa"]):
                    mejor = {"hoja": hoja, "fila": i, "mapa": mapa}
    return mejor

print("Prueba:", mapear_columnas(["Código CUM", "Valor unitario ofertado", "Cant. mensual", "IVA"],
      {"CUM": ["CUM"], "TARIFA": ["TARIFA", "VALOR UNITARIO", "PRECIO"],
       "CANTIDAD": ["CANTIDAD", "CANT MENSUAL"], "IVA": ["APLICA IVA", "IVA"]}))

# %% [markdown]
# ## Paso 1B. Tabla oficial de precios máximos (Circular 022 de 2026)
#
# El notebook busca en la carpeta `referencia` **cualquier archivo** que tenga columnas de CUM y de precio máximo institucional, sin importar su nombre, y lo convierte a un CSV estándar. Si ya existe el CSV, lo usa directamente. Como las licitaciones y convenios son ventas a compradores institucionales, la validación usa el **PMV institucional**.

# %%
FECHA_CORTE = pd.Timestamp("2026-11-20")
EMPRESA = "CRM TECNOLOGÍAS EN SALUD"
RUTA_PMV = "data/referencia/pmv_circular_22_2026.csv"

SINONIMOS_PMV = {
    "ID_MR": ["ID MR", "ID MERCADO RELEVANTE", "CODIGO MERCADO RELEVANTE"],
    "MERCADO_RELEVANTE": ["MERCADO RELEVANTE"],
    "CUM": ["CUM IUM", "CUM", "IUM", "EXPEDIENTE CONSECUTIVO", "CODIGO CUM"],
    "MEDICAMENTO": ["MEDICAMENTO", "NOMBRE DEL MEDICAMENTO", "PRODUCTO"],
    "CANTIDAD_UNIDAD_MEDIDA": ["CANTIDAD POR UNIDAD DE MEDIDA", "CANTIDAD"],
    "UNIDAD_MEDIDA": ["UNIDAD DE MEDIDA"],
    "PMV_INSTITUCIONAL": ["PRECIO MAXIMO DE VENTA TRANSACCION PRIMARIA SECUNDARIA Y FINAL INSTITUCIONAL",
                          "PMV INSTITUCIONAL", "PRECIO MAXIMO INSTITUCIONAL"],
    "MARGEN_IPS": ["MARGEN PARA IPS", "MARGEN IPS"],
    "PMV_COMERCIAL_MAYORISTA": ["PRECIO MAXIMO DE VENTA TRANSACCION PRIMARIA Y SECUNDARIA COMERCIAL",
                                "PMV COMERCIAL MAYORISTA"],
    "PMV_COMERCIAL_FINAL": ["PRECIO MAXIMO DE VENTA TRANSACCION FINAL COMERCIAL", "PMV COMERCIAL FINAL"],
    "CIRCULAR": ["CIRCULAR CNPMDM", "CIRCULAR"],
    "FECHA_VIGENCIA": ["FECHA DE INICIO VIGENCIA PRECIO MAXIMO DE VENTA", "FECHA DE VIGENCIA"],
}

def convertir_tabla_oficial(ruta, ubic):
    df = pd.read_excel(ruta, sheet_name=ubic["hoja"], header=ubic["fila"],
                       engine={".xlsb": "pyxlsb", ".xls": "xlrd"}.get(os.path.splitext(ruta)[1].lower())) \
         if not ruta.lower().endswith(".csv") else pd.read_csv(ruta, header=ubic["fila"])
    df = df[[ubic["mapa"][k] for k in ubic["mapa"]]].copy()
    df.columns = list(ubic["mapa"])
    for k in SINONIMOS_PMV:
        if k not in df: df[k] = np.nan
    if pd.api.types.is_numeric_dtype(df["FECHA_VIGENCIA"]):
        df["FECHA_VIGENCIA"] = pd.to_datetime(df["FECHA_VIGENCIA"], unit="D", origin="1899-12-30").dt.date
    for c in ["MERCADO_RELEVANTE", "MEDICAMENTO", "CUM", "ID_MR"]:
        df[c] = df[c].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    df[list(SINONIMOS_PMV)].to_csv(RUTA_PMV, index=False)
    print(f"Tabla oficial convertida desde '{os.path.basename(ruta)}' (hoja {ubic['hoja']}, fila {ubic['fila'] + 1}): {len(df):,} filas")
    print("Columnas reconocidas:", {k: str(v)[:40] for k, v in ubic["mapa"].items()})

if not os.path.exists(RUTA_PMV):
    candidatos = [f for f in glob.glob("data/referencia/*") if f.lower().endswith(EXTENSIONES)
                  and not os.path.basename(f).startswith(("~$", "catalogo_empresa"))]
    for f in candidatos:
        ubic = encontrar_tabla(f, SINONIMOS_PMV, ["CUM", "PMV_INSTITUCIONAL"])
        if ubic:
            convertir_tabla_oficial(f, ubic); break
    else:
        raise FileNotFoundError("No encontré en 'referencia' un archivo con columnas de CUM y precio máximo institucional.")

pmv = pd.read_csv(RUTA_PMV, dtype={"CUM": str, "ID_MR": str})
pmv["PMV_INSTITUCIONAL"] = pd.to_numeric(pmv["PMV_INSTITUCIONAL"], errors="coerce")
pmv["CANTIDAD_UNIDAD_MEDIDA"] = pd.to_numeric(pmv["CANTIDAD_UNIDAD_MEDIDA"], errors="coerce")
print(f"Presentaciones con precio máximo: {len(pmv):,} · Mercados relevantes: {pmv['ID_MR'].nunique():,}")
display(pmv.head(3))

# %% [markdown]
# ## Paso 1C. Catálogo de la empresa (precio piso)
#
# Lee `catalogo_empresa.xlsx` (o cualquier archivo de `referencia` con columnas de CUM y precio piso). El **precio piso** es el costo más el margen mínimo de **CRM TECNOLOGÍAS EN SALUD**: si una tarifa queda por debajo, la empresa **pierde dinero**.

# %%
SINONIMOS_CATALOGO = {
    "CUM": ["CUM", "CODIGO CUM", "CUM IUM"],
    "PRECIO_PISO": ["PRECIO PISO", "COSTO", "COSTO UNITARIO", "PISO"],
    "NOMBRE_CORTO": ["NOMBRE CORTO", "PRINCIPIO ACTIVO", "NOMBRE GENERICO"],
}
cat_archivo = None
for f in sorted(glob.glob("data/referencia/*")):
    if f.lower().endswith(EXTENSIONES) and not os.path.basename(f).startswith("~$") and f != RUTA_PMV:
        ubic = encontrar_tabla(f, SINONIMOS_CATALOGO, ["CUM", "PRECIO_PISO"], filas_busqueda=10)
        if ubic: cat_archivo = (f, ubic); break
if cat_archivo is None:
    raise FileNotFoundError("Sube a 'referencia' el catálogo con columnas de CUM y precio piso (catalogo_empresa.xlsx).")
f, ubic = cat_archivo
crudo = leer_hojas(f)[ubic["hoja"]]
catalogo = crudo.iloc[ubic["fila"] + 1:].dropna(how="all")
catalogo.columns = [str(c) for c in crudo.iloc[ubic["fila"]].values]
catalogo = catalogo[[str(ubic["mapa"][k]) for k in ubic["mapa"]]]; catalogo.columns = list(ubic["mapa"])
catalogo["CUM"] = catalogo["CUM"].astype(str).str.strip().str.upper()
catalogo["PRECIO_PISO"] = pd.to_numeric(catalogo["PRECIO_PISO"], errors="coerce")
catalogo = catalogo.merge(pmv[["CUM", "MEDICAMENTO", "PMV_INSTITUCIONAL"]].drop_duplicates("CUM"), on="CUM", how="left")
if "NOMBRE_CORTO" not in catalogo:
    catalogo["NOMBRE_CORTO"] = catalogo["MEDICAMENTO"].str.split(" - ").str[1].fillna(catalogo["MEDICAMENTO"]).str.title()
print(f"Catálogo leído de '{os.path.basename(f)}': {len(catalogo)} productos")
display(catalogo[["CUM", "NOMBRE_CORTO", "PMV_INSTITUCIONAL", "PRECIO_PISO"]].head())

# %% [markdown]
# ## Paso 1D. Precios del SISMED (medicamentos no regulados)
#
# Los medicamentos que **no están en la Circular 022** no tienen precio máximo, pero sus precios de venta se reportan al **SISMED** (Sistema de Información de Precios de Medicamentos). Ofertar muy por encima del valor publicado puede leerse como **especulación** y generar reporte ante la SIC.
#
# Reglas de la empresa (puedes cambiarlas aquí):
# - **Crédito a 90 días** (hay que cubrir nómina y logística mientras se cobra): margen viable sobre el SISMED entre **22 % y 32 %**. Menos de 22 %: margen insuficiente. Más de 32 %: riesgo de especulación.
# - **Contado:** ajustarse al SISMED, con una tolerancia de ±5 %.
#
# > **Precio de referencia:** el archivo `sismed_precios_institucionales.xlsx` usa, por CUM, el precio de **2024** del Termómetro de Precios de Clicsalud (Ministerio de Salud, construido con reportes al SISMED; canal institucional, por presentación) y, cuando no existe, el de **2019** de la consulta pública del SISMED (`3he6-m866`, último periodo publicado). La columna `AÑO_REFERENCIA` indica cuál se usó. Si tienes una consulta más reciente, reemplaza el archivo: el notebook lo reconoce por sus columnas.

# %%
MARGEN_MIN_CREDITO = 22    # % mínimo sobre SISMED a crédito
MARGEN_MAX_CREDITO = 32    # % máximo sobre SISMED a crédito
TOLERANCIA_CONTADO = 5     # ± % sobre SISMED de contado

SINONIMOS_SISMED = {
    "CUM": ["CUM", "CODIGO CUM", "EXPEDIENTE CONSECUTIVO"],
    "PRECIO_SISMED": ["PRECIO SISMED REFERENCIA", "PRECIO SISMED PROMEDIO", "PRECIO SISMED", "VALOR PROMEDIO", "PRECIO PROMEDIO PONDERADO", "PRECIO PROMEDIO"],
    "ANO_REFERENCIA": ["ANO REFERENCIA", "ANIO REFERENCIA", "PERIODO REFERENCIA"],
    "PRINCIPIO_ACTIVO": ["PRINCIPIO ACTIVO", "DESCRIPCION ATC"],
    "DESCRIPCION_SISMED": ["DESCRIPCION COMERCIAL"],
    "ULTIMO_CORTE": ["ULTIMO CORTE", "FECHA CORTE"],
}
sismed = pd.DataFrame(columns=list(SINONIMOS_SISMED))
for f in sorted(glob.glob("data/referencia/*")):
    if not f.lower().endswith(EXTENSIONES) or os.path.basename(f).startswith("~$") or f == RUTA_PMV:
        continue
    ubic = encontrar_tabla(f, SINONIMOS_SISMED, ["CUM", "PRECIO_SISMED"], filas_busqueda=5)
    if ubic:
        crudo = pd.read_excel(f, sheet_name=ubic["hoja"], header=ubic["fila"]) if not f.lower().endswith(".csv") else pd.read_csv(f, header=ubic["fila"])
        sismed = crudo[[ubic["mapa"][k] for k in ubic["mapa"]]].copy(); sismed.columns = list(ubic["mapa"])
        for k in SINONIMOS_SISMED:
            if k not in sismed: sismed[k] = np.nan
        sismed["CUM"] = sismed["CUM"].astype(str).str.strip().str.upper()
        sismed["PRECIO_SISMED"] = pd.to_numeric(sismed["PRECIO_SISMED"], errors="coerce")
        sismed = sismed.dropna(subset=["PRECIO_SISMED"]).drop_duplicates("CUM")
        print(f"SISMED leído de '{os.path.basename(f)}': {len(sismed):,} CUM")
        if sismed["ANO_REFERENCIA"].notna().any(): print(sismed["ANO_REFERENCIA"].value_counts().to_string())
        break
else:
    print("No hay archivo del SISMED en 'referencia': los medicamentos no regulados quedarán como 'Sin referencia'.")
# Termómetro 2024 por expediente (todos los consecutivos): precio por unidad de dispensación
termo = pd.DataFrame(columns=["EXPEDIENTE", "MEDICAMENTO_2024", "UNIDAD_DISPENSACION", "PRECIO_UNIDAD_REFERENCIA", "CANAL_REFERENCIA"])
for f in sorted(glob.glob("data/referencia/*.xls*")):
    try:
        hojas_f = pd.ExcelFile(f).sheet_names
    except Exception:
        continue
    h = next((x for x in hojas_f if "TERMOMETRO" in normalizar(x)), None)
    if h:
        termo = pd.read_excel(f, sheet_name=h, dtype={"EXPEDIENTE": str})
        print(f"Termómetro 2024 por expediente: {len(termo):,} expedientes (hoja '{h}' de {os.path.basename(f)})")
        break
termo_idx = termo.drop_duplicates("EXPEDIENTE").set_index("EXPEDIENTE")

PALABRAS_UNIDAD = r"(TABLETAS?|CAPSULAS?|COMPRIMIDOS?|GRAGEAS?|SOBRES?|AMPOLLAS?|VIALES?|UNIDADES?|FRASCOS?|BOLSAS?|JERINGAS?|OVULOS?|PARCHES?|SUPOSITORIOS?|PLUMAS?|CARTUCHOS?)"
def unidades_en_texto(desc):
    # unidades de dispensación que trae la presentación, según la descripción ('CAJA X 30 TABLETAS' = 30)
    d = normalizar(desc)
    m = re.search(r"\b(?:X|POR|CON|DE)\s*(\d{1,4})\s*" + PALABRAS_UNIDAD, d)
    if m: return int(m.group(1))
    if re.search(r"\b(?:X|POR|CON)\s*(?:1|UN|UNA)\s+(FRASCO|VIAL|AMPOLLA|BOLSA|JERINGA|INHALADOR|TUBO|PLUMA|CARTUCHO)", d): return 1
    return None

def precio_por_expediente(cum, desc):
    # (precio por presentación o None, precio por unidad, unidad, medicamento) a partir del expediente del CUM
    e = str(cum).split("-")[0]
    if e not in termo_idx.index: return None, None, None, None
    r = termo_idx.loc[e]; pu = float(r["PRECIO_UNIDAD_REFERENCIA"])
    n = unidades_en_texto(desc)
    if n is None and str(r["UNIDAD_DISPENSACION"]) in ("Frasco", "Vial", "Ampolla", "Bolsa", "Inhalador", "Jeringa Prellenada", "Tubo") and re.search(r"AMPOLLA|VIAL|FRASCO|BOLSA|JERINGA|INHALADOR|TUBO", normalizar(desc)):
        n = 1
    return (round(pu * n, 2) if n else None), pu, r["UNIDAD_DISPENSACION"], r["MEDICAMENTO_2024"]
import re
display(sismed.head(3))

# %% [markdown]
# ## Paso 2. Inventario de la carpeta `entrada`
#
# Cada archivo se clasifica **por su contenido**: si tiene columnas de radicado, fecha y categoría es la **base de PQRS**; si tiene CUM y tarifa es una **oferta de comprador**.

# %%
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

SINONIMOS_PQRS = {
    "RADICADO": ["RADICADO", "NUMERO DE RADICADO", "ID PQRS", "CONSECUTIVO PQRS"],
    "FECHA_RADICACION": ["FECHA RADICACION", "FECHA DE RADICACION", "FECHA RECIBIDO", "FECHA DE RECEPCION"],
    "MEDIO": ["MEDIO", "MEDIO DE RECEPCION", "CANAL DE RECEPCION"],
    "CANAL": ["CANAL", "LINEA DE NEGOCIO"],
    "CIUDAD": ["CIUDAD", "MUNICIPIO"],
    "PUNTO": ["PUNTO", "PUNTO DE DISPENSACION", "SEDE"],
    "SOLICITANTE": ["SOLICITANTE", "TIPO DE SOLICITANTE", "TIPO DE USUARIO"],
    "ID_USUARIO": ["ID USUARIO", "ID PACIENTE", "CODIGO USUARIO"],
    "COMPRADOR": ["COMPRADOR", "CLIENTE", "ENTIDAD"],
    "CATEGORIA": ["CATEGORIA", "MOTIVO", "TIPOLOGIA", "CAUSA"],
    "MEDICAMENTO": ["MEDICAMENTO", "PRODUCTO"],
    "DESCRIPCION": ["DESCRIPCION", "DETALLE", "RESUMEN DEL CASO"],
    "SEVERIDAD": ["SEVERIDAD", "PRIORIDAD", "CRITICIDAD"],
    "ENTE_RELACIONADO": ["ENTE RELACIONADO", "ENTE DE CONTROL", "ENTE"],
    "NORMATIVA": ["NORMATIVA", "NORMA", "MARCO NORMATIVO"],
    "ESTADO": ["ESTADO", "ESTADO DEL CASO"],
    "FECHA_RESPUESTA": ["FECHA RESPUESTA", "FECHA DE RESPUESTA", "FECHA DE CIERRE"],
    "ESCALADO_ENTE": ["ESCALADO ENTE", "ESCALADO A ENTE", "ESCALADO"],
    "TUTELA": ["TUTELA", "ACCION DE TUTELA"],
    "IMPACTO_COP": ["IMPACTO COP", "IMPACTO ECONOMICO", "VALOR EN RIESGO"],
    "SATISFACCION": ["SATISFACCION", "CALIFICACION", "ENCUESTA"],
}
OBLIG_PQRS = ["RADICADO", "FECHA_RADICACION", "CATEGORIA"]

archivos_oferta, archivo_pqrs = [], None
for ruta in sorted(f for f in glob.glob("data/entrada/*") if f.lower().endswith(EXTENSIONES)
                   and not os.path.basename(f).startswith("~$")):
    ubic = encontrar_tabla(ruta, SINONIMOS_PQRS, OBLIG_PQRS)
    if ubic and archivo_pqrs is None:
        archivo_pqrs = (ruta, ubic); print("Base de PQRS:   ", os.path.basename(ruta))
    else:
        archivos_oferta.append(ruta); print("Oferta:         ", os.path.basename(ruta))
if archivo_pqrs is None:
    raise FileNotFoundError("No encontré en 'entrada' la base de PQRS (columnas de radicado, fecha y categoría).")

# %% [markdown]
# ## Paso 3. Leer y consolidar las bases (lectura inteligente)
#
# Para cada archivo de `entrada`: busca la hoja y la fila del encabezado, reconoce las columnas por sinónimos, toma el nombre del comprador del encabezado del archivo (o del nombre del archivo si no lo tiene) y deja un **reporte de lectura** con lo que encontró y lo que falta. Un archivo que no se pueda leer **no detiene el proceso**: queda reportado.

# %%
SINONIMOS_ANEXO = {
    "CONSECUTIVO": ["CONSECUTIVO", "ITEM", "NO", "NUMERO"],
    "ID MR": ["ID MR", "MERCADO ID", "ID MERCADO"],
    "MERCADO RELEVANTE": ["MERCADO RELEVANTE", "MERCADO"],
    "CUM": ["CUM", "CODIGO CUM", "CUM IUM", "IUM", "EXPEDIENTE CONSECUTIVO", "COD CUM"],
    "DESCRIPCION": ["DESCRIPCION", "MEDICAMENTO", "PRODUCTO", "NOMBRE DEL MEDICAMENTO", "DESCRIPCION DEL PRODUCTO"],
    "AGRUPADOR": ["AGRUPADOR", "GRUPO", "CATEGORIA"],
    "CANTIDAD ESTIMADA MES": ["CANTIDAD ESTIMADA MES", "CANTIDAD", "CANT MENSUAL", "CANTIDAD MENSUAL", "UNIDADES", "VOLUMEN"],
    "TARIFA PRESENTACION COMERCIAL": ["TARIFA PRESENTACION COMERCIAL", "TARIFA", "TARIFA UNITARIA", "VALOR UNITARIO",
                                      "VALOR UNITARIO OFERTADO", "PRECIO", "PRECIO UNITARIO", "VALOR OFERTADO"],
    "APLICA IVA (SI-NO)": ["APLICA IVA SI NO", "APLICA IVA", "IVA"],
    "TIPO PLAN": ["TIPO PLAN", "PLAN"],
    "OBSERVACIONES": ["OBSERVACIONES", "OBSERVACION", "NOTAS"],
}
OBLIGATORIAS = ["CUM", "TARIFA PRESENTACION COMERCIAL"]
CLAVES_COMPRADOR = ["NOMBRE DEL COMPRADOR", "COMPRADOR", "CLIENTE", "RAZON SOCIAL", "ENTIDAD"]
CLAVES_MODALIDAD = ["MODALIDAD", "TIPO DE PROCESO"]
CLAVES_TAMANO = ["TAMANO DEL COMPRADOR", "TAMANO", "SEGMENTO", "TIPO DE COMPRADOR"]
CLAVES_PAGO = ["FORMA DE PAGO", "CONDICION DE PAGO", "PLAZO DE PAGO"]

def valor_de_clave(crudo, fila_tope, claves):
    """Busca en las filas de encabezado el par 'clave: valor' que mejor coincide con las claves."""
    mejor, valor = 0.9, None
    for _, r in crudo.iloc[:fila_tope].iterrows():
        vals = [v for v in r.values if pd.notna(v) and str(v).strip()]
        if len(vals) >= 2:
            s_ = max(puntaje(vals[0], c) for c in claves)
            if s_ >= mejor:
                mejor, valor = s_ + 1e-9, str(vals[1]).strip()
    return valor

def leer_archivo(ruta):
    ubic = encontrar_tabla(ruta, SINONIMOS_ANEXO, OBLIGATORIAS)
    if ubic is None:
        raise ValueError("no se encontró una fila con columnas de CUM y tarifa")
    crudo = leer_hojas(ruta)[ubic["hoja"]]
    df = crudo.iloc[ubic["fila"] + 1:].dropna(how="all").copy()
    df.columns = [str(c) for c in crudo.iloc[ubic["fila"]].values]
    df = df[[str(ubic["mapa"][k]) for k in ubic["mapa"]]]
    df.columns = list(ubic["mapa"])
    for k in SINONIMOS_ANEXO:
        if k not in df: df[k] = np.nan
    comprador = (valor_de_clave(crudo, ubic["fila"], CLAVES_COMPRADOR) or
                 normalizar(os.path.splitext(os.path.basename(ruta))[0])).upper()
    df.insert(0, "COMPRADOR", comprador)
    df["TAMANO"] = (valor_de_clave(crudo, ubic["fila"], CLAVES_TAMANO) or "Sin clasificar").capitalize()
    pago = (valor_de_clave(crudo, ubic["fila"], CLAVES_PAGO) or "CRÉDITO 90 DÍAS").upper()
    df["FORMA_PAGO"] = "CONTADO" if "CONTADO" in normalizar(pago) else "CRÉDITO"
    df.insert(1, "MODALIDAD", (valor_de_clave(crudo, ubic["fila"], CLAVES_MODALIDAD) or "SIN DATO").upper())
    df.insert(2, "ARCHIVO", os.path.basename(ruta))
    df["FILA_ORIGEN"] = df.index + 1
    reporte = {"ARCHIVO": os.path.basename(ruta), "HOJA": ubic["hoja"], "FILA_ENCABEZADO": ubic["fila"] + 1,
               "COMPRADOR": comprador, "LINEAS": len(df), "ESTADO": "LEÍDO",
               "COLUMNAS_RECONOCIDAS": "; ".join(f"{v} → {k}" for k, v in ubic["mapa"].items()),
               "COLUMNAS_FALTANTES": ", ".join(k for k in SINONIMOS_ANEXO if k not in ubic["mapa"]) or "Ninguna"}
    return df[["COMPRADOR", "TAMANO", "MODALIDAD", "FORMA_PAGO", "ARCHIVO", "FILA_ORIGEN"] + list(SINONIMOS_ANEXO)], reporte

partes, reportes = [], []
for ruta in archivos_oferta:
    try:
        df, rep = leer_archivo(ruta); partes.append(df); reportes.append(rep)
    except Exception as e:
        reportes.append({"ARCHIVO": os.path.basename(ruta), "ESTADO": f"NO LEÍDO: {e}"})
lectura = pd.DataFrame(reportes)
if not partes:
    raise RuntimeError("Ningún archivo de 'entrada' tiene columnas de CUM y tarifa. Revisa el reporte de lectura.")
consolidado = pd.concat(partes, ignore_index=True)


def a_numero(v):
    if isinstance(v, (int, float, np.integer, np.floating)) and not pd.isna(v):
        return float(v)
    s = re.sub(r"[^\d,.-]", "", str(v))
    if s.count(".") > 1 or ("." in s and "," not in s and len(s.split(".")[-1]) == 3):
        s = s.replace(".", "")                      # 1.250.000 o 1.250 → miles
    s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return np.nan

consolidado["TARIFA PRESENTACION COMERCIAL"] = consolidado["TARIFA PRESENTACION COMERCIAL"].map(a_numero)
consolidado["CANTIDAD ESTIMADA MES"] = consolidado["CANTIDAD ESTIMADA MES"].map(a_numero)
consolidado["DESCRIPCION"] = consolidado["DESCRIPCION"].map(lambda x: str(x).upper().strip())
consolidado["CUM"] = consolidado["CUM"].astype(str).str.strip().str.upper().str.replace(" ", "")
consolidado["APLICA IVA (SI-NO)"] = consolidado["APLICA IVA (SI-NO)"].fillna("").astype(str).str.strip().str.upper()
display(lectura)

# %% [markdown]
# ## Paso 4. Validar las propuestas y cruzar con la Circular 022
#
# 1. **CUM** con formato válido (expediente-consecutivo, como `19911481-1`, o IUM de 15 caracteres) y **presente** en la Circular 022.
# 2. **Duplicados** dentro del mismo comprador.
# 3. **Tarifa, IVA y cantidad** correctamente diligenciados.
# 4. **Tarifa frente al PMV institucional** del CUM: si lo supera, la línea **incumple la Circular 022 de 2026**.
# 5. **Tarifa frente al precio piso interno:** riesgo de margen.
#
# **Cruce de cada línea** (hoja `CRUCE_CIRCULAR_022` del Excel):
#
# | Columna | Cálculo |
# |---|---|
# | Tarifa ofertada | Precio por presentación que envía el comprador |
# | PMV Circular 022 | Precio máximo institucional del CUM |
# | % ganancia sobre PMV | (Tarifa − PMV) / PMV, cuando la tarifa **supera** el PMV. Es un sobreprecio no permitido |
# | % margen | (Tarifa − precio piso) / tarifa: lo que gana la empresa en cada unidad |
# | % pérdida | (Precio piso − tarifa) / precio piso, cuando la tarifa queda **por debajo del costo** |
# | Precio SISMED y % sobre SISMED | Solo para medicamentos **no regulados**: (Tarifa − SISMED) / SISMED, comparado con el rango según la forma de pago |

# %%
v = (consolidado
     .merge(pmv[["CUM", "MERCADO_RELEVANTE", "CANTIDAD_UNIDAD_MEDIDA", "UNIDAD_MEDIDA", "PMV_INSTITUCIONAL"]]
            .drop_duplicates("CUM"), on="CUM", how="left")
     .merge(catalogo[["CUM", "PRECIO_PISO"]], on="CUM", how="left")
     .merge(sismed[["CUM", "PRECIO_SISMED", "ANO_REFERENCIA"]], on="CUM", how="left"))
T = "TARIFA PRESENTACION COMERCIAL"
# Si el CUM exacto no tiene PMV ni precio SISMED, se busca su expediente en el Termómetro 2024
sin_ref = v["PMV_INSTITUCIONAL"].isna() & v["PRECIO_SISMED"].isna()
fx = [precio_por_expediente(c, d) if s_ else (None, None, None, None) for c, d, s_ in zip(v["CUM"], v["DESCRIPCION"], sin_ref)]
v["PRECIO_2024_UNIDAD_EXPEDIENTE"] = [x[1] for x in fx]
usa = sin_ref & pd.Series([x[0] is not None for x in fx], index=v.index)
v["PRECIO_SISMED"] = v["PRECIO_SISMED"].astype(float).where(~usa, pd.Series([np.nan if x[0] is None else x[0] for x in fx], index=v.index, dtype=float))
v.loc[usa, "ANO_REFERENCIA"] = "2024 por expediente × unidades"
print(f"Precio SISMED completado por expediente: {int(usa.sum())} líneas · con precio solo por unidad: {int((sin_ref & ~usa & v['PRECIO_2024_UNIDAD_EXPEDIENTE'].notna()).sum())}")
v["H_CUM_FORMATO"] = ~(v["CUM"].str.fullmatch(r"\d{4,9}-\d{1,3}") | v["CUM"].str.fullmatch(r"[0-9A-Z]{15}"))
v["H_CUM_NO_EXISTE"] = ~v["H_CUM_FORMATO"] & v["PMV_INSTITUCIONAL"].isna() & v["PRECIO_SISMED"].isna()
v["H_DUPLICADO"] = v.duplicated(["COMPRADOR", "CUM", T], keep="first")
v["H_TARIFA_VACIA"] = v[T].isna() | (v[T] <= 0)
v["H_IVA_VACIO"] = ~v["APLICA IVA (SI-NO)"].isin(["SI", "NO"])
v["H_CANTIDAD_INVALIDA"] = v["CANTIDAD ESTIMADA MES"].isna() | (v["CANTIDAD ESTIMADA MES"] <= 0)
v["H_SOBRE_PMV"] = v[T] > v["PMV_INSTITUCIONAL"]
v["H_BAJO_PISO"] = v[T] < v["PRECIO_PISO"]
cant = v["CANTIDAD ESTIMADA MES"].clip(lower=0).fillna(0)
v["PCT_SOBRE_PMV"] = np.where(v["H_SOBRE_PMV"], (v[T] / v["PMV_INSTITUCIONAL"] - 1) * 100, 0).round(1)
v["EXCEDENTE_MES"] = np.where(v["H_SOBRE_PMV"], (v[T] - v["PMV_INSTITUCIONAL"]) * cant, 0).round(0)
v["VALOR_OFERTA_MES"] = (v[T] * cant).round(0)
v["DIF_VS_PMV_PCT"] = ((v[T] / v["PMV_INSTITUCIONAL"] - 1) * 100).round(1)
v["PCT_GANANCIA_SOBRE_PMV"] = v["DIF_VS_PMV_PCT"].clip(lower=0)
v["PCT_MARGEN"] = ((v[T] - v["PRECIO_PISO"]) / v[T] * 100).round(1)
v["PCT_PERDIDA"] = np.where(v["H_BAJO_PISO"], (v["PRECIO_PISO"] - v[T]) / v["PRECIO_PISO"] * 100, 0).round(1)
v.loc[v["PRECIO_PISO"].isna(), "PCT_PERDIDA"] = np.nan
v["PERDIDA_MES"] = np.where(v["H_BAJO_PISO"], (v["PRECIO_PISO"] - v[T]) * cant, 0).round(0)
# Medicamentos no regulados: margen sobre el precio publicado en el SISMED
no_reg = v["PMV_INSTITUCIONAL"].isna() & v["PRECIO_SISMED"].notna() & ~v["H_CUM_FORMATO"]
v["PCT_SOBRE_SISMED"] = np.where(no_reg, (v[T] / v["PRECIO_SISMED"] - 1) * 100, np.nan).round(1)
contado = v["FORMA_PAGO"] == "CONTADO"
tope = np.where(contado, TOLERANCIA_CONTADO, MARGEN_MAX_CREDITO)
v["H_ESPECULACION"] = no_reg & (v["PCT_SOBRE_SISMED"] > tope)
v["H_MARGEN_SISMED"] = no_reg & np.where(contado, v["PCT_SOBRE_SISMED"] < -TOLERANCIA_CONTADO, v["PCT_SOBRE_SISMED"] < MARGEN_MIN_CREDITO)
v["EXCEDENTE_MES"] = np.where(v["H_ESPECULACION"], (v[T] - v["PRECIO_SISMED"] * (1 + tope / 100)) * cant, v["EXCEDENTE_MES"]).round(0)
v["RESULTADO_CRUCE"] = np.select(
    [v["H_CUM_FORMATO"], v["H_SOBRE_PMV"], v["H_ESPECULACION"], v["H_BAJO_PISO"],
     no_reg & v["H_MARGEN_SISMED"] & contado, no_reg & v["H_MARGEN_SISMED"],
     no_reg & contado, no_reg, v["PMV_INSTITUCIONAL"].isna()],
    ["CUM con formato inválido", "Sobreprecio: supera el PMV", "Riesgo de especulación (reporte SIC)", "Pérdida: tarifa bajo el costo",
     "Bajo SISMED: revisar margen", "Margen insuficiente para crédito",
     "Alineado a SISMED (contado)", "Margen viable sobre SISMED (crédito)", "Sin referencia: consultar SISMED"], "Dentro del rango")

ETIQUETAS = {"H_SOBRE_PMV": "Supera PMV (Circular 022)", "H_BAJO_PISO": "Bajo precio piso (riesgo margen)",
             "H_CUM_FORMATO": "CUM con formato inválido", "H_CUM_NO_EXISTE": "CUM sin referencia (ni Circular 022 ni SISMED)",
             "H_ESPECULACION": "Riesgo de especulación sobre SISMED", "H_MARGEN_SISMED": "Margen insuficiente sobre SISMED",
             "H_DUPLICADO": "Línea duplicada", "H_TARIFA_VACIA": "Tarifa vacía o cero",
             "H_IVA_VACIO": "IVA sin diligenciar", "H_CANTIDAD_INVALIDA": "Cantidad negativa o vacía"}
cols_h = list(ETIQUETAS)
v["N_HALLAZGOS"] = v[cols_h].sum(axis=1)
v["HALLAZGOS"] = v[cols_h].apply(lambda r: "; ".join(ETIQUETAS[c] for c in cols_h if r[c]), axis=1)
v["ESTADO"] = np.select([v["H_SOBRE_PMV"] | v["H_ESPECULACION"] | v["H_CUM_NO_EXISTE"] | v["H_CUM_FORMATO"] | v["H_TARIFA_VACIA"],
                         v["N_HALLAZGOS"] > 0], ["RECHAZADA", "CON OBSERVACIONES"], "APROBADA")

resumen = v.groupby(["COMPRADOR", "TAMANO"]).agg(
    MODALIDAD=("MODALIDAD", "first"), LINEAS=("CUM", "size"),
    APROBADAS=("ESTADO", lambda s: (s == "APROBADA").sum()),
    CON_OBSERVACIONES=("ESTADO", lambda s: (s == "CON OBSERVACIONES").sum()),
    RECHAZADAS=("ESTADO", lambda s: (s == "RECHAZADA").sum()),
    FORMA_PAGO=("FORMA_PAGO", "first"), SOBRE_PMV=("H_SOBRE_PMV", "sum"), ESPECULACION=("H_ESPECULACION", "sum"), BAJO_PISO=("H_BAJO_PISO", "sum"),
    EXCEDENTE_MES=("EXCEDENTE_MES", "sum"), VALOR_OFERTA_MES=("VALOR_OFERTA_MES", "sum")).reset_index()
resumen["PCT_APROBACION"] = (resumen["APROBADAS"] / resumen["LINEAS"] * 100).round(1)
hallazgos_tipo = v[cols_h].groupby(v["COMPRADOR"]).sum().rename(columns=ETIQUETAS)
resumen["PERDIDA_MES"] = resumen["COMPRADOR"].map(v.groupby("COMPRADOR")["PERDIDA_MES"].sum())

cruce = v[["COMPRADOR", "TAMANO", "FORMA_PAGO", "CUM", "DESCRIPCION", "CANTIDAD ESTIMADA MES", T, "PMV_INSTITUCIONAL", "PRECIO_SISMED", "PRECIO_PISO",
           "DIF_VS_PMV_PCT", "PCT_GANANCIA_SOBRE_PMV", "PCT_SOBRE_SISMED", "PCT_MARGEN", "PCT_PERDIDA", "EXCEDENTE_MES", "PERDIDA_MES",
           "RESULTADO_CRUCE", "ANO_REFERENCIA"]].rename(columns={"FORMA_PAGO": "FORMA DE PAGO", "PRECIO_SISMED": "PRECIO SISMED", "ANO_REFERENCIA": "AÑO PRECIO SISMED", "PCT_SOBRE_SISMED": "% SOBRE SISMED",
    "TAMANO": "TAMAÑO", "DESCRIPCION": "MEDICAMENTO", "CANTIDAD ESTIMADA MES": "CANTIDAD MES", T: "TARIFA OFERTADA",
    "PMV_INSTITUCIONAL": "PMV CIRCULAR 022", "PRECIO_PISO": "PRECIO PISO (COSTO)", "DIF_VS_PMV_PCT": "% DIFERENCIA VS PMV",
    "PCT_GANANCIA_SOBRE_PMV": "% GANANCIA SOBRE PMV", "PCT_MARGEN": "% MARGEN", "PCT_PERDIDA": "% PÉRDIDA",
    "EXCEDENTE_MES": "EXCEDENTE SOBRE TOPE / MES", "PERDIDA_MES": "PÉRDIDA / MES", "RESULTADO_CRUCE": "RESULTADO"})
print(cruce["RESULTADO"].value_counts().to_string())
display(cruce.head(10))

# %% [markdown]
# ## Paso 5. Seguimiento de PQRS
#
# Se lee la base de PQRS que está en `entrada` (440 casos de junio a noviembre de 2026, en dos canales: *B2B – convenios y licitaciones* y *Dispensación – pacientes*). El notebook **calcula** a partir de las fechas y la severidad:
# - **Términos:** el legal general es de 15 días hábiles (Ley 1755 de 2015). Los pendientes deben entregarse en 48 horas (Resolución 1604 de 2013). Internamente hay un acuerdo de servicio por severidad: crítica 2, alta 5, media 10 y baja 15 días hábiles.
# - **Campos calculados:** días hábiles, oportunidad legal, cumplimiento del acuerdo de servicio, casos vencidos, envejecimiento y reincidencia.

# %%
FESTIVOS_2026 = pd.to_datetime(["2026-01-01", "2026-01-12", "2026-03-23", "2026-04-02", "2026-04-03",
    "2026-05-01", "2026-05-18", "2026-06-08", "2026-06-15", "2026-06-29", "2026-07-20", "2026-08-07",
    "2026-08-17", "2026-10-12", "2026-11-02", "2026-11-16", "2026-12-08", "2026-12-25"]).values.astype("datetime64[D]")
def dias_habiles(ini, fin):
    return np.busday_count(pd.to_datetime(ini).values.astype("datetime64[D]"),
                           pd.to_datetime(fin).values.astype("datetime64[D]"), holidays=FESTIVOS_2026)
def sumar_habiles(f, d):
    return pd.Timestamp(np.busday_offset(np.datetime64(f.date()), int(d), roll="forward", holidays=FESTIVOS_2026))

DIAS_SLA = {"Crítica": 2, "Alta": 5, "Media": 10, "Baja": 15}

ruta, ubic = archivo_pqrs
crudo = leer_hojas(ruta)[ubic["hoja"]]
pqr = crudo.iloc[ubic["fila"] + 1:].dropna(how="all").copy()
pqr.columns = [str(c) for c in crudo.iloc[ubic["fila"]].values]
pqr = pqr[[str(ubic["mapa"][k]) for k in ubic["mapa"]]]; pqr.columns = list(ubic["mapa"])
for k in SINONIMOS_PQRS:
    if k not in pqr: pqr[k] = np.nan
pqr = pqr[list(SINONIMOS_PQRS)].reset_index(drop=True)
for c in ["FECHA_RADICACION", "FECHA_RESPUESTA"]:
    pqr[c] = pd.to_datetime(pqr[c], errors="coerce")
for c in ["ESCALADO_ENTE", "TUTELA"]:
    pqr[c] = pqr[c].astype(str).map(normalizar).isin(["SI", "S", "TRUE", "1", "X"])
pqr["IMPACTO_COP"] = pqr["IMPACTO_COP"].map(a_numero).fillna(0)
pqr["SATISFACCION"] = pd.to_numeric(pqr["SATISFACCION"], errors="coerce")
pqr["SEVERIDAD"] = pqr["SEVERIDAD"].astype(str).str.strip().str.capitalize().replace({"Critica": "Crítica"})
pqr["COMPRADOR"] = pqr["COMPRADOR"].fillna("—").astype(str).str.strip().str.upper().replace({"—": "—", "NAN": "—"})
for c in ["MEDIO", "CANAL", "CIUDAD", "PUNTO", "SOLICITANTE", "ID_USUARIO", "MEDICAMENTO", "ENTE_RELACIONADO", "NORMATIVA", "ESTADO"]:
    pqr[c] = pqr[c].fillna("—").astype(str).str.strip()
pqr = pqr.dropna(subset=["FECHA_RADICACION"]).sort_values("FECHA_RADICACION").reset_index(drop=True)
pqr["DIAS_SLA"] = pqr["SEVERIDAD"].map(DIAS_SLA).fillna(15).astype(int)

abierta = pqr["FECHA_RESPUESTA"].isna()
pqr["DIAS_HABILES"] = dias_habiles(pqr["FECHA_RADICACION"], pqr["FECHA_RESPUESTA"].fillna(FECHA_CORTE))
pqr["OPORTUNA_LEGAL"] = np.where(abierta, np.nan, (pqr["DIAS_HABILES"] <= 15).astype(float))
pqr["CUMPLE_SLA"] = np.where(abierta, np.nan, (pqr["DIAS_HABILES"] <= pqr["DIAS_SLA"]).astype(float))
pqr["VENCIDA_ABIERTA"] = abierta & (pqr["DIAS_HABILES"] > 15)
pqr["ENVEJECIMIENTO"] = pd.cut(pqr["DIAS_HABILES"], [-1, 5, 15, 30, 999], labels=["0-5 días", "6-15 días", "16-30 días", "> 30 días"]).astype(str)
# Reincidencia: mismo paciente (o comprador y medicamento) con la misma categoría en los 30 días anteriores
pqr["CLAVE"] = np.where(pqr["CANAL"] == "B2B", pqr["COMPRADOR"] + "|" + pqr["MEDICAMENTO"], pqr["ID_USUARIO"]) + "|" + pqr["CATEGORIA"]
pqr["REINCIDENTE"] = pqr.groupby("CLAVE")["FECHA_RADICACION"].diff().dt.days.le(30).fillna(False)
pqr["MES"] = pqr["FECHA_RADICACION"].dt.to_period("M").astype(str)
pqr = pqr.drop(columns="CLAVE")
print(f"PQRS leídas de '{os.path.basename(ruta)}': {len(pqr)}")

cerradas = pqr[~abierta]
kpi_pqr = {
    "total": int(len(pqr)), "abiertas": int(abierta.sum()), "vencidas_abiertas": int(pqr["VENCIDA_ABIERTA"].sum()),
    "pct_oportunidad": round(float(cerradas["OPORTUNA_LEGAL"].mean() * 100), 1),
    "pct_sla": round(float(cerradas["CUMPLE_SLA"].mean() * 100), 1),
    "promedio_dias": round(float(cerradas["DIAS_HABILES"].mean()), 1),
    "pct_internos": round(float(pqr["SOLICITANTE"].str.startswith("Usuario interno").mean() * 100), 1),
    "tutelas": int(pqr["TUTELA"].sum()), "escaladas": int(pqr["ESCALADO_ENTE"].sum()),
    "pct_reincidencia": round(float(pqr["REINCIDENTE"].mean() * 100), 1),
    "satisfaccion": round(float(cerradas["SATISFACCION"].mean()), 2),
    "impacto_cop": float(pqr["IMPACTO_COP"].sum()),
}
print(kpi_pqr)
display(pqr.groupby(["MES", "CANAL"]).size().unstack(fill_value=0))

# %% [markdown]
# ## Paso 5B. Metas regulatorias y semáforo
#
# Indicadores que vigilan los entes de control, comparados con el resultado del periodo. Verde: cumple; amarillo: a menos de 10 puntos; rojo: incumple.

# %%
pend = cerradas[cerradas["CATEGORIA"] == "Entrega incompleta (pendientes)"]
cal = cerradas[cerradas["CATEGORIA"].isin(["Calidad y cadena de frío", "Error de dispensación", "Trazabilidad y documentación"])]
metas = pd.DataFrame([
    ["SIC", "Líneas ofertadas dentro del PMV institucional", "Circular 022 de 2026 (CNPMDM)", 100.0,
     round((1 - v["H_SOBRE_PMV"].mean()) * 100, 1), "mayor"],
    ["Supersalud", "Pendientes entregados en 48 horas (2 días hábiles)", "Resolución 1604 de 2013", 100.0,
     round(float((pend["DIAS_HABILES"] <= 2).mean() * 100), 1), "mayor"],
    ["Supersalud / Ley", "PQRS respondidas dentro de 15 días hábiles", "Ley 1755 de 2015", 100.0,
     kpi_pqr["pct_oportunidad"], "mayor"],
    ["INVIMA", "Casos de calidad y trazabilidad cerrados dentro del acuerdo de servicio", "Resolución 1403 de 2007", 95.0,
     round(float(cal["CUMPLE_SLA"].mean() * 100), 1), "mayor"],
    ["Interna", "Cumplimiento del acuerdo de servicio por severidad", "Política de servicio CRM", 90.0,
     kpi_pqr["pct_sla"], "mayor"],
    ["Interna", "Reincidencia de PQRS (misma causa en 30 días)", "Política de servicio CRM", 10.0,
     kpi_pqr["pct_reincidencia"], "menor"],
], columns=["ENTE", "INDICADOR", "NORMATIVA", "META", "RESULTADO", "SENTIDO"])
def semaforo(r):
    brecha = (r.RESULTADO - r.META) if r.SENTIDO == "mayor" else (r.META - r.RESULTADO)
    return "Verde" if brecha >= 0 else ("Amarillo" if brecha >= -10 else "Rojo")
metas["SEMAFORO"] = metas.apply(semaforo, axis=1)
display(metas)

# %% [markdown]
# ## Paso 5C. Casos reales del mercado (anonimizados)
#
# Casos públicos de vigilancia a distribuidores y gestores farmacéuticos en Colombia, tomados de prensa especializada y de entidades oficiales. **Se omiten los nombres de las empresas**; se conservan fecha, ente, cifras y medida, y se relacionan con la categoría equivalente de la base de PQRS para usarlos como referencia de riesgo.

# %%
casos_mercado = pd.DataFrame([
    ["2026-09-22", "SIC", "Distribuidor mayorista (Cúcuta)", "Ajuste de precios Circular 022",
     "Venta de 47 medicamentos por encima del precio máximo regulado; un producto con exceso de hasta 7.001 %.",
     "Multa de $4.614 millones confirmada en segunda instancia (Circular 13 de 2022).",
     "Validar cada línea contra el PMV antes de facturar.", "Consultorsalud",
     "https://consultorsalud.com/sic-multa-precios-de-medicamentos/"],
    ["2026-04-20", "SIC", "Laboratorio y distribuidor del canal institucional (Armenia)", "Facturación superior al pactado",
     "Al menos 26 medicamentos vendidos con sobrecostos de hasta 400 % en el canal institucional.",
     "Multas por más de $1.379 millones en primera instancia.",
     "Alertar tarifas que superen el PMV en ofertas de convenios.", "Consultorsalud",
     "https://consultorsalud.com/sic-bayer-comfenalco-sobrecostos-medicamentos/"],
    ["2025-08-29", "SIC", "Tres mayoristas e IPS (Cali y Cúcuta)", "Ajuste de precios Circular 022",
     "Excesos entre 0,09 % y 7.000 % sobre el precio máximo en ventas de 2022 y 2023.",
     "Sanciones en primera instancia por exceder precios máximos (Circular 13 de 2022).",
     "Revisar el precio por canal: institucional vs. comercial.", "El Espectador",
     "https://www.elespectador.com/salud/sancionan-a-tres-mayoristas-e-ips-por-exceder-precios-maximos-de-venta-de-medicamentos/"],
    ["2025-03-25", "Supersalud", "Gestor farmacéutico nacional (Bogotá / Funza)", "Retención / desabastecimiento",
     "Auditoría encontró 6.116 unidades pendientes; 7 de 22 medicamentos priorizados tenían existencias en bodega.",
     "Orden de entregar los pendientes en 24 horas.",
     "Cruzar pendientes del punto con el inventario del centro de distribución.", "Consultorsalud",
     "https://consultorsalud.com/?p=311264"],
    ["2025-07-15", "Supersalud", "Gestor farmacéutico de caja de compensación (Bogotá / Funza)", "Entrega incompleta (pendientes)",
     "11.855 unidades figuraban como pendientes pese a haber inventario; 10 de 27 medicamentos sin existencias.",
     "Medida cautelar y orden de entrega en 48 horas hábiles; prohibición de cerrar puntos.",
     "Controlar el indicador de entrega de pendientes en 48 horas.", "Consultorsalud",
     "https://consultorsalud.com/medida-cautelar-gestor-farmaceutico-colsubsidio/"],
    ["2026-05-29", "Supersalud", "Gestor farmacéutico nacional", "Entrega incompleta (pendientes)",
     "Más de 1 millón de pendientes abiertos por entrega de medicamentos a nivel nacional.",
     "Medida cautelar con plan de normalización: 30 % el primer mes, 60 % el segundo y cierre total.",
     "Plan de choque con metas de cierre de pendientes por mes.", "Vanguardia",
     "https://www.vanguardia.com/colombia/2026/05/29/supersalud-ordena-acciones-urgentes-a-cafam-por-retrasos-en-entrega-de-farmacos/"],
    ["2026-06-18", "Supersalud", "Gestor farmacéutico (Bogotá, Cajicá, Medellín)", "Entrega incompleta (pendientes)",
     "256.296 solicitudes de entrega pendientes en varias regiones, con cinco EPS afectadas.",
     "Medida cautelar: cerrar 30 % en 15 días, 60 % en 30 y 100 % en 45; informes quincenales.",
     "Priorizar pacientes crónicos, de alto costo, menores y adultos mayores.", "El Tiempo",
     "https://www.eltiempo.com/salud/supersalud-ordena-medida-cautelar-contra-el-gestor-farmaceutica-medic-colombia-por-retrasos-en-entrega-de-medicamentos-3565469"],
    ["2026-03-05", "INVIMA", "Sector de medicamentos (nacional)", "Calidad y cadena de frío",
     "En 2025 el INVIMA ejecutó $6.058 millones en multas sanitarias; el 51 % ($3.060 millones) correspondió a medicamentos.",
     "Multas sanitarias ejecutoriadas; Cundinamarca concentró el 44,2 %.",
     "Mantener trazabilidad de temperatura y buenas prácticas de almacenamiento.", "INVIMA",
     "https://www.invima.gov.co/blog/sala-de-prensa-13/el-invima-impuso-mas-de-6-000-millones-en-multas-sanitarias-en-2025-384"],
], columns=["FECHA", "ENTE", "TIPO_DE_ENTIDAD", "CATEGORIA_RELACIONADA", "HALLAZGO", "MEDIDA", "LECCION_PARA_CRM",
            "FUENTE", "URL"])
riesgo = pqr.groupby("CATEGORIA").size().rename("PQRS_CRM")
casos_mercado["PQRS_CRM_MISMA_CATEGORIA"] = casos_mercado["CATEGORIA_RELACIONADA"].map(riesgo).fillna(0).astype(int)
display(casos_mercado[["FECHA", "ENTE", "TIPO_DE_ENTIDAD", "CATEGORIA_RELACIONADA", "PQRS_CRM_MISMA_CATEGORIA"]])

# %% [markdown]
# ## Paso 6. Gráficos para los informes

# %%
import matplotlib.pyplot as plt, textwrap
def pesos(x): return f"${x:,.0f}".replace(",", ".")
plt.rcParams.update({"figure.dpi": 110, "axes.spines.top": False, "axes.spines.right": False, "font.size": 9})
AZUL, NARANJA, GRIS, VERDE, ROJO, MORADO = "#1F4E79", "#D9822B", "#8A8F98", "#2F7D4F", "#B4372F", "#8A6FB0"
def p(a, b): return f"{(a / b * 100 if b else 0):.1f} %".replace(".", ",")
NOTAS = {}   # nota de cada gráfico: se imprime debajo de la imagen y se usa en el informe y el memorando

def guardar_fig(fig, nombre, nota):
    # escribe la nota debajo del gráfico y guarda la imagen
    NOTAS[nombre] = nota
    fig.text(0.01, 0.01, "\n".join(textwrap.wrap(nota, 125)).replace("$", r"\$"), fontsize=8, color="#333", va="bottom", ha="left",
             bbox=dict(boxstyle="round,pad=0.5", fc="#E3ECF5", ec="none"))
    fig.savefig(f"reportes/graficos/{nombre}.png", bbox_inches="tight"); plt.show()

def etiquetar_apilado(ax, tabla, horizontal=False, minimo=3):
    # número y % dentro de cada segmento; total y % del gran total al final de la barra
    tot = tabla.sum(axis=1); gran = tot.sum()
    for cont, col in zip(ax.containers, tabla.columns):
        etiquetas = [f"{int(v)}" if v >= minimo else "" for v, t in zip(tabla[col], tot)]
        ax.bar_label(cont, labels=etiquetas, label_type="center", color="white", fontsize=7 if len(tabla) > 4 else 7.5, fontweight="bold")
    for k, t in enumerate(tot):
        if horizontal: ax.text(t, k, f"  {int(t)}", va="center", fontsize=9, fontweight="bold")
        else: ax.text(k, t, f"{int(t)}", ha="center", va="bottom", fontsize=9, fontweight="bold")
    (ax.set_xlim if horizontal else ax.set_ylim)(0, tot.max() * 1.25)

def barras_h(ax, serie, color, total=None):
    total = total or serie.sum()
    b = ax.barh(serie.index, serie.values, color=color)
    ax.bar_label(b, labels=[f" {int(v)}" for v in serie.values], fontsize=8.5, fontweight="bold")
    ax.set_xlim(0, serie.max() * 1.3)

meses = sorted(pqr["MES"].unique())
B, Dsp = pqr[pqr["CANAL"] == "B2B"], pqr[pqr["CANAL"] != "B2B"]

# 1. Estado de las líneas por comprador
fig, ax = plt.subplots(figsize=(8, 3.8))
r = resumen.set_index("COMPRADOR")[["APROBADAS", "CON_OBSERVACIONES", "RECHAZADAS"]]
r.plot(kind="barh", stacked=True, color=[VERDE, NARANJA, ROJO], ax=ax, legend=False)
etiquetar_apilado(ax, r, horizontal=True, minimo=1)
ax.set_title("Estado de las líneas ofertadas por comprador (número de líneas)"); ax.set_xlabel("Líneas"); ax.set_ylabel("")
ax.legend(["Aprobadas", "Con observaciones", "Rechazadas"], fontsize=7.5, loc="lower right")
peor = resumen.assign(t=resumen["RECHAZADAS"] / resumen["LINEAS"]).sort_values("t").iloc[-1]
fig.subplots_adjust(bottom=0.30)
guardar_fig(fig, "estado_lineas", f"De {len(v)} líneas, {(v['ESTADO'] == 'APROBADA').sum()} ({p((v['ESTADO'] == 'APROBADA').sum(), len(v))}) quedaron aprobadas y "
            f"{(v['ESTADO'] == 'RECHAZADA').sum()} ({p((v['ESTADO'] == 'RECHAZADA').sum(), len(v))}) rechazadas. La mayor proporción de rechazo la tiene "
            f"{peor.COMPRADOR.title()}: {peor.RECHAZADAS} de {peor.LINEAS} líneas ({p(peor.RECHAZADAS, peor.LINEAS)}).")

# 2. Cruce con la Circular 022 y el SISMED
fig, ax = plt.subplots(figsize=(8, 4))
rc = cruce.groupby(["COMPRADOR", "RESULTADO"]).size().unstack(fill_value=0)
colores = {"Dentro del rango": VERDE, "Margen viable sobre SISMED (crédito)": "#5BA57A", "Alineado a SISMED (contado)": "#5BA57A",
           "Sobreprecio: supera el PMV": ROJO, "Riesgo de especulación (reporte SIC)": "#7A1F1A", "Pérdida: tarifa bajo el costo": NARANJA,
           "Margen insuficiente para crédito": "#E0A458", "Bajo SISMED: revisar margen": "#E0A458", "CUM con formato inválido": MORADO,
           "Sin referencia: consultar SISMED": GRIS}
rc.plot(kind="barh", stacked=True, color=[colores.get(c, GRIS) for c in rc.columns], ax=ax, legend=False)
etiquetar_apilado(ax, rc, horizontal=True, minimo=2)
ax.set_title("Cruce con la Circular 022 y el SISMED por comprador (número de líneas)"); ax.set_xlabel(""); ax.set_ylabel("")
ax.legend(rc.columns, fontsize=6.5, loc="upper center", bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False)
fig.subplots_adjust(bottom=0.45)
rcount = cruce["RESULTADO"].value_counts()
guardar_fig(fig, "cruce_circular", f"{rcount.get('Sobreprecio: supera el PMV', 0)} líneas superan el PMV de la Circular 022 y "
            f"{rcount.get('Riesgo de especulación (reporte SIC)', 0)} superan el margen permitido sobre el SISMED (riesgo de reporte ante la SIC); "
            f"juntas suman {pesos(v['EXCEDENTE_MES'].sum())} al mes por encima del tope. {rcount.get('Pérdida: tarifa bajo el costo', 0)} líneas quedan bajo el costo "
            f"({pesos(v['PERDIDA_MES'].sum())}/mes de pérdida) y {rcount.get('Sin referencia: consultar SISMED', 0) + rcount.get('CUM con formato inválido', 0)} no se pueden validar por CUM inválido o sin referencia.")

# 3. PQRS por mes: B2B frente a dispensación
fig, ax = plt.subplots(figsize=(8, 4))
t = pqr.groupby(["MES", "CANAL"]).size().unstack(fill_value=0)[["B2B", "Dispensación"]]
t.plot(kind="bar", stacked=True, color=[AZUL, NARANJA], ax=ax, rot=0, legend=False)
etiquetar_apilado(ax, t)
ax.set_title("PQRS por mes: B2B comercial frente a dispensación"); ax.set_xlabel(""); ax.legend(["B2B comercial", "Dispensación"], fontsize=8, loc="upper left")
fig.subplots_adjust(bottom=0.25)
pico = t.sum(axis=1).idxmax()
guardar_fig(fig, "pqr_mes_canal", f"Qué muestra: cuántas PQRS (quejas y solicitudes, no ventas) llegaron cada mes. Azul = B2B (compradores institucionales y equipos de licitaciones y convenios); "
            f"naranja = dispensación (pacientes). Cifras: {len(pqr)} PQRS en total = {len(B)} B2B ({p(len(B), len(pqr))}) + {len(Dsp)} de dispensación ({p(len(Dsp), len(pqr))}). "
            f"Mes con más casos: {pico}, con {t.sum(axis=1).max()} ({t.loc[pico, 'B2B']} B2B y {t.loc[pico, 'Dispensación']} de dispensación). Las B2B subieron de {t['B2B'].iloc[0]} en junio a "
            f"{t.loc['2026-10', 'B2B']} en octubre por los ajustes de precio de la Circular 022. Noviembre tiene datos hasta el día 20.")

# 4. B2B: mes por solicitante y categoría
fig, (a1, a2) = plt.subplots(1, 2, figsize=(11, 4.2), gridspec_kw={"width_ratios": [1, 1.1]})
sb = B.groupby(["MES", "SOLICITANTE"]).size().unstack(fill_value=0)
sb.columns = [c.replace("Usuario interno – ", "Interno: ") for c in sb.columns]
sb.plot(kind="bar", stacked=True, color=[AZUL, "#4F8A8B", MORADO], ax=a1, rot=0, legend=False, width=0.8)
etiquetar_apilado(a1, sb); a1.set_title("B2B por mes y tipo de solicitante"); a1.set_xlabel(""); a1.legend(sb.columns, fontsize=7, loc="upper left")
barras_h(a2, B["CATEGORIA"].value_counts().sort_values(), AZUL); a2.set_title("B2B por categoría (número de casos)")
fig.subplots_adjust(bottom=0.25, wspace=0.55)
cb = B["CATEGORIA"].value_counts(); internos = B["SOLICITANTE"].str.startswith("Usuario interno").sum()
guardar_fig(fig, "pqr_b2b", f"De {len(B)} casos B2B, {internos} ({p(internos, len(B))}) los radicaron los equipos internos de licitaciones y convenios y "
            f"{len(B) - internos} ({p(len(B) - internos, len(B))}) los compradores. La categoría principal es {cb.index[0]} con {cb.iloc[0]} casos ({p(cb.iloc[0], len(B))}). "
            f"Impacto económico B2B: {pesos(B['IMPACTO_COP'].sum())}.")

# 5. Dispensación: mes con tutelas y categoría
fig, (a1, a2) = plt.subplots(1, 2, figsize=(11, 4.2), gridspec_kw={"width_ratios": [1, 1.1]})
dm = Dsp.groupby("MES").size(); dt = Dsp[Dsp["TUTELA"]].groupby("MES").size().reindex(dm.index, fill_value=0)
x = np.arange(len(dm)); b1 = a1.bar(x - .2, dm.values, .4, color=NARANJA, label="Casos"); b2 = a1.bar(x + .2, dt.values, .4, color=ROJO, label="Con tutela")
a1.bar_label(b1, labels=[f"{v}" for v in dm.values], fontsize=8.5, fontweight="bold")
a1.bar_label(b2, labels=[f"{v}" if v else "" for v in dt.values], fontsize=8, color=ROJO)
a1.set_xticks(x, dm.index); a1.set_ylim(0, dm.max() * 1.3); a1.set_title("Dispensación por mes"); a1.legend(fontsize=7)
barras_h(a2, Dsp["CATEGORIA"].value_counts().sort_values(), NARANJA); a2.set_title("Dispensación por categoría (número de casos)")
fig.subplots_adjust(bottom=0.25, wspace=0.55)
pend = cerradas[cerradas["CATEGORIA"] == "Entrega incompleta (pendientes)"]; p48 = (pend["DIAS_HABILES"] <= 2).sum()
cd = Dsp["CATEGORIA"].value_counts()
guardar_fig(fig, "pqr_dispensacion", f"{len(Dsp)} casos de dispensación; el pico fue {dm.idxmax()} con {dm.max()} ({p(dm.max(), len(Dsp))}). La causa principal es "
            f"{cd.index[0]} ({cd.iloc[0]}, {p(cd.iloc[0], len(Dsp))}). Solo {p48} de {len(pend)} pendientes ({p(p48, len(pend))}) se entregaron en 48 horas "
            f"(Resolución 1604 de 2013) y {Dsp['TUTELA'].sum()} casos ({p(Dsp['TUTELA'].sum(), len(Dsp))}) terminaron en tutela.")

# 6. Acuerdo de servicio por mes y canal
fig, ax = plt.subplots(figsize=(8, 3.8))
for df_, nom, col in [(cerradas[cerradas["CANAL"] == "B2B"], "B2B comercial", AZUL), (cerradas[cerradas["CANAL"] != "B2B"], "Dispensación", NARANJA)]:
    serie = df_.groupby("MES")["CUMPLE_SLA"].mean().mul(100)
    ax.plot(serie.index, serie.values, marker="o", color=col, label=nom)
    for xm, yv in zip(serie.index, serie.values): ax.annotate(f"{yv:.1f} %".replace(".", ","), (xm, yv), textcoords="offset points", xytext=(0, 7 if col == AZUL else -13), ha="center", fontsize=7.5, color=col, fontweight="bold")
ax.axhline(90, color=GRIS, ls="--", lw=1, label="Meta 90 %"); ax.set_ylim(40, 105); ax.legend(fontsize=7.5, loc="lower left")
ax.set_title("Cumplimiento del acuerdo de servicio por mes y canal (%)"); fig.subplots_adjust(bottom=0.27)
slaB = cerradas[cerradas["CANAL"] == "B2B"]["CUMPLE_SLA"].mean() * 100; slaD = cerradas[cerradas["CANAL"] != "B2B"]["CUMPLE_SLA"].mean() * 100
guardar_fig(fig, "oportunidad_mes", f"Acuerdo de servicio del periodo: B2B {slaB:.1f} % y dispensación {slaD:.1f} %, frente a una meta de 90 %; ningún canal la alcanza. "
            f"La oportunidad legal (15 días hábiles) es {kpi_pqr['pct_oportunidad']} %, porque el plazo interno por severidad es más exigente que el legal.".replace(".", ",", 2))

top_excedentes = (v[v["H_SOBRE_PMV"] | v["H_ESPECULACION"]].sort_values("EXCEDENTE_MES", ascending=False)
                  [["COMPRADOR", "CUM", "DESCRIPCION", T, "PMV_INSTITUCIONAL", "PRECIO_SISMED", "PCT_SOBRE_PMV", "PCT_SOBRE_SISMED", "EXCEDENTE_MES"]])
display(top_excedentes.head(10))

# %% [markdown]
# ## Paso 7. Exportar el consolidado en Excel y los datos del portal

# %%
import json, shutil
from openpyxl.formatting.rule import CellIsRule

RUTA_XLSX = "data/salida/consolidado_validacion.xlsx"
cols_salida = ["COMPRADOR", "TAMANO", "MODALIDAD", "ARCHIVO", "FILA_ORIGEN", "CUM", "DESCRIPCION",
               "MERCADO_RELEVANTE", "CANTIDAD ESTIMADA MES", T, "PMV_INSTITUCIONAL", "PRECIO_PISO",
               "PRECIO_SISMED", "PCT_SOBRE_PMV", "PCT_SOBRE_SISMED", "PCT_MARGEN", "PCT_PERDIDA", "EXCEDENTE_MES", "PERDIDA_MES",
               "VALOR_OFERTA_MES", "FORMA_PAGO", "RESULTADO_CRUCE", "ESTADO", "HALLAZGOS"]
indicadores_pqr = (pqr.groupby("CANAL").agg(TOTAL=("RADICADO", "size"), ABIERTAS=("ESTADO", lambda s: s.isin(["En gestión", "Escalada a ente"]).sum()),
                    VENCIDAS=("VENCIDA_ABIERTA", "sum"), OPORTUNIDAD_PCT=("OPORTUNA_LEGAL", lambda s: round(s.dropna().mean() * 100, 1)),
                    SLA_PCT=("CUMPLE_SLA", lambda s: round(s.dropna().mean() * 100, 1)),
                    DIAS_PROMEDIO=("DIAS_HABILES", "mean"), TUTELAS=("TUTELA", "sum")).round(1).reset_index())

with pd.ExcelWriter(RUTA_XLSX, engine="openpyxl") as xw:
    resumen.to_excel(xw, sheet_name="RESUMEN", index=False)
    cruce.to_excel(xw, sheet_name="CRUCE_CIRCULAR_022", index=False)
    lectura.to_excel(xw, sheet_name="LECTURA_ARCHIVOS", index=False)
    v[cols_salida].to_excel(xw, sheet_name="CONSOLIDADO_VALIDADO", index=False)
    v[v["N_HALLAZGOS"] > 0][cols_salida].to_excel(xw, sheet_name="HALLAZGOS", index=False)
    hallazgos_tipo.reset_index().to_excel(xw, sheet_name="HALLAZGOS_POR_TIPO", index=False)
    pqr.to_excel(xw, sheet_name="PQRS", index=False)
    indicadores_pqr.to_excel(xw, sheet_name="INDICADORES_PQRS", index=False)
    metas.to_excel(xw, sheet_name="METAS_REGULATORIAS", index=False)
    casos_mercado.to_excel(xw, sheet_name="CASOS_MERCADO", index=False)

wb = load_workbook(RUTA_XLSX)
for ws in wb.worksheets:
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(50, max(10, max(len(str(x.value or "")) for x in col[:60]) + 2))
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
for nombre in ["CONSOLIDADO_VALIDADO", "HALLAZGOS"]:
    ws = wb[nombre]; col = chr(65 + cols_salida.index("ESTADO")); rango = f"{col}2:{col}{ws.max_row}"
    ws.conditional_formatting.add(rango, CellIsRule(operator="equal", formula=['"RECHAZADA"'], fill=PatternFill("solid", fgColor="F8CBAD")))
    ws.conditional_formatting.add(rango, CellIsRule(operator="equal", formula=['"APROBADA"'], fill=PatternFill("solid", fgColor="C6EFCE")))
ws = wb["CRUCE_CIRCULAR_022"]; cab = [c.value for c in ws[1]]
for nombre, fmt_ in [("TARIFA OFERTADA", '"$"#,##0'), ("PMV CIRCULAR 022", '"$"#,##0'), ("PRECIO PISO (COSTO)", '"$"#,##0'), ("PRECIO SISMED", '"$"#,##0'), ("% SOBRE SISMED", '0.0" %"'),
                     ("EXCEDENTE SOBRE TOPE / MES", '"$"#,##0'), ("PÉRDIDA / MES", '"$"#,##0'), ("% DIFERENCIA VS PMV", '0.0" %"'),
                     ("% GANANCIA SOBRE PMV", '0.0" %"'), ("% MARGEN", '0.0" %"'), ("% PÉRDIDA", '0.0" %"'), ("CANTIDAD MES", '#,##0')]:
    j = cab.index(nombre) + 1
    for (cel,) in ws.iter_rows(min_row=2, min_col=j, max_col=j): cel.number_format = fmt_
col = ws.cell(1, cab.index("RESULTADO") + 1).column_letter; rango = f"{col}2:{col}{ws.max_row}"
from openpyxl.formatting.rule import FormulaRule
for txt, color in [("Sobreprecio", "F8CBAD"), ("Riesgo", "F4B183"), ("Pérdida", "FFC7CE"), ("Margen insuf", "FFE699"), ("Bajo SISMED", "FFE699"),
                   ("Dentro", "C6EFCE"), ("Margen viable", "C6EFCE"), ("Alineado", "C6EFCE"), ("Sin referencia", "D9D9D9"), ("CUM con", "D9D2E9")]:
    ws.conditional_formatting.add(rango, FormulaRule(formula=[f'LEFT({col}2,{len(txt)})="{txt}"'], fill=PatternFill("solid", fgColor=color)))
wb.save(RUTA_XLSX)
shutil.copy(RUTA_XLSX, "docs/descargas/consolidado_validacion.xlsx")

def a_json(df):
    return json.loads(df.to_json(orient="records", date_format="iso", force_ascii=False))
def guardar(nombre, obj):
    json.dump(obj, open(f"docs/data/{nombre}.json", "w", encoding="utf-8"), ensure_ascii=False)

kpis = {"fecha_corte": str(FECHA_CORTE.date()), "empresa": EMPRESA, "compradores": int(len(resumen)),
        "lineas": int(len(v)), "pct_aprobacion": round(float((v["ESTADO"] == "APROBADA").mean() * 100), 1),
        "sobre_pmv": int(v["H_SOBRE_PMV"].sum()), "bajo_piso": int(v["H_BAJO_PISO"].sum()),
        "excedente_mes": float(v["EXCEDENTE_MES"].sum()), "perdida_mes": float(v["PERDIDA_MES"].sum()),
        "especulacion": int(v["H_ESPECULACION"].sum()), "no_regulados": int(no_reg.sum()),
        "margen_promedio": round(float(v["PCT_MARGEN"].mean()), 1), "valor_oferta_mes": float(v["VALOR_OFERTA_MES"].sum()),
        "pqr": kpi_pqr, "presentaciones_circular": int(len(pmv)),
        "nota": "PMV, CUM y medicamentos reales (Circular 022 de 2026). Empresa, compradores, tarifas y PQRS ficticios. Casos del mercado anonimizados."}
guardar("kpis", kpis)
guardar("lineas", a_json(v[["COMPRADOR", "TAMANO", "CUM", "DESCRIPCION", "CANTIDAD ESTIMADA MES", T, "PMV_INSTITUCIONAL",
                            "PRECIO_PISO", "PRECIO_SISMED", "PCT_SOBRE_PMV", "PCT_SOBRE_SISMED", "PCT_MARGEN", "PCT_PERDIDA", "EXCEDENTE_MES", "PERDIDA_MES",
                            "VALOR_OFERTA_MES", "FORMA_PAGO", "RESULTADO_CRUCE", "ESTADO", "HALLAZGOS"]]
                     .rename(columns={"CANTIDAD ESTIMADA MES": "CANTIDAD", T: "TARIFA"})))
pqr_web = pqr.copy()
for c in ["FECHA_RADICACION", "FECHA_RESPUESTA"]:
    pqr_web[c] = pqr_web[c].dt.strftime("%Y-%m-%d")
guardar("pqr", a_json(pqr_web))
guardar("metas", a_json(metas))
guardar("casos", a_json(casos_mercado))
guardar("lectura", a_json(lectura.fillna("")))
guardar("notas", NOTAS)
# referencias para la página "Analizar un archivo" del portal (se procesa en el navegador)
_p = pmv.dropna(subset=["PMV_INSTITUCIONAL"]).drop_duplicates("CUM")
guardar("ref_pmv", {c: [round(float(x), 2), str(m)[:90]] for c, x, m in zip(_p["CUM"], _p["PMV_INSTITUCIONAL"], _p["MEDICAMENTO"])})
guardar("ref_sismed", {c: [round(float(x), 2), str(a)[:40], str(d)[:70], str(y)[:4] if pd.notna(y) else ""] for c, x, a, d, y in zip(sismed["CUM"], sismed["PRECIO_SISMED"], sismed["PRINCIPIO_ACTIVO"], sismed["DESCRIPCION_SISMED"], sismed["ANO_REFERENCIA"])})
guardar("ref_termometro", {str(e): [round(float(p), 2), str(u), str(m)[:90], str(c)] for e, p, u, m, c in zip(termo["EXPEDIENTE"], termo["PRECIO_UNIDAD_REFERENCIA"], termo["UNIDAD_DISPENSACION"], termo["MEDICAMENTO_2024"], termo["CANAL_REFERENCIA"]) if pd.notna(p)})
guardar("ref_catalogo", {c: round(float(x), 2) for c, x in zip(catalogo["CUM"], catalogo["PRECIO_PISO"])})
print("Exportado:", RUTA_XLSX, "y docs/data/*.json")

# %% [markdown]
# ## Paso 8. Documento gerencial (Word)

# %%
from docx import Document
from docx.shared import Pt, Inches

doc = Document(); st = doc.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(10.5)
doc.add_heading("Memorando gerencial – Validación de ofertas y seguimiento de PQRS", level=1)
doc.add_paragraph(f"Para: Gerencia Comercial y de Licitaciones\nDe: Coordinación de PQRS y apoyo a licitaciones – {EMPRESA}\n"
                  f"Fecha de corte: {FECHA_CORTE:%d/%m/%Y}")
peor = resumen.sort_values("EXCEDENTE_MES", ascending=False).iloc[0]
doc.add_heading("1. Resumen ejecutivo", level=2)
resumen_texto = (
    f"Se consolidaron {len(v)} líneas de {len(resumen)} compradores, cruzadas con {len(pmv):,} presentaciones de la Circular 022 de 2026. "
    f"El {kpis['pct_aprobacion']} % quedó aprobado; {kpis['sobre_pmv']} líneas superan el PMV institucional "
    f"(excedente estimado {pesos(kpis['excedente_mes'])}/mes, concentrado en {peor.COMPRADOR}). "
    f"En PQRS se recibieron {kpi_pqr['total']} casos: oportunidad legal {kpi_pqr['pct_oportunidad']} %, acuerdo de servicio "
    f"{kpi_pqr['pct_sla']} %, {kpi_pqr['vencidas_abiertas']} abiertos fuera de término, {kpi_pqr['tutelas']} tutelas y "
    f"satisfacción promedio de {kpi_pqr['satisfaccion']}/5.")
doc.add_paragraph(resumen_texto)
doc.add_heading("2. Resultado por comprador", level=2)
t = doc.add_table(rows=1, cols=6); t.style = "Light Grid Accent 1"
for c, txt in zip(t.rows[0].cells, ["Comprador", "Líneas", "% aprobación", "Sobre PMV", "Bajo piso", "Excedente/mes"]): c.text = txt
for _, r in resumen.iterrows():
    for c, txt in zip(t.add_row().cells, [r.COMPRADOR, r.LINEAS, f"{r.PCT_APROBACION} %", r.SOBRE_PMV, r.BAJO_PISO, pesos(r.EXCEDENTE_MES)]):
        c.text = str(txt)
doc.add_picture("reportes/graficos/estado_lineas.png", width=Inches(6.2))
doc.add_heading("2B. Cruce con la Circular 022 y el SISMED", level=2)
doc.add_picture("reportes/graficos/cruce_circular.png", width=Inches(6.2))
doc.add_heading("3. Metas regulatorias", level=2)
for _, m in metas.iterrows():
    doc.add_paragraph(f"[{m.SEMAFORO}] {m.INDICADOR} ({m.ENTE}, {m.NORMATIVA}): resultado {m.RESULTADO} % frente a meta {m.META} %.",
                      style="List Bullet")
doc.add_heading("4. Gestión de PQRS", level=2)
top_m = pqr["CATEGORIA"].value_counts().head(3)
doc.add_paragraph("Principales categorías: " + "; ".join(f"{m} ({n})" for m, n in top_m.items()) +
                  f". El {kpi_pqr['pct_internos']} % de los casos provino de usuarios internos de licitaciones y convenios.")
doc.add_heading("4A. PQRS comerciales B2B", level=3)
doc.add_picture("reportes/graficos/pqr_b2b.png", width=Inches(6.4))
doc.add_heading("4B. PQRS de dispensación", level=3)
doc.add_picture("reportes/graficos/pqr_dispensacion.png", width=Inches(6.4))
doc.add_picture("reportes/graficos/oportunidad_mes.png", width=Inches(6.2))
doc.add_heading("5. Referentes del mercado", level=2)
for _, c in casos_mercado.head(4).iterrows():
    doc.add_paragraph(f"{c.FECHA} · {c.ENTE} · {c.TIPO_DE_ENTIDAD}: {c.HALLAZGO} Lección: {c.LECCION_PARA_CRM}", style="List Bullet")
doc.add_heading("6. Recomendaciones", level=2)
recomendaciones = [
    f"Devolver las {kpis['sobre_pmv']} líneas que superan el PMV institucional y ajustar tarifas antes de legalizar los convenios.",
    f"Ajustar las {int(v['H_ESPECULACION'].sum())} líneas no reguladas que superan el margen permitido sobre el SISMED (22 % a 32 % a crédito; ±5 % de contado) para evitar reportes por especulación ante la SIC.",
    "Solicitar corrección de los CUM inválidos o que no figuran en la Circular 022.",
    f"Revisar con el área comercial las {kpis['bajo_piso']} líneas bajo el precio piso.",
    f"Plan de choque para los {kpi_pqr['vencidas_abiertas']} casos abiertos fuera de término y para la entrega de pendientes en 48 horas.",
    "Estandarizar respuestas tipo para ajustes de precio por la Circular 022, la categoría que más creció desde agosto.",
]
for txt in recomendaciones: doc.add_paragraph(txt, style="List Number")
doc.add_paragraph("Nota: PMV reales; empresa, compradores, tarifas y PQRS ficticios; casos del mercado anonimizados.").runs[0].italic = True
doc.save("reportes/memorando_gerencial.docx")
shutil.copy("reportes/memorando_gerencial.docx", "docs/descargas/memorando_gerencial.docx")
for g in glob.glob("reportes/graficos/*.png"): shutil.copy(g, "docs/img/")
print("Documento gerencial: reportes/memorando_gerencial.docx")

# %% [markdown]
# ## Paso 8B. Informe ejecutivo en HTML
#
# Genera `docs/informe_ejecutivo.html` (copia en `salida`): una sola página con indicadores, metas, tablas, gráficos incrustados, casos del mercado y recomendaciones. Se abre sin internet, se puede enviar por correo o imprimir en PDF.

# %%
import base64, html as H
def img64(ruta): return "data:image/png;base64," + base64.b64encode(open(ruta, "rb").read()).decode()
def fmt_num(x):
    return (f"{x:,.1f}" if isinstance(x, float) and not float(x).is_integer() else f"{int(x):,}").replace(",", "X").replace(".", ",").replace("X", ".")
def tabla_html(df, num=(), dinero=()):
    cab = "".join(f"<th>{H.escape(str(c))}</th>" for c in df.columns)
    filas = ""
    for _, r in df.iterrows():
        tds = ""
        for c in df.columns:
            x = r[c]
            if c in dinero and pd.notna(x): tds += f'<td class="n">{pesos(x)}</td>'
            elif c in num and pd.notna(x): tds += f'<td class="n">{fmt_num(x)}</td>'
            else: tds += f"<td>{'' if pd.isna(x) else H.escape(str(x))}</td>"
        filas += f"<tr>{tds}</tr>"
    return f'<div class="tw"><table><thead><tr>{cab}</tr></thead><tbody>{filas}</tbody></table></div>'

COLOR_SEM = {"Verde": "ok", "Amarillo": "warn", "Rojo": "bad"}
kp = [("Líneas validadas", fmt_num(kpis["lineas"]), ""), ("Aprobadas", f"{kpis['pct_aprobacion']} %", "warn" if kpis["pct_aprobacion"] < 80 else "ok"),
      ("Sobre el PMV", str(kpis["sobre_pmv"]), "bad" if kpis["sobre_pmv"] else "ok"),
      ("Excedente / mes", f"${kpis['excedente_mes']/1e6:,.1f} M".replace(".", ","), "bad"),
      ("PQRS", fmt_num(kpi_pqr["total"]), ""), ("Oportunidad legal", f"{kpi_pqr['pct_oportunidad']} %", "ok" if kpi_pqr["pct_oportunidad"] >= 90 else "warn"),
      ("Fuera de término", str(kpi_pqr["vencidas_abiertas"]), "bad" if kpi_pqr["vencidas_abiertas"] else "ok"),
      ("Satisfacción", f"{kpi_pqr['satisfaccion']}/5", "")]
metas_html = "".join(f'<div class="meta {COLOR_SEM[m.SEMAFORO]}"><b>{m.RESULTADO} %</b><span>meta {"≤" if m.SENTIDO=="menor" else "≥"} {m.META:g} %</span>'
                     f'<p>{H.escape(m.INDICADOR)}</p><small>{m.ENTE} · {H.escape(m.NORMATIVA)}</small></div>' for m in metas.itertuples())
casos_html = "".join(f'<div class="caso"><small>{c.FECHA} · {c.ENTE}</small><b>{H.escape(c.TIPO_DE_ENTIDAD)}</b><p>{H.escape(c.HALLAZGO)}</p>'
                     f'<p class="mut">{H.escape(c.MEDIDA)}</p><p class="lec">Lección: {H.escape(c.LECCION_PARA_CRM)}</p>'
                     f'<a href="{c.URL}" target="_blank" rel="noopener">Fuente: {c.FUENTE}</a></div>' for c in casos_mercado.itertuples())
t_res = resumen[["COMPRADOR", "TAMANO", "MODALIDAD", "LINEAS", "PCT_APROBACION", "SOBRE_PMV", "BAJO_PISO", "EXCEDENTE_MES"]].rename(
    columns={"TAMANO": "TAMAÑO", "PCT_APROBACION": "% APROB.", "SOBRE_PMV": "SOBRE PMV", "BAJO_PISO": "BAJO PISO", "EXCEDENTE_MES": "EXCEDENTE/MES"})
t_top = top_excedentes.head(8).rename(columns={T: "TARIFA", "PMV_INSTITUCIONAL": "PMV", "PRECIO_SISMED": "SISMED", "PCT_SOBRE_PMV": "% SOBRE PMV", "PCT_SOBRE_SISMED": "% SOBRE SISMED", "EXCEDENTE_MES": "EXCEDENTE/MES"})
t_cruce = cruce[~cruce["RESULTADO"].isin(["Dentro del rango"])].sort_values(["RESULTADO", "% GANANCIA SOBRE PMV", "% PÉRDIDA"], ascending=[True, False, False])[
    ["COMPRADOR", "FORMA DE PAGO", "MEDICAMENTO", "TARIFA OFERTADA", "PMV CIRCULAR 022", "% GANANCIA SOBRE PMV", "PRECIO SISMED", "% SOBRE SISMED", "% PÉRDIDA", "RESULTADO"]]
def fig_html(nombre, alt):
    return f'<div class="card"><img src="{img64(f"reportes/graficos/{nombre}.png")}" alt="{alt}"><p class="nota">{H.escape(NOTAS.get(nombre, ""))}</p></div>'
ind_b2b = indicadores_pqr[indicadores_pqr["CANAL"] == "B2B"]; ind_dis = indicadores_pqr[indicadores_pqr["CANAL"] != "B2B"]
t_venc = (pqr[pqr["VENCIDA_ABIERTA"]].sort_values("DIAS_HABILES", ascending=False)
          [["RADICADO", "CANAL", "CIUDAD", "CATEGORIA", "SEVERIDAD", "DIAS_HABILES"]].rename(columns={"DIAS_HABILES": "DÍAS HÁBILES"}))

INFORME = f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Informe Ejecutivo – {EMPRESA}</title><style>
:root{{--bg:#f5f6f4;--card:#fff;--ink:#1b2230;--mut:#5d6675;--line:#dfe3e8;--brand:#1f4e79;--soft:#e3ecf5;--ok:#2f7d4f;--warn:#b86b12;--bad:#b4372f}}
@media (prefers-color-scheme:dark){{:root{{--bg:#11161d;--card:#18202a;--ink:#e7ebf0;--mut:#9aa4b2;--line:#2a3542;--brand:#7fb0de;--soft:#1e3044;--ok:#6fcf97;--warn:#f2b356;--bad:#f08a80}} img{{background:#fff;border-radius:6px}}}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);font:15px/1.55 system-ui,"Segoe UI",Roboto,sans-serif}}
.w{{max-width:1080px;margin:0 auto;padding:28px 16px 48px}}h1{{font-size:clamp(22px,3vw,30px);margin:0 0 4px}}h2{{font-size:18px;margin:32px 0 10px;color:var(--brand)}}
.tag{{display:inline-block;font-size:12px;color:var(--warn);border:1px solid var(--warn);padding:2px 8px;border-radius:4px;margin-bottom:10px}}.mut{{color:var(--mut)}}
.k{{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin:20px 0}}.kpi{{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px}}
.kpi .l{{color:var(--mut);font-size:12.5px}}.kpi .v{{font-size:21px;font-weight:700}}.ok .v,.meta.ok b{{color:var(--ok)}}.warn .v,.meta.warn b{{color:var(--warn)}}.bad .v,.meta.bad b{{color:var(--bad)}}
.metas,.casos{{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:10px}}
.meta,.caso{{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px 14px;border-left:5px solid var(--line)}}
.meta.ok{{border-left-color:var(--ok)}}.meta.warn{{border-left-color:var(--warn)}}.meta.bad{{border-left-color:var(--bad)}}
.meta b{{font-size:22px}}.meta span{{color:var(--mut);margin-left:8px;font-size:13px}}.meta p{{margin:6px 0 4px}}.meta small,.caso small{{color:var(--mut)}}
.caso b{{display:block;margin:2px 0 6px}}.caso p{{margin:4px 0;font-size:14px}}.caso .lec{{color:var(--brand)}}.caso a{{font-size:13px;color:var(--brand)}}
.g{{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:12px}}.card{{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px}}img{{width:100%;height:auto}}
.nota{{font-size:13px;background:var(--soft);border-radius:6px;padding:8px 10px;margin:8px 0 0}}
.tw{{overflow-x:auto;border:1px solid var(--line);border-radius:8px;background:var(--card)}}table{{border-collapse:collapse;width:100%;font-size:13px}}
th,td{{padding:7px 10px;border-bottom:1px solid var(--line);text-align:left}}th{{background:var(--soft);white-space:nowrap}}td.n{{text-align:right;white-space:nowrap}}
footer{{margin-top:36px;color:var(--mut);font-size:12.5px;border-top:1px solid var(--line);padding-top:12px}}
@media print{{body{{background:#fff}}.card,.tw,.meta,.caso{{break-inside:avoid}}}}</style></head><body><div class="w">
<span class="tag">PMV REALES · EMPRESA, COMPRADORES Y PQRS FICTICIOS · CASOS DEL MERCADO ANONIMIZADOS</span>
<h1>Informe ejecutivo: validación de ofertas y seguimiento de PQRS</h1>
<p class="mut"><b>{EMPRESA}</b> · Coordinación de PQRS y apoyo a licitaciones · Corte {FECHA_CORTE:%d/%m/%Y}</p>
<div class="k">{"".join(f'<div class="kpi {c}"><div class="l">{l}</div><div class="v">{x}</div></div>' for l, x, c in kp)}</div>
<h2>1. Resumen</h2><p>{H.escape(resumen_texto)}</p>
<h2>2. Metas regulatorias</h2><div class="metas">{metas_html}</div>
<h2>3. Resultado por comprador</h2>{tabla_html(t_res, num=("LINEAS", "% APROB.", "SOBRE PMV", "BAJO PISO"), dinero=("EXCEDENTE/MES",))}
<div class="g" style="margin-top:12px">{fig_html("estado_lineas", "Estado de líneas")}</div>
<h2>4. Cruce con la Circular 022 y el SISMED: sobreprecios, especulación y pérdidas</h2>
<p>Regulados: la tarifa no puede superar el PMV institucional. No regulados: margen viable sobre el SISMED de {MARGEN_MIN_CREDITO} % a {MARGEN_MAX_CREDITO} % a crédito de 90 días y ±{TOLERANCIA_CONTADO} % de contado; por encima hay riesgo de reporte por especulación ante la SIC. Precio SISMED: 2024 (Termómetro de Precios de Clicsalud) cuando existe; si no, 2019 (consulta pública del SISMED).</p>
<p>Pérdida estimada por tarifas bajo el costo: <b>{pesos(kpis['perdida_mes'])}</b> al mes · Margen promedio de las ofertas: <b>{kpis['margen_promedio']} %</b>.</p>
{tabla_html(t_cruce, num=("% GANANCIA SOBRE PMV", "% SOBRE SISMED", "% PÉRDIDA"), dinero=("TARIFA OFERTADA", "PMV CIRCULAR 022", "PRECIO SISMED"))}
<div class="g" style="margin-top:12px">{fig_html("cruce_circular", "Cruce con la Circular 022 y el SISMED")}</div>
<h2>5. PQRS: visión general</h2>
<div class="g">{fig_html("pqr_mes_canal", "PQRS por mes y canal")}{fig_html("oportunidad_mes", "Acuerdo de servicio por mes y canal")}</div>
<h2>5A. PQRS comerciales B2B (convenios y licitaciones)</h2>{tabla_html(ind_b2b, num=("TOTAL", "ABIERTAS", "VENCIDAS", "OPORTUNIDAD_PCT", "SLA_PCT", "DIAS_PROMEDIO", "TUTELAS"))}
<div style="margin-top:12px">{fig_html("pqr_b2b", "PQRS B2B")}</div>
<h2>5B. PQRS de dispensación a pacientes</h2>{tabla_html(ind_dis, num=("TOTAL", "ABIERTAS", "VENCIDAS", "OPORTUNIDAD_PCT", "SLA_PCT", "DIAS_PROMEDIO", "TUTELAS"))}
<div style="margin-top:12px">{fig_html("pqr_dispensacion", "PQRS de dispensación")}</div>
<h2>PQRS abiertas fuera de término</h2>{tabla_html(t_venc, num=("DÍAS HÁBILES",)) if len(t_venc) else "<p>No hay casos fuera de término.</p>"}
<h2>6. Referentes del mercado (casos reales anonimizados)</h2><div class="casos">{casos_html}</div>
<h2>7. Recomendaciones</h2><ol>{"".join(f"<li>{H.escape(x)}</li>" for x in recomendaciones)}</ol>
<footer>Generado automáticamente desde Google Colab · PMV: tabla pública de la Circular 022 de 2026 (CNPMDM). Precios de referencia de no regulados: Termómetro de Precios de Clicsalud 2024 (datos.gov.co n4dj-8r7k) y SISMED 2019 (3he6-m866). {EMPRESA}, compradores, tarifas y PQRS son ficticios; los casos del mercado son públicos y se presentan sin nombres de empresas. Elaboró: Ivan Santiago Herrera Baquero.</footer>
</div></body></html>"""
open("docs/informe_ejecutivo.html", "w", encoding="utf-8").write(INFORME)
shutil.copy("docs/informe_ejecutivo.html", "data/salida/informe_ejecutivo.html")
print("Informe HTML: docs/informe_ejecutivo.html (copia en salida)")

# %% [markdown]
# ## Paso 9. Publicar en GitHub (actualiza el portal)
#
# **Una sola vez:** crea en GitHub un token *fine-grained* con permiso **Contents: Read and write** solo para este repositorio y guárdalo en Colab (llave 🔑 *Secrets*) como `GITHUB_TOKEN`. Nunca lo pegues en el código.

# %%
PUBLICAR = False   # cámbialo a True cuando quieras subir los resultados

if PUBLICAR and EN_COLAB:
    from google.colab import userdata
    token = userdata.get("GITHUB_TOKEN")
    url = f"https://{USUARIO_GITHUB}:{token}@github.com/{USUARIO_GITHUB}/{REPO}.git"
    destino = "/content/repo"; shutil.rmtree(destino, ignore_errors=True)
    r = subprocess.run(["git", "clone", url, destino], capture_output=True, text=True)
    if r.returncode != 0: raise RuntimeError(r.stderr.replace(token, "***"))
    for origen, dest in [("data/entrada", "data/entrada"), ("data/referencia", "data/referencia"), ("data/salida", "data/salida"),
                         ("docs/data", "docs/data"), ("docs/descargas", "docs/descargas"), ("docs/img", "docs/img"), ("reportes", "reportes")]:
        shutil.copytree(origen, os.path.join(destino, dest), dirs_exist_ok=True, ignore=shutil.ignore_patterns("*.xlsb", "~$*"))
    shutil.copy("docs/informe_ejecutivo.html", os.path.join(destino, "docs/informe_ejecutivo.html"))
    git = lambda *a: subprocess.run(["git", "-C", destino, *a], capture_output=True, text=True)
    git("config", "user.name", "Ivan Santiago Herrera Baquero")
    git("config", "user.email", f"{USUARIO_GITHUB}@users.noreply.github.com")
    git("add", "-A"); c = git("commit", "-m", f"Actualiza validación y PQRS – corte {FECHA_CORTE.date()}")
    p = git("push", "origin", "HEAD:main")
    print(c.stdout.strip()[:200] or c.stderr.strip()[:200])
    print("push →", "ok" if p.returncode == 0 else p.stderr.replace(token, "***")[:300])
    print(f"Portal: https://{USUARIO_GITHUB}.github.io/{REPO}/")
else:
    print("Publicación desactivada. Cambia PUBLICAR = True para subir los resultados.")

# %% [markdown]
# ## Paso 10. Asistente de IA para leer cualquier documento
#
# Para archivos que no siguen ningún formato (cotizaciones con títulos raros, listas de precios, PDF de circulares o resoluciones, documentos en Word), este paso usa un modelo de lenguaje para **entender el documento** y devolver de forma estructurada:
#
# - qué tipo de documento es (oferta, lista de precios, norma, base de PQRS u otro);
# - en qué hoja y fila están los datos, y qué columna es el **CUM**, la **tarifa**, la **cantidad**, el **precio máximo**, el código **MAPIPOS/CUPS** (si aparece) y la descripción;
# - los **topes de regulación** que mencione (circular, precio máximo, fecha de vigencia) y los CUM con su precio cuando el documento es una norma o un PDF.
#
# Después, el cruce lo hace el código (no la IA): cada CUM se compara con la **Circular 022** y, si no está regulado, con el **SISMED**, aplicando las reglas de margen del Paso 1D. La IA solo lee; no inventa precios.
#
# **Cómo usarlo:** deja los documentos en la carpeta `TECNOLOGIAS EN SALUD/ia` de tu Drive y ejecuta la celda. Los resultados quedan en `salida/analisis_ia_<archivo>.xlsx`.
#
# **Modelo:** usa la IA integrada de Colab si tu cuenta la tiene. Si no, usa Gemini con una llave gratuita de Google AI Studio guardada en los Secrets de Colab (🔑) como `GOOGLE_API_KEY`, o Claude con `ANTHROPIC_API_KEY`. Si no hay ninguna, usa la lectura por contenido (sin IA), que reconoce las columnas por la forma de los datos.

# %%
import json, re, io
CARPETA_IA = os.path.join(CARPETA_DRIVE, "ia") if EN_COLAB else "data/ia"
os.makedirs(CARPETA_IA, exist_ok=True)
EXT_IA = (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv", ".txt", ".pdf", ".docx")
if EN_COLAB:
    subprocess.run([sys.executable, "-m", "pip", "-q", "install", "pdfplumber", "google-generativeai", "anthropic"])

# ---------- 1. Modelo de lenguaje disponible
def preparar_ia():
    try:                                    # IA integrada de Colab
        from google.colab import ai as colab_ai
        colab_ai.generate_text("Responde solo: ok")
        return "Colab AI", lambda prompt: colab_ai.generate_text(prompt)
    except Exception:
        pass
    clave = lambda n: (__import__("google.colab", fromlist=["userdata"]).userdata.get(n) if EN_COLAB else os.environ.get(n))
    try:                                    # Gemini (llave gratuita de Google AI Studio)
        k = clave("GOOGLE_API_KEY")
        if k:
            import google.generativeai as genai
            genai.configure(api_key=k); modelo = genai.GenerativeModel("gemini-1.5-flash")
            return "Gemini", lambda prompt: modelo.generate_content(prompt).text
    except Exception:
        pass
    try:                                    # Claude (Anthropic)
        k = clave("ANTHROPIC_API_KEY")
        if k:
            import anthropic
            cli = anthropic.Anthropic(api_key=k)
            return "Claude", lambda prompt: cli.messages.create(model="claude-sonnet-4-5", max_tokens=4000,
                                                              messages=[{"role": "user", "content": prompt}]).content[0].text
    except Exception:
        pass
    return None, None

NOMBRE_IA, preguntar = preparar_ia()
print("Asistente:", NOMBRE_IA or "sin modelo de IA disponible → lectura por contenido")

# ---------- 2. Contenido del documento como texto
def contenido(ruta, filas=60):
    ext = os.path.splitext(ruta)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv"):
        hojas = leer_hojas(ruta)
        return "\n\n".join(f"### HOJA: {h} (fila 0 = primera fila)\n" +
                           df.head(filas).fillna("").astype(str).to_csv(sep="|", header=False) for h, df in hojas.items()), hojas
    if ext == ".pdf":
        import pdfplumber
        with pdfplumber.open(ruta) as pdf:
            return "\n".join((pg.extract_text() or "") for pg in pdf.pages[:15]), None
    if ext == ".docx":
        from docx import Document as _D
        d = _D(ruta)
        texto = "\n".join(p.text for p in d.paragraphs)
        for t in d.tables:
            texto += "\n" + "\n".join("|".join(c.text for c in r.cells) for r in t.rows)
        return texto, None
    return open(ruta, encoding="utf-8", errors="ignore").read(), None

INSTRUCCION = (
    'Eres analista de precios de medicamentos en Colombia. Lee el documento y responde SOLO con un JSON válido, sin texto adicional:\n'
    '{"tipo": "oferta|lista_precios|norma|pqrs|otro",\n'
    ' "hoja": "nombre de la hoja con los datos o null",\n'
    ' "fila_encabezado": número de la fila del encabezado (0 = primera fila) o null,\n'
    ' "columnas": {"CUM": "texto exacto del encabezado o null", "TARIFA": "...", "CANTIDAD": "...", "DESCRIPCION": "...", "PRECIO_MAXIMO": "...", "MAPIPOS_CUPS": "..."},\n'
    ' "forma_de_pago": "CONTADO|CRÉDITO|null",\n'
    ' "topes_regulatorios": [{"norma": "p. ej. Circular 022 de 2026", "detalle": "qué regula", "vigencia": "fecha o null"}],\n'
    ' "registros": [{"cum": "expediente-consecutivo", "descripcion": "...", "precio": número, "tipo_precio": "tarifa|precio_maximo|otro"}],\n'
    ' "resumen": "una o dos frases en español"}\n'
    'Reglas: el CUM tiene la forma número-número (p. ej. 19911481-1). En "registros" incluye solo CUM y precios que aparezcan literalmente en el texto (máximo 200); si el documento es una tabla con encabezado, deja "registros" vacío y llena "columnas". No inventes valores.\n'
    'DOCUMENTO:\n'
)

def json_de(texto):
    m = re.search(r"\{.*\}", texto, re.S)
    return json.loads(m.group(0)) if m else {}

# ---------- 3. Lectura por contenido (sin IA): columna con forma de CUM + columna numérica de precio
RX_CUM = re.compile(r"^\s*(\d{4,9})\s*-\s*(\d{1,3})\s*$")
def leer_por_contenido(hojas):
    mejor = None
    for h, df in hojas.items():
        txt = df.astype(str)
        conteo = {c: txt[c].str.match(RX_CUM).sum() for c in df.columns}
        col_cum = max(conteo, key=conteo.get)
        if conteo[col_cum] < 2: continue
        datos = df[txt[col_cum].str.match(RX_CUM)]
        largos0 = {c: datos[c].astype(str).str.len().mean() for c in df.columns if c != col_cum}
        cand = []
        for c in df.columns:
            if c == col_cum: continue
            nums = datos[c].map(a_numero)
            if nums.notna().mean() >= .6 and (nums > 0).mean() >= .6:
                cand.append((c, nums.median(), (nums.dropna() % 1 == 0).mean()))
        if not cand:   # lista sin precio: consulta de referencias
            desc = max(largos0, key=largos0.get) if largos0 else None
            t = pd.DataFrame({"CUM": datos[col_cum].astype(str).str.replace(" ", ""), "TARIFA": np.nan, "CANTIDAD": np.nan,
                              "DESCRIPCION": datos[desc].astype(str) if desc is not None else ""})
            if mejor is None or len(t) > len(mejor[1]): mejor = (h, t)
            continue
        precio = max(cand, key=lambda x: x[1])[0]
        cant = min([x for x in cand if x[0] != precio and x[2] > .9], key=lambda x: x[1], default=(None,))[0]
        largos = {c: datos[c].astype(str).str.len().mean() for c in df.columns if c not in (col_cum, precio, cant)}
        desc = max(largos, key=largos.get) if largos else None
        t = pd.DataFrame({"CUM": datos[col_cum].astype(str).str.replace(" ", ""), "TARIFA": datos[precio].map(a_numero),
                          "CANTIDAD": datos[cant].map(a_numero) if cant is not None else np.nan,
                          "DESCRIPCION": datos[desc].astype(str) if desc is not None else ""})
        if mejor is None or len(t) > len(mejor[1]): mejor = (h, t)
    return mejor

# ---------- 4. Cruce con Circular 022 y SISMED (reglas del Paso 1D)
ref_pmv = pmv.drop_duplicates("CUM").set_index("CUM")["PMV_INSTITUCIONAL"]
ref_sis = sismed.drop_duplicates("CUM").set_index("CUM")["PRECIO_SISMED"]
def cruzar(t, forma_pago="CRÉDITO"):
    t = t.copy()
    t["CUM"] = t["CUM"].astype(str).str.upper().str.replace(" ", "")
    t["PMV_CIRCULAR_022"] = t["CUM"].map(ref_pmv); t["PRECIO_SISMED"] = t["CUM"].map(ref_sis)
    fx = [precio_por_expediente(c, d) for c, d in zip(t["CUM"], t["DESCRIPCION"])]
    t["PRECIO_2024_UNIDAD_EXPEDIENTE"] = [x[1] for x in fx]; t["MEDICAMENTO_TERMOMETRO"] = [x[3] for x in fx]
    falta = t["PMV_CIRCULAR_022"].isna() & t["PRECIO_SISMED"].isna()
    t["PRECIO_SISMED"] = t["PRECIO_SISMED"].astype(float).where(~falta, pd.Series([np.nan if x[0] is None else x[0] for x in fx], index=t.index, dtype=float))
    t["%_SOBRE_PMV"] = ((t["TARIFA"] / t["PMV_CIRCULAR_022"] - 1) * 100).round(1)
    t["%_SOBRE_SISMED"] = np.where(t["PMV_CIRCULAR_022"].isna(), (t["TARIFA"] / t["PRECIO_SISMED"] - 1) * 100, np.nan).round(1)
    contado = "CONTADO" in str(forma_pago).upper()
    def regla(r):
        if not RX_CUM.match(str(r.CUM)): return "CUM con formato inválido"
        if pd.isna(r.TARIFA):   # lista sin precio: solo se informa la referencia
            if pd.notna(r.PMV_CIRCULAR_022): return "Regulado: PMV Circular 022"
            if pd.notna(r.PRECIO_SISMED): return "No regulado: precio SISMED por presentación"
            if pd.notna(r.PRECIO_2024_UNIDAD_EXPEDIENTE): return "No regulado: precio 2024 por unidad (expediente)"
            return "Sin referencia pública"
        if pd.notna(r.PMV_CIRCULAR_022): return "Sobreprecio: supera el PMV" if r.TARIFA > r.PMV_CIRCULAR_022 else "Dentro del PMV"
        if pd.notna(r.PRECIO_SISMED):
            m = r["%_SOBRE_SISMED"]
            if contado: return "Riesgo de especulación (reporte SIC)" if m > TOLERANCIA_CONTADO else ("Bajo SISMED: revisar margen" if m < -TOLERANCIA_CONTADO else "Alineado a SISMED (contado)")
            return "Margen insuficiente para crédito" if m < MARGEN_MIN_CREDITO else ("Margen viable (crédito)" if m <= MARGEN_MAX_CREDITO else "Riesgo de especulación (reporte SIC)")
        if pd.notna(r.PRECIO_2024_UNIDAD_EXPEDIENTE): return "Solo precio por unidad 2024: indicar unidades por caja"
        return "Sin referencia: consultar SISMED"
    t["RESULTADO"] = t.apply(regla, axis=1)
    return t

# ---------- 5. Recorrer la carpeta
resultados_ia = []
for ruta in sorted(glob.glob(os.path.join(CARPETA_IA, "*"))):
    if not ruta.lower().endswith(EXT_IA) or os.path.basename(ruta).startswith("~$"):
        continue
    nombre = os.path.basename(ruta); print(f"\n=== {nombre}")
    texto, hojas = contenido(ruta)
    info = {}
    if preguntar:
        try:
            info = json_de(preguntar(INSTRUCCION + texto[:30000]))
            print("IA:", info.get("resumen", ""), "| tipo:", info.get("tipo"))
            for tope in info.get("topes_regulatorios", [])[:5]:
                print("  Tope:", tope.get("norma"), "-", tope.get("detalle"), "-", tope.get("vigencia"))
        except Exception as e:
            print("La IA no respondió bien, se usa lectura por contenido:", str(e)[:120])
    tabla = None
    col = info.get("columnas") or {}
    if hojas and info.get("hoja") in hojas and col.get("CUM"):
        df = hojas[info["hoja"]]; f = int(info.get("fila_encabezado") or 0)
        cab = [str(x).strip() for x in df.iloc[f].values]; datos = df.iloc[f + 1:].copy(); datos.columns = cab
        pick = lambda k: datos[col[k]] if col.get(k) in datos.columns else np.nan
        tabla = pd.DataFrame({"CUM": pick("CUM"), "TARIFA": pick("TARIFA") if col.get("TARIFA") else (pick("PRECIO_MAXIMO") if col.get("PRECIO_MAXIMO") else np.nan),
                              "CANTIDAD": pick("CANTIDAD"), "DESCRIPCION": pick("DESCRIPCION")})
        tabla["TARIFA"] = tabla["TARIFA"].map(a_numero); tabla["CANTIDAD"] = pd.to_numeric(tabla["CANTIDAD"].map(a_numero), errors="coerce")
        tabla = tabla.dropna(subset=["CUM"]); tabla = tabla[tabla["CUM"].astype(str).str.strip() != ""]
        metodo = f"IA ({NOMBRE_IA}): columnas"
    elif info.get("registros"):
        tabla = pd.DataFrame(info["registros"]).rename(columns={"cum": "CUM", "precio": "TARIFA", "descripcion": "DESCRIPCION"})
        tabla["CANTIDAD"] = np.nan; tabla["TARIFA"] = pd.to_numeric(tabla["TARIFA"], errors="coerce")
        metodo = f"IA ({NOMBRE_IA}): registros del texto"
    elif hojas:
        r = leer_por_contenido(hojas)
        if r: tabla = r[1]; metodo = f"lectura por contenido (hoja {r[0]})"
    if tabla is None or tabla.empty:
        print("No se encontraron CUM con precio en este documento."); continue
    res = cruzar(tabla, info.get("forma_de_pago") or "CRÉDITO")
    destino = f"data/salida/analisis_ia_{os.path.splitext(nombre)[0]}.xlsx"
    with pd.ExcelWriter(destino, engine="openpyxl") as xw:
        res.to_excel(xw, sheet_name="CRUCE", index=False)
        pd.DataFrame([{"ARCHIVO": nombre, "METODO": metodo, "TIPO": info.get("tipo"), "RESUMEN_IA": info.get("resumen"),
                       "TOPES": json.dumps(info.get("topes_regulatorios", []), ensure_ascii=False)}]).to_excel(xw, sheet_name="LECTURA", index=False)
    cuenta = res["RESULTADO"].value_counts()
    print(f"Leído con {metodo}: {len(res)} líneas → " + "; ".join(f"{k}: {v}" for k, v in cuenta.items()))
    print("Guardado en", destino)
    resultados_ia.append(res.assign(ARCHIVO=nombre))

if not resultados_ia:
    print(f"\nDeja documentos en {CARPETA_IA} y vuelve a ejecutar esta celda.")
else:
    display(pd.concat(resultados_ia)[["ARCHIVO", "CUM", "DESCRIPCION", "TARIFA", "PMV_CIRCULAR_022", "PRECIO_SISMED", "%_SOBRE_PMV", "%_SOBRE_SISMED", "RESULTADO"]].head(30))

# %% [markdown]
# ## Conclusiones
#
# - La **lectura inteligente** permite consolidar archivos con nombres, hojas y encabezados distintos, y deja trazabilidad de qué se leyó y qué faltó.
# - Cruzar cada CUM con la Circular 022 detecta en segundos las tarifas que superan el precio máximo institucional, el mismo tipo de hallazgo que hoy sanciona la SIC.
# - El seguimiento de PQRS por severidad, canal y ente permite anticipar los riesgos que muestran los casos reales del mercado: pendientes, desabastecimiento y sobreprecios.
# - El asistente de IA del Paso 10 lee documentos sin formato (incluidos PDF y Word) y deja el cruce listo para revisar.
# - Basta con dejar nuevos archivos en `entrada` y volver a ejecutar para actualizar el Excel, el memorando, el informe HTML y el portal.
