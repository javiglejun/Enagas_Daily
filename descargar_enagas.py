import csv
import re
import time
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from urllib.parse import urlencode

import openpyxl
import requests


# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_URL = (
    "https://www.enagas.es/content/enagas/es/gestion-tecnica-sistema/"
    "energy-data/parametros-fisicos/capacidades-tecnicas-flujos-fisicos/"
    "seguimiento-diario-sistema/jcr:content/responsiveGrid/"
    "dailysystemtracking.dailysystemtrackingdto.xls"
)

CARPETA_DATOS = Path("data")

HISTORICO = CARPETA_DATOS / "historico_sistema_gasista.csv"

# El histórico comenzará aquí
FECHA_INICIO_HISTORICO = date(2025, 1, 1)

COLUMNAS = [
    "Fecha",
    "Demanda Nacional",
    "Convencional",
    "Sector Eléctrico",
    "Total Salidas",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) "
        "AppleWebKit/537.36 "
        "Chrome/130.0 Safari/537.36"
    ),
    "Accept": (
        "application/vnd.ms-excel,"
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet,"
        "application/octet-stream,*/*"
    ),
    "Accept-Language": "es-ES,es;q=0.9",
    "Referer": "https://www.enagas.es/",
}


# ============================================================
# CONEXIÓN CON ENAGÁS
# ============================================================

def crear_sesion():

    sesion = requests.Session()
    sesion.headers.update(HEADERS)

    return sesion


def construir_url(fecha):

    fecha_texto = fecha.strftime("%d/%m/%Y")

    parametros = urlencode({
        "date": fecha_texto
    })

    return f"{BASE_URL}?{parametros}"


def es_excel_valido(contenido):

    if len(contenido) < 1000:
        return False

    # Los archivos descargados tienen extensión .xls,
    # pero realmente son archivos XLSX (ZIP)
    return contenido.startswith(b"PK")


def descargar_excel(sesion, fecha):

    url = construir_url(fecha)

    fecha_texto = fecha.strftime("%d/%m/%Y")

    print(f"Descargando {fecha_texto}")

    try:

        respuesta = sesion.get(
            url,
            timeout=60
        )

    except requests.RequestException as error:

        print(
            f"Error de conexión para {fecha_texto}: "
            f"{error}"
        )

        return None

    if respuesta.status_code != 200:

        print(
            f"No disponible {fecha_texto}. "
            f"HTTP {respuesta.status_code}"
        )

        return None

    if not es_excel_valido(respuesta.content):

        print(
            f"No existe un Excel válido para "
            f"{fecha_texto}"
        )

        return None

    return respuesta.content


# ============================================================
# EXTRACCIÓN DE DATOS DEL EXCEL
# ============================================================

def buscar_fecha_en_excel(hoja, fecha_solicitada):

    for fila in hoja.iter_rows():

        for celda in fila:

            if not isinstance(celda.value, str):
                continue

            texto = celda.value.strip()

            if (
                "Día de gas" in texto
                or "Dia de gas" in texto
            ):

                coincidencia = re.search(
                    r"\d{2}/\d{2}/\d{4}",
                    texto
                )

                if coincidencia:

                    fecha_excel = datetime.strptime(
                        coincidencia.group(),
                        "%d/%m/%Y"
                    )

                    return fecha_excel.strftime(
                        "%Y-%m-%d"
                    )

    return fecha_solicitada.strftime("%Y-%m-%d")


def convertir_numero(valor, nombre, celda):

    if valor is None:

        raise ValueError(
            f"La celda {celda} de "
            f"'{nombre}' está vacía."
        )

    if isinstance(valor, (int, float)):
        return valor

    texto = str(valor).strip()

    texto = texto.replace(" ", "")

    if "," in texto and "." in texto:

        texto = texto.replace(".", "")
        texto = texto.replace(",", ".")

    elif "," in texto:

        texto = texto.replace(",", ".")

    try:

        return float(texto)

    except ValueError as error:

        raise ValueError(
            f"El valor de {nombre} "
            f"en {celda} no es numérico: {valor}"
        ) from error


def extraer_datos(contenido_excel, fecha_solicitada):

    libro = openpyxl.load_workbook(
        BytesIO(contenido_excel),
        data_only=True,
        read_only=True
    )

    if "Resumen" not in libro.sheetnames:

        libro.close()

        raise ValueError(
            "El Excel no contiene la hoja 'Resumen'."
        )

    hoja = libro["Resumen"]

    datos = {

        "Fecha": buscar_fecha_en_excel(
            hoja,
            fecha_solicitada
        ),

        "Demanda Nacional": convertir_numero(
            hoja["E8"].value,
            "Demanda Nacional",
            "E8"
        ),

       
