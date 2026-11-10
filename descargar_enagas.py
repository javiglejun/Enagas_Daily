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
    "Total Regasificación",
    "Barcelona",
    "Cartagena",
    "Huelva",
    "BBG",
    "Sagunto",
    "Reganosa",
    "El Musel"
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

        "Convencional": convertir_numero(
            hoja["H8"].value,
            "Convencional",
            "H8"
        ),

        "Sector Eléctrico": convertir_numero(
            hoja["H10"].value,
            "Sector Eléctrico",
            "H10"
        ),

        "Total Salidas": convertir_numero(
            hoja["E21"].value,
            "Total Salidas",
            "E21"
        ),

        "Total Regasificación": convertir_numero(
            hoja["E31"].value,
            "Total Regasificación",
            "E31"
        ),

        "Barcelona": convertir_numero(
            hoja["F31"].value,
            "Barcelona",
            "F31"
        ),

         "Cartagena": convertir_numero(
            hoja["G31"].value,
            "Cartagena",
            "G31"
        ),

         "Huelva": convertir_numero(
            hoja["G31"].value,
            "Huelva",
            "G31"
        ),

         "BBG": convertir_numero(
            hoja["I31"].value,
            "BBG",
            "I31"
        ),

         "Sagunto": convertir_numero(
            hoja["J31"].value,
            "Sagunto",
            "J31"
        ),

         "Reganosa": convertir_numero(
            hoja["K31"].value,
            "Reganosa",
            "K31"
         ),

           "El Musel": convertir_numero(
            hoja["L31"].value,
            "El Musel",
            "L31"
         ),
    }

    libro.close()

    return datos


# ============================================================
# GESTIÓN DEL HISTÓRICO
# ============================================================

def leer_historico():

    registros = {}

    if not HISTORICO.exists():

        print(
            "No existe todavía histórico."
        )

        return registros

    with open(
        HISTORICO,
        "r",
        encoding="utf-8-sig",
        newline=""
    ) as archivo:

        lector = csv.DictReader(
            archivo,
            delimiter=";"
        )

        for fila in lector:

            fecha = fila.get("Fecha")

            if fecha:

                registros[fecha] = fila

    print(
        f"Registros existentes en el histórico: "
        f"{len(registros)}"
    )

    return registros


def guardar_historico(registros):

    CARPETA_DATOS.mkdir(
        parents=True,
        exist_ok=True
    )

    filas_ordenadas = sorted(
        registros.values(),
        key=lambda fila: fila["Fecha"]
    )

    with open(
        HISTORICO,
        "w",
        encoding="utf-8-sig",
        newline=""
    ) as archivo:

        escritor = csv.DictWriter(
            archivo,
            fieldnames=COLUMNAS,
            delimiter=";"
        )

        escritor.writeheader()

        escritor.writerows(
            filas_ordenadas
        )

    print()
    print(
        "========================================"
    )
    print(
        "HISTÓRICO ACTUALIZADO"
    )
    print(
        "========================================"
    )

    print(
        f"Archivo: {HISTORICO}"
    )

    print(
        f"Registros totales: "
        f"{len(filas_ordenadas)}"
    )


# ============================================================
# DETECCIÓN DE FECHAS QUE FALTAN
# ============================================================

def obtener_fechas_a_descargar(registros):

    """
    Busca todas las fechas que faltan entre
    el 01/01/2025 y hoy.

    De esta forma:

    - Si solamente existe 09/11/2026,
      descargará todo lo anterior.

    - Si ya existe todo el histórico,
      solamente descargará el nuevo día.

    - Si falta algún día intermedio,
      también intentará recuperarlo.
    """

    hoy = datetime.now().date()

    fechas_a_descargar = []

    fecha_actual = FECHA_INICIO_HISTORICO

    while fecha_actual <= hoy:

        fecha_texto = fecha_actual.strftime(
            "%Y-%m-%d"
        )

        if fecha_texto not in registros:

            fechas_a_descargar.append(
                fecha_actual
            )

        fecha_actual += timedelta(days=1)

    return fechas_a_descargar


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

def main():

    CARPETA_DATOS.mkdir(
        parents=True,
        exist_ok=True
    )

    print()
    print(
        "========================================"
    )
    print(
        "ACTUALIZACIÓN SISTEMA GASISTA ENAGÁS"
    )
    print(
        "========================================"
    )
    print()

    # Leer histórico existente
    registros = leer_historico()

    # Determinar qué fechas faltan realmente
    fechas_a_descargar = obtener_fechas_a_descargar(
        registros
    )

    print()

    if not fechas_a_descargar:

        print(
            "El histórico ya está completamente "
            "actualizado."
        )

        return

    print(
        f"Fechas pendientes de comprobar: "
        f"{len(fechas_a_descargar)}"
    )

    print(
        f"Primera fecha pendiente: "
        f"{fechas_a_descargar[0].strftime('%d/%m/%Y')}"
    )

    print(
        f"Última fecha pendiente: "
        f"{fechas_a_descargar[-1].strftime('%d/%m/%Y')}"
    )

    print()

    sesion = crear_sesion()

    nuevos_registros = 0
    errores = 0

    for numero, fecha in enumerate(
        fechas_a_descargar,
        start=1
    ):

        print()
        print(
            f"[{numero}/{len(fechas_a_descargar)}]"
        )

        contenido_excel = descargar_excel(
            sesion,
            fecha
        )

        if contenido_excel is None:

            errores += 1

            continue

        try:

            datos = extraer_datos(
                contenido_excel,
                fecha
            )

            registros[
                datos["Fecha"]
            ] = datos

            nuevos_registros += 1

            print(
                f"OK | {datos['Fecha']} | "
                f"Demanda Nacional: "
                f"{datos['Demanda Nacional']} | "
                f"Convencional: "
                f"{datos['Convencional']} | "
                f"Sector Eléctrico: "
                f"{datos['Sector Eléctrico']} | "
                f"Total Salidas: "
                f"{datos['Total Salidas']} | "
                f"Total Regasificación: "
                f"{datos['Total Regasificación']} | "
                f"Barcelona: "
                f"{datos['Barcelona']} | "
                f"Cartagena: "
                f"{datos['Cartagena']} | "
                f"Huelva: "
                f"{datos['Huelva']} | "
                f"BBG: "
                f"{datos['BBG']} | "
                f"Sagunto: "
                f"{datos['Sagunto']} | "
                f"Reganosa: "
                f"{datos['Reganosa']} | "
                f"El Musel: "
                f"{datos['El Musel']}"
            )

        except Exception as error:

            errores += 1

            print(
                f"ERROR procesando "
                f"{fecha.strftime('%d/%m/%Y')}: "
                f"{error}"
            )

        # Pequeña pausa para no saturar
        # el servidor de Enagás.
        time.sleep(0.25)

    # Guardamos el CSV una vez terminada
    # la actualización
    guardar_historico(
        registros
    )

    print()
    print(
        f"Nuevos registros incorporados: "
        f"{nuevos_registros}"
    )

    print(
        f"Fechas no disponibles/error: "
        f"{errores}"
    )


if __name__ == "__main__":
    main()
