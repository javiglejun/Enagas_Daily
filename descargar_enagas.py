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
# CONFIGURACIÓN GENERAL
# ============================================================

BASE_URL = (
    "https://www.enagas.es/content/enagas/es/gestion-tecnica-sistema/"
    "energy-data/parametros-fisicos/capacidades-tecnicas-flujos-fisicos/"
    "seguimiento-diario-sistema/jcr:content/responsiveGrid/"
    "dailysystemtracking.dailysystemtrackingdto.xls"
)

CARPETA_DATOS = Path("data")

HISTORICO = CARPETA_DATOS / "historico_sistema_gasista.csv"

FECHA_INICIO_HISTORICO = date(2025, 1, 1)

COLUMNAS = [
    "Fecha",
    "Demanda Nacional",
    "Convencional",
    "Sector Eléctrico",
    "Total Salidas",
    "Total Entradas",
    "Gas regasificado",
    "CCII",
    "AASS",
    "Yacimientos",
    "Transporte biometano",
    "Distribución biometano",
]

CELDAS = {
    "Demanda Nacional": "E8",
    "Convencional": "H8",
    "Sector Eléctrico": "H10",
    "Total Salidas": "E21",
    "Total Entradas": "E47",
    "Gas regasificado": "F47",
    "CCII": "G47",
    "AASS": "H47",
    "Yacimientos": "I47",
    "Transporte biometano": "J47",
    "Distribución biometano": "K47",
}

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

TIMEOUT_SEGUNDOS = 60

MAXIMO_INTENTOS = 3

PAUSA_ENTRE_FECHAS = 0.25


# ============================================================
# CONEXIÓN CON ENAGÁS
# ============================================================

def crear_sesion():
    """
    Crea una sesión HTTP que se reutiliza durante
    toda la ejecución.
    """

    sesion = requests.Session()
    sesion.headers.update(HEADERS)

    return sesion


def construir_url(fecha):
    """
    Construye la URL del Excel correspondiente
    a la fecha indicada.
    """

    fecha_texto = fecha.strftime("%d/%m/%Y")

    parametros = urlencode({
        "date": fecha_texto
    })

    return f"{BASE_URL}?{parametros}"


def es_excel_valido(contenido):
    """
    Comprueba que la respuesta descargada tenga
    realmente formato Excel.

    Los ficheros publicados por Enagás tienen una URL
    terminada en .xls, pero internamente son archivos
    XLSX, que comienzan con la firma ZIP 'PK'.
    """

    if contenido is None:
        return False

    if len(contenido) < 1000:
        return False

    return contenido.startswith(b"PK")


def descargar_excel(sesion, fecha):
    """
    Descarga en memoria el Excel de una fecha.

    Realiza hasta tres intentos si se produce
    un error temporal de conexión.
    """

    url = construir_url(fecha)

    fecha_texto = fecha.strftime("%d/%m/%Y")

    print(f"Descargando {fecha_texto}")

    for intento in range(1, MAXIMO_INTENTOS + 1):

        try:

            respuesta = sesion.get(
                url,
                timeout=TIMEOUT_SEGUNDOS
            )

        except requests.RequestException as error:

            print(
                f"Error de conexión en el intento "
                f"{intento}/{MAXIMO_INTENTOS}: {error}"
            )

            if intento < MAXIMO_INTENTOS:
                time.sleep(2 * intento)

            continue

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

    print(
        f"No se ha podido descargar {fecha_texto} "
        f"después de {MAXIMO_INTENTOS} intentos."
    )

    return None


# ============================================================
# EXTRACCIÓN DE DATOS
# ============================================================

def buscar_fecha_en_excel(hoja, fecha_solicitada):
    """
    Busca la fecha del Día de gas dentro del Excel.

    Si no se localiza, utiliza la fecha solicitada
    en la descarga.
    """

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
    """
    Convierte el contenido de una celda en número.

    Admite valores numéricos y textos con coma
    o punto decimal.
    """

    if valor is None:

        raise ValueError(
            f"La celda {celda}, correspondiente a "
            f"'{nombre}', está vacía."
        )

    if isinstance(valor, bool):

        raise ValueError(
            f"La celda {celda}, correspondiente a "
            f"'{nombre}', contiene un valor booleano."
        )

    if isinstance(valor, (int, float)):
        return valor

    texto = str(valor).strip()

    texto = texto.replace(" ", "")

    if "," in texto and "." in texto:

        ultima_coma = texto.rfind(",")
        ultimo_punto = texto.rfind(".")

        if ultima_coma > ultimo_punto:

            texto = texto.replace(".", "")
            texto = texto.replace(",", ".")

        else:

            texto = texto.replace(",", "")

    elif "," in texto:

        texto = texto.replace(",", ".")

    try:

        return float(texto)

    except ValueError as error:

        raise ValueError(
            f"El valor de '{nombre}' en la celda "
            f"{celda} no es numérico: {valor}"
        ) from error


def extraer_datos(contenido_excel, fecha_solicitada):
    """
    Abre el Excel en memoria y extrae todas
    las variables configuradas en CELDAS.
    """

    libro = openpyxl.load_workbook(
        BytesIO(contenido_excel),
        data_only=True,
        read_only=True
    )

    try:

        if "Resumen" not in libro.sheetnames:

            raise ValueError(
                "El Excel no contiene la hoja 'Resumen'."
            )

        hoja = libro["Resumen"]

        datos = {
            "Fecha": buscar_fecha_en_excel(
                hoja,
                fecha_solicitada
            )
        }

        for nombre, celda in CELDAS.items():

            datos[nombre] = convertir_numero(
                hoja[celda].value,
                nombre,
                celda
            )

        return datos

    finally:

        libro.close()


# ============================================================
# LECTURA Y VALIDACIÓN DEL HISTÓRICO
# ============================================================

def leer_historico():
    """
    Lee el CSV histórico existente.

    Las filas se almacenan en un diccionario cuya
    clave es la fecha, para evitar duplicados.
    """

    registros = {}

    if not HISTORICO.exists():

        print("No existe todavía el histórico.")

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

        columnas_existentes = lector.fieldnames or []

        columnas_faltantes = [
            columna
            for columna in COLUMNAS
            if columna not in columnas_existentes
        ]

        if columnas_faltantes:

            print(
                "El histórico existente no contiene "
                "todas las nuevas columnas."
            )

            print(
                "Columnas que deben completarse: "
                + ", ".join(columnas_faltantes)
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


def fila_completa(fila):
    """
    Comprueba si una fila contiene todas las columnas
    y todos los valores requeridos.

    Las filas del histórico antiguo no tienen las nuevas
    variables, por lo que se marcarán como incompletas
    y se volverán a descargar una sola vez.
    """

    if not fila:
        return False

    for columna in COLUMNAS:

        if columna not in fila:
            return False

        valor = fila.get(columna)

        if valor is None:
            return False

        if str(valor).strip() == "":
            return False

    return True


def obtener_fechas_a_descargar(registros):
    """
    Devuelve todas las fechas ausentes o incompletas
    entre el 01/01/2025 y hoy.

    Primera ejecución con las nuevas columnas:
    vuelve a procesar todo el histórico.

    Ejecuciones posteriores:
    procesa únicamente las fechas nuevas o aquellas
    que hayan quedado incompletas.
    """

    hoy = datetime.now().date()

    fechas_a_descargar = []

    fecha_actual = FECHA_INICIO_HISTORICO

    while fecha_actual <= hoy:

        fecha_texto = fecha_actual.strftime(
            "%Y-%m-%d"
        )

        fila = registros.get(fecha_texto)

        if not fila_completa(fila):

            fechas_a_descargar.append(
                fecha_actual
            )

        fecha_actual += timedelta(days=1)

    return fechas_a_descargar


# ============================================================
# GUARDADO DEL CSV
# ============================================================

def preparar_fila_para_guardar(fila):
    """
    Devuelve únicamente las columnas definitivas
    y en el orden establecido.
    """

    return {
        columna: fila.get(columna, "")
        for columna in COLUMNAS
    }


def guardar_historico(registros):
    """
    Guarda el histórico ordenado cronológicamente.

    El archivo utiliza punto y coma como delimitador
    y UTF-8 con BOM para facilitar su lectura en Excel
    y Power BI.
    """

    CARPETA_DATOS.mkdir(
        parents=True,
        exist_ok=True
    )

    filas_ordenadas = sorted(
        registros.values(),
        key=lambda fila: fila["Fecha"]
    )

    filas_ordenadas = [
        preparar_fila_para_guardar(fila)
        for fila in filas_ordenadas
    ]

    archivo_temporal = HISTORICO.with_suffix(
        ".csv.tmp"
    )

    with open(
        archivo_temporal,
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
        escritor.writerows(filas_ordenadas)

    archivo_temporal.replace(HISTORICO)

    print()
    print("========================================")
    print("HISTÓRICO ACTUALIZADO")
    print("========================================")
    print(f"Archivo: {HISTORICO}")
    print(
        f"Registros totales: "
        f"{len(filas_ordenadas)}"
    )


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

def main():

    CARPETA_DATOS.mkdir(
        parents=True,
        exist_ok=True
    )

    print()
    print("========================================")
    print("ACTUALIZACIÓN SISTEMA GASISTA ENAGÁS")
    print("========================================")
    print()

    registros = leer_historico()

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
        f"Fechas pendientes de procesar: "
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
    fechas_sin_fichero = []

    total_fechas = len(fechas_a_descargar)

    for numero, fecha in enumerate(
        fechas_a_descargar,
        start=1
    ):

        print()
        print(f"[{numero}/{total_fechas}]")

        contenido_excel = descargar_excel(
            sesion,
            fecha
        )

        if contenido_excel is None:

            errores += 1

            fechas_sin_fichero.append(
                fecha.strftime("%Y-%m-%d")
            )

            continue

        try:

            datos = extraer_datos(
                contenido_excel,
                fecha
            )

            registros[datos["Fecha"]] = datos

            nuevos_registros += 1

            print(
                f"OK | {datos['Fecha']} | "
                f"Demanda Nacional: "
                f"{datos['Demanda Nacional']} | "
                f"Total Salidas: "
                f"{datos['Total Salidas']} | "
                f"Total Entradas: "
                f"{datos['Total Entradas']}"
            )

        except Exception as error:

            errores += 1

            fechas_sin_fichero.append(
                fecha.strftime("%Y-%m-%d")
            )

            print(
                f"ERROR procesando "
                f"{fecha.strftime('%d/%m/%Y')}: "
                f"{error}"
            )

        time.sleep(PAUSA_ENTRE_FECHAS)

    guardar_historico(
        registros
    )

    print()
    print(
        f"Registros incorporados o actualizados: "
        f"{nuevos_registros}"
    )

    print(
        f"Fechas no disponibles o con error: "
        f"{errores}"
    )

    if fechas_sin_fichero:

        print()
        print(
            "Fechas que se volverán a intentar "
            "en la próxima ejecución:"
        )

        for fecha in fechas_sin_fichero:
            print(f"- {fecha}")


if __name__ == "__main__":
    main()
