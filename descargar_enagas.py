import csv
import re
import time
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from urllib.parse import urlencode

import openpyxl
import requests


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


def crear_sesion():
    """
    Crea una sesión para reutilizar la conexión
    durante toda la ejecución.
    """

    sesion = requests.Session()
    sesion.headers.update(HEADERS)

    return sesion


def construir_url(fecha):
    """
    Construye la URL del fichero correspondiente
    a la fecha indicada.
    """

    fecha_texto = fecha.strftime("%d/%m/%Y")

    parametros = urlencode({
        "date": fecha_texto
    })

    return f"{BASE_URL}?{parametros}"


def es_excel_valido(contenido):
    """
    Comprueba que la descarga sea realmente un Excel.

    Los ficheros de Enagás terminan en .xls,
    pero internamente tienen formato XLSX.
    Los XLSX son archivos ZIP y empiezan por PK.
    """

    if len(contenido) < 1000:
        return False

    return contenido.startswith(b"PK")


def descargar_excel(sesion, fecha):
    """
    Descarga en memoria el Excel de una fecha.

    No guarda cientos de Excel en el repositorio.
    Devuelve el contenido descargado.
    """

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
            f"La respuesta del {fecha_texto} "
            "no es un Excel válido."
        )

        return None

    return respuesta.content


def buscar_fecha_en_excel(hoja, fecha_solicitada):
    """
    Obtiene el Día de gas desde el propio Excel.

    Si no se localiza, utiliza la fecha solicitada.
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
    Comprueba que el dato obtenido sea numérico.
    """

    if valor is None:

        raise ValueError(
            f"La celda {celda}, correspondiente a "
            f"'{nombre}', está vacía."
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
            f"El valor de {nombre} en la celda "
            f"{celda} no es numérico: {valor}"
        ) from error


def extraer_datos(contenido_excel, fecha_solicitada):
    """
    Abre el Excel en memoria y extrae las cuatro celdas.
    """

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
    }

    libro.close()

    return datos


def leer_historico():
    """
    Lee el CSV existente.

    Devuelve un diccionario cuya clave es la fecha.
    Esto permite evitar duplicados.
    """

    registros = {}

    if not HISTORICO.exists():
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

    return registros


def guardar_historico(registros):
    """
    Guarda el histórico completo ordenado por fecha.
    """

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
        f"Histórico guardado: {HISTORICO}"
    )
    print(
        f"Número total de registros: "
        f"{len(filas_ordenadas)}"
    )


def obtener_ultima_fecha(registros):
    """
    Devuelve la fecha más reciente que ya existe
    en el histórico.
    """

    if not registros:
        return None

    ultima_fecha_texto = max(
        registros.keys()
    )

    return datetime.strptime(
        ultima_fecha_texto,
        "%Y-%m-%d"
    ).date()


def determinar_fecha_inicial(registros):
    """
    Primera ejecución:
    empieza el 01/01/2025.

    Ejecuciones posteriores:
    empieza el día siguiente a la última fecha
    ya incluida en el CSV.
    """

    ultima_fecha = obtener_ultima_fecha(
        registros
    )

    if ultima_fecha is None:

        print(
            "No existe histórico previo."
        )
        print(
            "Se realizará la carga inicial "
            "desde el 01/01/2025."
        )

        return FECHA_INICIO_HISTORICO

    fecha_inicial = ultima_fecha + timedelta(
        days=1
    )

    print(
        f"Última fecha del histórico: "
        f"{ultima_fecha.strftime('%d/%m/%Y')}"
    )

    print(
        f"La actualización comenzará en: "
        f"{fecha_inicial.strftime('%d/%m/%Y')}"
    )

    return fecha_inicial


def main():

    CARPETA_DATOS.mkdir(
        parents=True,
        exist_ok=True
    )

    registros = leer_historico()

    fecha_actual = determinar_fecha_inicial(
        registros
    )

    hoy = datetime.now().date()

    if fecha_actual > hoy:

        print(
            "El histórico ya está actualizado."
        )

        return

    sesion = crear_sesion()

    nuevos_registros = 0
    fechas_no_disponibles = []

    while fecha_actual <= hoy:

        contenido_excel = descargar_excel(
            sesion,
            fecha_actual
        )

        if contenido_excel is None:

            fechas_no_disponibles.append(
                fecha_actual.strftime("%Y-%m-%d")
            )

        else:

            try:

                datos = extraer_datos(
                    contenido_excel,
                    fecha_actual
                )

                registros[datos["Fecha"]] = datos

                nuevos_registros += 1

                print(
                    f"Datos añadidos: "
                    f"{datos['Fecha']} | "
                    f"Demanda Nacional: "
                    f"{datos['Demanda Nacional']}"
                )

            except Exception as error:

                print(
                    f"Error procesando "
                    f"{fecha_actual.strftime('%d/%m/%Y')}: "
                    f"{error}"
                )

                fechas_no_disponibles.append(
                    fecha_actual.strftime("%Y-%m-%d")
                )

        fecha_actual += timedelta(days=1)

        # Pequeña pausa para no hacer todas las peticiones
        # al servidor de Enagás de forma instantánea.
        time.sleep(0.25)

    guardar_historico(registros)

    print()
    print(
        f"Nuevos registros incorporados: "
        f"{nuevos_registros}"
    )

    if fechas_no_disponibles:

        print()
        print(
            "Fechas sin fichero disponible "
            "o con error:"
        )

        for fecha in fechas_no_disponibles:
            print(f"- {fecha}")


if __name__ == "__main__":
    main()
