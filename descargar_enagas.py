import csv
import re
from datetime import datetime, timedelta
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

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": (
        "application/vnd.ms-excel,"
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,"
        "application/octet-stream,*/*"
    ),
}


def descargar_excel(fecha):
    """
    Descarga el Excel correspondiente a la fecha indicada.
    Devuelve la ruta del fichero si la descarga es correcta.
    """

    fecha_texto = fecha.strftime("%d/%m/%Y")

    url = f"{BASE_URL}?{urlencode({'date': fecha_texto})}"

    print(f"Intentando descargar el fichero del {fecha_texto}")
    print(f"URL: {url}")

    response = requests.get(
        url,
        headers=HEADERS,
        timeout=60
    )

    if response.status_code != 200:
        print(f"Fichero no disponible. HTTP {response.status_code}")
        return None

    if len(response.content) < 1000:
        print("La respuesta recibida no parece un Excel válido.")
        return None

    CARPETA_DATOS.mkdir(exist_ok=True)

    nombre = (
        f"Enagas_Sistema_Gasista_"
        f"{fecha.strftime('%Y-%m-%d')}.xls"
    )

    ruta = CARPETA_DATOS / nombre

    ruta.write_bytes(response.content)

    print(f"Fichero descargado correctamente: {ruta}")

    return ruta


def obtener_fecha_excel(hoja, fecha_descarga):
    """
    Obtiene la fecha del día de gas desde el propio Excel.

    En el fichero analizado aparece como:
    'Día de gas: 09/11/2026'

    Si no se localiza, utiliza la fecha empleada
    para descargar el fichero.
    """

    for fila in hoja.iter_rows():
        for celda in fila:

            if not isinstance(celda.value, str):
                continue

            texto = celda.value.strip()

            if "Día de gas" in texto or "Dia de gas" in texto:

                coincidencia = re.search(
                    r"\d{2}/\d{2}/\d{4}",
                    texto
                )

                if coincidencia:
                    fecha = datetime.strptime(
                        coincidencia.group(),
                        "%d/%m/%Y"
                    )

                    return fecha.strftime("%Y-%m-%d")

    return fecha_descarga.strftime("%Y-%m-%d")


def convertir_numero(valor, nombre):
    """
    Comprueba que el valor extraído sea numérico.
    """

    if valor is None:
        raise ValueError(
            f"La celda correspondiente a '{nombre}' está vacía."
        )

    if isinstance(valor, (int, float)):
        return valor

    texto = str(valor).strip().replace(",", ".")

    try:
        return float(texto)

    except ValueError as error:
        raise ValueError(
            f"El valor extraído para '{nombre}' "
            f"no es numérico: {valor}"
        ) from error


def extraer_datos(ruta_excel, fecha_descarga):
    """
    Extrae los datos de la hoja Resumen.
    """

    libro = openpyxl.load_workbook(
        ruta_excel,
        data_only=True
    )

    if "Resumen" not in libro.sheetnames:
        raise ValueError(
            "No se encuentra la hoja 'Resumen' en el Excel."
        )

    hoja = libro["Resumen"]

    datos = {
        "Fecha": obtener_fecha_excel(
            hoja,
            fecha_descarga
        ),
        "Demanda Nacional": convertir_numero(
            hoja["E8"].value,
            "Demanda Nacional"
        ),
        "Convencional": convertir_numero(
            hoja["H8"].value,
            "Convencional"
        ),
        "Sector Eléctrico": convertir_numero(
            hoja["H10"].value,
            "Sector Eléctrico"
        ),
        "Total Salidas": convertir_numero(
            hoja["E21"].value,
            "Total Salidas"
        ),
    }

    libro.close()

    print("Datos extraídos:")
    print(datos)

    return datos


def leer_historico():
    """
    Lee el histórico existente.
    """

    if not HISTORICO.exists():
        return []

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

        return list(lector)


def guardar_historico(datos):
    """
    Añade una nueva fecha al histórico.

    Si la fecha ya existe, sustituye esa fila.
    De esta forma no se generan duplicados.
    """

    columnas = [
        "Fecha",
        "Demanda Nacional",
        "Convencional",
        "Sector Eléctrico",
        "Total Salidas",
    ]

    filas = leer_historico()

    filas = [
        fila for fila in filas
        if fila.get("Fecha") != datos["Fecha"]
    ]

    filas.append({
        columna: datos[columna]
        for columna in columnas
    })

    filas.sort(
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
            fieldnames=columnas,
            delimiter=";"
        )

        escritor.writeheader()
        escritor.writerows(filas)

    print(f"Histórico actualizado correctamente: {HISTORICO}")
    print(f"Número total de fechas: {len(filas)}")


def main():
    """
    Busca el fichero más reciente disponible,
    extrae los datos y actualiza el histórico.
    """

    hoy = datetime.now()

    for dias_atras in range(7):

        fecha_descarga = hoy - timedelta(
            days=dias_atras
        )

        ruta_excel = descargar_excel(
            fecha_descarga
        )

        if ruta_excel is not None:

            datos = extraer_datos(
                ruta_excel,
                fecha_descarga
            )

            guardar_historico(datos)

            return

    raise Exception(
        "No se ha podido descargar ningún fichero "
        "de los últimos 7 días."
    )


if __name__ == "__main__":
    main()
