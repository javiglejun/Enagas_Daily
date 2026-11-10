import requests
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlencode


BASE_URL = (
    "https://www.enagas.es/content/enagas/es/gestion-tecnica-sistema/"
    "energy-data/parametros-fisicos/capacidades-tecnicas-flujos-fisicos/"
    "seguimiento-diario-sistema/jcr:content/responsiveGrid/"
    "dailysystemtracking.dailysystemtrackingdto.xls"
)

CARPETA = Path("data")

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/vnd.ms-excel,application/octet-stream,*/*",
}


def descargar_excel(fecha):

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
        return False

    if len(response.content) < 1000:
        print("La respuesta recibida no parece un Excel válido.")
        return False

    CARPETA.mkdir(exist_ok=True)

    nombre = f"Enagas_Sistema_Gasista_{fecha.strftime('%Y-%m-%d')}.xls"

    ruta = CARPETA / nombre

    ruta.write_bytes(response.content)

    print(f"Fichero descargado correctamente: {ruta}")

    return True


def main():

    hoy = datetime.now()

    # Probamos primero con hoy.
    # Si Enagás todavía no lo ha publicado,
    # buscamos hacia atrás hasta 7 días.

    for dias_atras in range(7):

        fecha = hoy - timedelta(days=dias_atras)

        if descargar_excel(fecha):
            return

    raise Exception(
        "No se ha podido descargar ningún fichero "
        "de los últimos 7 días."
    )


if __name__ == "__main__":
    main()
