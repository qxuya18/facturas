#==============================================================
#VALIACIONES SI TIENE INSTALAADAS LAS LIBRERIAS SI NO EL SCRIPTLAS INSTALARA
#=============================================================
import subprocess
import sys
import importlib.util

def instalar_si_falta(pip_name, import_name):
    """Instala el paquete si no está disponible."""
    if importlib.util.find_spec(import_name) is None:
        print(f"Instalando {pip_name}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])

instalar_si_falta("PyPDF2", "PyPDF2")
instalar_si_falta("pandas", "pandas")
instalar_si_falta("openpyxl", "openpyxl")



#===============================================================
# INICIO DE SCRIPT IMPORTAR LIBRERIAS PARA PODER EJECUTARSE
#===============================================================
import re
import os
import PyPDF2
import pandas as pd


def buscar_dato(patron, texto, flags=re.IGNORECASE):

    resultado = re.search(patron, texto, flags)

    if resultado:
        return resultado.group(1).strip()

    return "No encontrado"


def extraer_factura_desprocesada(pdf_file_path):

    texto = ""

    # Abrir y leer el PDF
    with open(pdf_file_path, "rb") as file:

        lector = PyPDF2.PdfReader(file)

        for pagina in lector.pages:

            texto_pagina = pagina.extract_text()

            if texto_pagina:
                texto += texto_pagina + "\n"

    # Normalizar espacios, manteniendo los saltos de línea
    texto = re.sub(r"[ \t]+", " ", texto)

    print("\n========== TEXTO EXTRAÍDO ==========")
    print(texto)
    print("====================================")

    # ============================
    # DATOS DEL ENCABEZADO
    # ============================

    tipo_factura = buscar_dato(
        r"^\s*(Factura(?:\s+Pequeño\s+Contribuyente)?)",
        texto,
        re.IGNORECASE | re.MULTILINE
    )

    proveedor = buscar_dato(
        r"Factura(?:\s+Pequeño\s+Contribuyente)?"
        r"\s*\d*\s*(.*?)\s+"
        r"N[ÚU]MERO\s+DE\s+AUTORIZACI[ÓO]N\s*:",
        texto,
        re.IGNORECASE | re.DOTALL
    )

    nit_emisor = buscar_dato(
        r"Nit\s+Emisor\s*:\s*([0-9-]+)",
        texto
    )

    serie = buscar_dato(
        r"Serie\s*:\s*([A-Z0-9]+)",
        texto
    )

    numero_dte = buscar_dato(
        r"N[úu]mero\s+de\s+DTE\s*:\s*(\d+)",
        texto
    )

    # ============================
    # DESCRIPCIONES
    # ============================

    descripciones = extraer_descripciones(texto)

    # Unimos todas las descripciones en un solo texto
    descripcion = " | ".join(descripciones)

    # ============================
    # TOTAL
    # ============================

    total = buscar_dato(
        r"TOTALES\s*:\s*[\d,]+\.\d{2}\s+([\d,]+\.\d{2})",
        texto
    )

    datos_factura = {
        "Tipo Factura": tipo_factura,
        "Proveedor": proveedor,
        "Nit Emisor": nit_emisor,
        "Serie": serie,
        "Número DTE": numero_dte,
        "Descripción": descripcion,
        "Total": total
    }

    return datos_factura


def extraer_descripciones(texto):

    descripciones = []

    patron = (
        r"\d+\s*(?:Bien|Servicio)\s*"
        r"[\d,.]+\s*"
        r"(.*?)"
        r"\s+[\d,]+\.\d{2}"
    )

    resultados = re.findall(
        patron,
        texto,
        re.IGNORECASE | re.DOTALL
    )

    for descripcion in resultados:

        descripcion_limpia = re.sub(
            r"\s+",
            " ",
            descripcion
        ).strip()

        descripciones.append(descripcion_limpia)

    return descripciones


    
if __name__ == "__main__":

    carpeta = "sin_procesar"
    facturas = []

    for archivo in os.listdir(carpeta):

        if archivo.lower().endswith(".pdf"):

            ruta_pdf = os.path.join(carpeta, archivo)

            print(f"\nProcesando: {archivo}")

            try:
                datos = extraer_factura_desprocesada(ruta_pdf)

                # Agregamos el nombre del archivo para identificarlo
                datos["Archivo"] = archivo

                facturas.append(datos)

                print("Procesada correctamente.")

            except Exception as error:
                print(f"Error al procesar {archivo}: {error}")

    print("\n=================================")
    print("RESULTADO DE TODAS LAS FACTURAS")
    print("=================================")

    for numero, factura in enumerate(facturas, start=1):

        print(f"\n----- FACTURA {numero} -----")

        for campo, valor in factura.items():
            print(f"{campo}: {valor}")

    print(f"\nTotal de facturas procesadas: {len(facturas)}")


    #Exortacion a un archivo xlsx
    df = pd.DataFrame(facturas)
    df.to_excel("facturas_procesadas.xlsx", index=False)
    print(f"\nArchivo guardado: facturas_procesadas.xlsx")
