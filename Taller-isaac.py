"""
Taller: Automatización de Flujos de Datos con Python
De Excel (openpyxl) a Consola (tabulate) e Informe Ejecutivo (ReportLab)

Autor: Isaac García
Curso: Nuevas Tecnologías - CESDE

Qué hace este programa:
  1. Lee ventas_tecnologia_v2.xlsx con openpyxl.
  2. Muestra los datos en consola con tabulate.
  3. Calcula el precio promedio.
  4. Filtra por precio y busca por ID.
  5. Exporta un informe ejecutivo en PDF con ReportLab.
  6. Todo se controla desde un menú interactivo (while True).
"""

import os
from xml.sax.saxutils import escape

import openpyxl
from tabulate import tabulate
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


# CONSTANTES (si algo cambia, se cambia en un solo lugar)

CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(CARPETA, "ventas_tecnologia_v2.xlsx")
ARCHIVO_PDF = os.path.join(CARPETA, "reporte_ejecutivo_ventas.pdf")
PRECIO_MINIMO = 100

ENCABEZADOS = ["ID", "Producto", "Categoría", "Precio", "Ventas (Unid)"]
ALINEACION = ("center", "left", "center", "right", "right")



# 1. LECTURA DEL EXCEL (openpyxl)

def cargar_datos(ruta):
    "Lee el Excel y devuelve una lista de listas (una por producto)."
    libro = openpyxl.load_workbook(ruta, data_only=True)
    hoja = libro.active  # primera pestaña

    datos = []
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is None:  # ignora filas vacías
            continue
        datos.append(list(fila))
    return datos



# 2. LÓGICA DE NEGOCIO (filtros, métricas, búsqueda)

def filtrar_por_precio(datos, minimo=PRECIO_MINIMO):
    "Reto 1: productos cuyo precio es MAYOR que el mínimo (estrictamente)."
    return [fila for fila in datos if fila[3] > minimo]


def calcular_metricas(datos):
    "Calcula exclusivamente el precio promedio."
    if not datos:
        return 0.0
    precio_promedio = sum(fila[3] for fila in datos) / len(datos)
    return precio_promedio


def buscar_por_id(datos, id_buscado):
    "Reto 3: devuelve la fila con ese ID o None si no existe."
    for fila in datos:
        if fila[0] == id_buscado:
            return fila
    return None



# 3. SALIDA EN CONSOLA (tabulate)

def mostrar_tabla(datos, titulo, formato="fancy_grid"):
    print(f"\n{titulo}")
    print(tabulate(datos, headers=ENCABEZADOS, tablefmt=formato, colalign=ALINEACION))


def mostrar_metricas(datos):
    precio_promedio = calcular_metricas(datos)
    print("\nMÉTRICAS DE RESUMEN")
    print(f"precio promedio            : {precio_promedio:,.0f}")



# 4. INFORME PDF (ReportLab)

def generar_pdf(datos, ruta_pdf, subtitulo="Catálogo completo"):
    "Construye el PDF apilando elementos en la lista 'story'."
    doc = SimpleDocTemplate(ruta_pdf, pagesize=letter)
    estilos = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        "EstiloTitulo", parent=estilos["Title"],
        fontName="Helvetica-Bold", fontSize=18,
        textColor=colors.HexColor("#1A365D"), spaceAfter=10,
    )
    estilo_cuerpo = ParagraphStyle("EstiloCuerpo", parent=estilos["Normal"], fontSize=9)
    estilo_encabezado = ParagraphStyle(
        "EstiloEncabezado", parent=estilos["Normal"],
        fontName="Helvetica-Bold", fontSize=10,
        textColor=colors.whitesmoke, alignment=1,
    )

    precio_promedio = calcular_metricas(datos)

    story = [
        Paragraph("Informe Ejecutivo de Ventas Tecno", estilo_titulo),
        Paragraph(f"<b>Reporte:</b> {escape(subtitulo)}", estilo_cuerpo),
        Paragraph(f"<b>Total de Productos Procesados:</b> {len(datos)}", estilo_cuerpo),
        Spacer(1, 12),
    ]

    estilo_c = ParagraphStyle("C", parent=estilo_cuerpo, alignment=1)
    estilo_l = ParagraphStyle("L", parent=estilo_cuerpo, alignment=0)
    estilo_r = ParagraphStyle("R", parent=estilo_cuerpo, alignment=2)

    encabezados_pdf = [Paragraph(f"<b>{escape(h)}</b>", estilo_encabezado) for h in ENCABEZADOS]
    filas_pdf = [
        [
            Paragraph(str(f[0]), estilo_c),
            Paragraph(escape(str(f[1])), estilo_l),
            Paragraph(escape(str(f[2])), estilo_c),
            Paragraph(f"{f[3]:,.0f}", estilo_r),
            Paragraph(str(f[4]), estilo_r),
        ]
        for f in datos
    ]

    # Fila final de resumen (Solo precio promedio)
    fila_resumen = [
        Paragraph("", estilo_c),
        Paragraph("<b>PROMEDIO</b>", estilo_l),
        Paragraph("", estilo_c),
        Paragraph(f"<b>{precio_promedio:,.0f}</b>", estilo_r),
        Paragraph("", estilo_r),
    ]

    anchos_cols = [40, 160, 100, 80, 80]
    tabla = Table([encabezados_pdf] + filas_pdf + [fila_resumen],
                  colWidths=anchos_cols, repeatRows=1)

    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F7FAFC")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2E8F0")),
    ]))

    story.append(tabla)
    doc.build(story)



# 5. MENÚ INTERACTIVO

def opcion_catalogo(datos):
    mostrar_tabla(datos, f"CATÁLOGO COMPLETO ({len(datos)} PRODUCTOS)", formato="fancy_grid")
    mostrar_metricas(datos)


def opcion_filtrar(datos):
    filtrados = filtrar_por_precio(datos)
    if filtrados:
        mostrar_tabla(filtrados, f"PRODUCTOS CON PRECIO > {PRECIO_MINIMO} ({len(filtrados)} PRODUCTOS)", formato="fancy_grid")
        mostrar_metricas(filtrados)
    else:
        print(f"\n No hay productos con precio superior a {PRECIO_MINIMO}.")


def opcion_buscar(datos):
    entrada = input("\nIngrese el ID del producto: ").strip()
    try:
        id_buscado = int(entrada)
    except ValueError:
        print("El ID debe ser un número entero.")
        return

    fila = buscar_por_id(datos, id_buscado)
    if fila:
        # Se usa formato simple/plano para evitar líneas excesivas de IA
        mostrar_tabla([fila], f"RESULTADO DE BÚSQUEDA - ID {id_buscado}", formato="simple")
    else:
        print(f"Advertencia: el ID {id_buscado} no existe en el catálogo.")


def opcion_exportar(datos):
    print("\n  1) Catálogo completo")
    print(f"  2) Solo productos con precio > {PRECIO_MINIMO}")
    eleccion = input("¿Qué desea exportar? (1/2): ").strip()

    if eleccion == "1":
        generar_pdf(datos, ARCHIVO_PDF, "Catálogo completo")
    elif eleccion == "2":
        generar_pdf(filtrar_por_precio(datos), ARCHIVO_PDF, f"Productos con precio > {PRECIO_MINIMO}")
    else:
        print("Opción no válida, exportación cancelada.")
        return
    print(f"\n[ÉXITO] Archivo '{os.path.basename(ARCHIVO_PDF)}' generado correctamente.")


def menu():
    try:
        datos = cargar_datos(ARCHIVO_EXCEL)
    except FileNotFoundError:
        print(f"⚠ No se encontró '{ARCHIVO_EXCEL}'. Colóquelo en la misma carpeta del script.")
        return
    print(f"Hoja cargada exitosamente. Total de registros leídos en memoria: {len(datos)}")

    while True:
        print("\n   MENÚ - REPORTE DE VENTAS TECNOVENTAS")
        print("1) Ver catálogo completo")
        print(f"2) Filtrar por precio (> {PRECIO_MINIMO})")
        print("3) Buscar por ID")
        print("4) Exportar informe PDF")
        print("5) Salir del programa")

        opcion = input("Seleccione una opción (1-5): ").strip()

        if opcion == "1":
            opcion_catalogo(datos)
        elif opcion == "2":
            opcion_filtrar(datos)
        elif opcion == "3":
            opcion_buscar(datos)
        elif opcion == "4":
            opcion_exportar(datos)
        elif opcion == "5":
            print("\nHasta luego.")
            break
        else:
            print("Opción no válida. Elija un número del 1 al 5.")


if __name__ == "__main__":
    menu()