
"""
Calculadora de eficiencia empresarial
Versión con historial para Excel y Power BI.

Cada ejecución guarda una evaluación nueva en Historial_Eficiencia.csv.
"""


import os
import sys
import tempfile
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# =========================================================
# CALCULADORA DE EFICIENCIA EMPRESARIAL
# Historial directo en Excel (.xlsx)
# =========================================================

PONDERACIONES = {
    "productividad": 0.25,
    "eficiencia_costos": 0.15,
    "calidad": 0.20,
    "cumplimiento_entregas": 0.15,
    "utilizacion_recursos": 0.10,
    "margen_operativo": 0.15,
}

MARGEN_OBJETIVO_PORCENTAJE = 30

ENCABEZADOS = [
    "Empresa",
    "Periodo",
    "Produccion real",
    "Produccion objetivo",
    "Costo presupuestado",
    "Costo real",
    "Unidades producidas",
    "Unidades conformes",
    "Entregas totales",
    "Entregas a tiempo",
    "Horas disponibles",
    "Horas productivas",
    "Ventas totales",
    "Costos operativos",
    "Productividad (%)",
    "Eficiencia de costos (%)",
    "Calidad (%)",
    "Cumplimiento de entregas (%)",
    "Utilizacion de recursos (%)",
    "Margen operativo (%)",
    "Indice global (%)",
    "Clasificacion",
    "Recomendaciones",
]


def obtener_ruta_historial():
    if getattr(sys, "frozen", False):
        carpeta = os.path.dirname(os.path.abspath(sys.executable))
        nombre = "Historial_Eficiencia.xlsx"
        ruta_local = os.path.join(carpeta, nombre)
        ruta_anterior = os.path.join(os.path.dirname(carpeta), nombre)

        if os.path.exists(ruta_local):
            return ruta_local
        if os.path.exists(ruta_anterior):
            return ruta_anterior
    else:
        carpeta = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(carpeta, "Historial_Eficiencia.xlsx")


def pedir_texto(mensaje):
    while True:
        valor = input(mensaje).strip()
        if valor:
            return valor
        print("Este campo es obligatorio.")


def pedir_periodo():
    actual = datetime.now().strftime("%Y-%m-%d")

    while True:
        periodo = input(
            f"Periodo (AAAA-MM-DD) [{actual}]: "
        ).strip()

        if not periodo:
            return actual

        try:
            datetime.strptime(periodo, "%Y-%m-%d")
            return periodo
        except ValueError:
            print("Formato incorrecto. Ejemplo: 2026-10-01.")

def pedir_numero(mensaje, positivo=False):
    while True:
        try:
            valor = float(
                input(mensaje).strip().replace(",", ".")
            )

            if valor < 0:
                print("El valor no puede ser negativo.")
                continue

            if positivo and valor <= 0:
                print("El valor debe ser mayor que cero.")
                continue

            return valor

        except ValueError:
            print("Entrada inválida. Escribe un número.")


def pedir_entero(mensaje, minimo=0):
    while True:
        try:
            valor = int(input(mensaje).strip())

            if valor < minimo:
                print(f"El valor mínimo permitido es {minimo}.")
                continue

            return valor

        except ValueError:
            print("Ingresa un número entero válido.")


def ingresar_datos():
    print("\n" + "=" * 55)
    print("     CALCULADORA DE EFICIENCIA EMPRESARIAL")
    print("=" * 55)

    # Empresa antes del periodo
    empresa = pedir_texto("Nombre de la empresa: ")
    periodo = pedir_periodo()

    print("\n--- PRODUCCIÓN ---")
    produccion_real = pedir_numero("Producción real (unidades): ")
    produccion_objetivo = pedir_numero(
        "Producción objetivo (unidades): ", positivo=True
    )

    print("\n--- COSTOS ---")
    costo_presupuestado = pedir_numero(
        "Costo presupuestado (COP): "
    )
    costo_real = pedir_numero("Costo real (COP): ")

    print("\n--- CALIDAD ---")
    unidades_producidas = pedir_entero(
        "Total de unidades producidas: ", minimo=1
    )
    unidades_conformes = pedir_entero(
        "Unidades conformes: "
    )

    while unidades_conformes > unidades_producidas:
        print("Las unidades conformes no pueden superar el total.")
        unidades_conformes = pedir_entero(
            "Unidades conformes: "
        )

    print("\n--- ENTREGAS ---")
    entregas_totales = pedir_entero(
        "Total de entregas: ", minimo=1
    )
    entregas_a_tiempo = pedir_entero(
        "Entregas realizadas a tiempo: "
    )

    while entregas_a_tiempo > entregas_totales:
        print("Las entregas a tiempo no pueden superar el total.")
        entregas_a_tiempo = pedir_entero(
            "Entregas realizadas a tiempo: "
        )

    print("\n--- UTILIZACIÓN DE RECURSOS ---")
    horas_disponibles = pedir_numero(
        "Horas disponibles: ", positivo=True
    )
    horas_productivas = pedir_numero(
        "Horas productivas: "
    )

    while horas_productivas > horas_disponibles:
        print("Las horas productivas no pueden superar las disponibles.")
        horas_productivas = pedir_numero(
            "Horas productivas: "
        )

    print("\n--- RESULTADOS FINANCIEROS ---")
    ventas_totales = pedir_numero(
        "Ventas totales (COP): ", positivo=True
    )
    costos_operativos = pedir_numero(
        "Costos operativos (COP): "
    )

    return {
        "Empresa": empresa,
        "Periodo": periodo,
        "Produccion real": produccion_real,
        "Produccion objetivo": produccion_objetivo,
        "Costo presupuestado": costo_presupuestado,
        "Costo real": costo_real,
        "Unidades producidas": unidades_producidas,
        "Unidades conformes": unidades_conformes,
        "Entregas totales": entregas_totales,
        "Entregas a tiempo": entregas_a_tiempo,
        "Horas disponibles": horas_disponibles,
        "Horas productivas": horas_productivas,
        "Ventas totales": ventas_totales,
        "Costos operativos": costos_operativos,
    }


def limitar_porcentaje(valor):
    return max(0.0, min(valor, 100.0))


def calcular_indicadores(datos):
    productividad = (
        datos["Produccion real"]
        / datos["Produccion objetivo"] * 100
    )

    if datos["Costo real"] == 0:
        eficiencia_costos = 100.0
    else:
        eficiencia_costos = (
            datos["Costo presupuestado"]
            / datos["Costo real"] * 100
        )

    calidad = (
        datos["Unidades conformes"]
        / datos["Unidades producidas"] * 100
    )

    cumplimiento_entregas = (
        datos["Entregas a tiempo"]
        / datos["Entregas totales"] * 100
    )

    utilizacion_recursos = (
        datos["Horas productivas"]
        / datos["Horas disponibles"] * 100
    )

    margen_operativo = (
        (datos["Ventas totales"] - datos["Costos operativos"])
        / datos["Ventas totales"] * 100
    )

    puntajes = {
        "productividad": limitar_porcentaje(productividad),
        "eficiencia_costos": limitar_porcentaje(eficiencia_costos),
        "calidad": limitar_porcentaje(calidad),
        "cumplimiento_entregas": limitar_porcentaje(
            cumplimiento_entregas
        ),
        "utilizacion_recursos": limitar_porcentaje(
            utilizacion_recursos
        ),
        "margen_operativo": limitar_porcentaje(
            margen_operativo / MARGEN_OBJETIVO_PORCENTAJE * 100
        ),
    }

    indice_global = sum(
        puntajes[nombre] * peso
        for nombre, peso in PONDERACIONES.items()
    )

    return {
        "Productividad (%)": productividad,
        "Eficiencia de costos (%)": eficiencia_costos,
        "Calidad (%)": calidad,
        "Cumplimiento de entregas (%)": cumplimiento_entregas,
        "Utilizacion de recursos (%)": utilizacion_recursos,
        "Margen operativo (%)": margen_operativo,
        "Indice global (%)": indice_global,
    }


def clasificar_eficiencia(indice):
    if indice >= 90:
        return "EXCELENTE"
    elif indice >= 80:
        return "MUY EFICIENTE"
    elif indice >= 70:
        return "EFICIENTE / ACEPTABLE"
    elif indice >= 60:
        return "REQUIERE MEJORAS"
    return "EFICIENCIA BAJA"


def generar_recomendaciones(indicadores):
    recomendaciones = []

    if indicadores["Productividad (%)"] < 80:
        recomendaciones.append(
            "Mejorar productividad y revisar tiempos de produccion."
        )

    if indicadores["Eficiencia de costos (%)"] < 90:
        recomendaciones.append(
            "Analizar sobrecostos y ajustar el presupuesto."
        )

    if indicadores["Calidad (%)"] < 95:
        recomendaciones.append(
            "Reducir defectos y fortalecer el control de calidad."
        )

    if indicadores["Cumplimiento de entregas (%)"] < 90:
        recomendaciones.append(
            "Analizar retrasos y mejorar la planificacion de entregas."
        )

    if indicadores["Utilizacion de recursos (%)"] < 75:
        recomendaciones.append(
            "Mejorar la utilizacion de horas y recursos disponibles."
        )

    if indicadores["Margen operativo (%)"] < 10:
        recomendaciones.append(
            "Revisar costos operativos y rentabilidad."
        )

    if not recomendaciones:
        recomendaciones.append(
            "Mantener el desempeno y monitorear los indicadores."
        )

    return recomendaciones


def mostrar_resultados(indicadores, clasificacion, recomendaciones):
    print("\n" + "=" * 55)
    print("                 RESULTADOS")
    print("=" * 55)

    for nombre, valor in indicadores.items():
        print(f"{nombre:<36}: {valor:.2f}%")

    print("-" * 55)
    print("Clasificacion:", clasificacion)
    print("\nRecomendaciones:")

    for numero, texto in enumerate(recomendaciones, start=1):
        print(f"{numero}. {texto}")

    print("=" * 55)


def preparar_hoja_historial(libro):
    if "Historial" not in libro.sheetnames:
        hoja = libro.create_sheet("Historial")
        hoja.append(ENCABEZADOS)
        return hoja

    hoja = libro["Historial"]
    filas = list(hoja.iter_rows(values_only=True))

    if not filas or all(valor is None for valor in filas[0]):
        hoja.delete_rows(1, hoja.max_row)
        hoja.append(ENCABEZADOS)
        return hoja

    encabezados_anteriores = list(filas[0])
    registros = []

    for fila in filas[1:]:
        if all(valor is None for valor in fila):
            continue

        registro = {}
        for i, encabezado in enumerate(encabezados_anteriores):
            if encabezado is not None:
                registro[encabezado] = (
                    fila[i] if i < len(fila) else None
                )
        registros.append(registro)

    # Reorganiza columnas para colocar Empresa antes de Periodo.
    hoja.delete_rows(1, hoja.max_row)
    hoja.append(ENCABEZADOS)

    for registro in registros:
        hoja.append([
            registro.get(encabezado, "")
            for encabezado in ENCABEZADOS
        ])

    return hoja


def aplicar_formato_excel(hoja):
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(
            "solid", fgColor="1F4E78"
        )
        celda.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    hoja.row_dimensions[1].height = 35

    for columna in range(1, hoja.max_column + 1):
        letra = get_column_letter(columna)
        longitud = max(
            (
                len(str(celda.value))
                for celda in hoja[letra]
                if celda.value is not None
            ),
            default=12
        )
        hoja.column_dimensions[letra].width = min(
            max(longitud + 2, 14), 42
        )

    # Indicadores en columnas O a U.
    for fila in range(2, hoja.max_row + 1):
        for columna in range(15, 22):
            hoja.cell(fila, columna).number_format = '0.00"%"'


def _obtener_ruta_archivado(ruta):
    carpeta = os.path.dirname(ruta)
    nombre, extension = os.path.splitext(os.path.basename(ruta))
    fecha = datetime.now().strftime("%d_%m_%Y")
    candidato = os.path.join(carpeta, f"{nombre}_{fecha}{extension}")
    consecutivo = 0

    while os.path.exists(candidato):
        consecutivo += 1
        candidato = os.path.join(
            carpeta,
            f"{nombre}_{fecha}_{consecutivo}{extension}"
        )

    return candidato


def guardar_historial_xlsx(registro, iniciar_nuevo=False):
    ruta = obtener_ruta_historial()

    try:
        if iniciar_nuevo:
            libro = Workbook()
            libro.active.title = "Historial"
            hoja = preparar_hoja_historial(libro)
            hoja.append([
                registro.get(encabezado, "")
                for encabezado in ENCABEZADOS
            ])
            aplicar_formato_excel(hoja)

            descriptor, ruta_temporal = tempfile.mkstemp(
                prefix=".Historial_Eficiencia_",
                suffix=".xlsx",
                dir=os.path.dirname(ruta)
            )
            os.close(descriptor)
            ruta_archivado = None

            try:
                try:
                    libro.save(ruta_temporal)
                finally:
                    libro.close()

                if os.path.exists(ruta):
                    ruta_archivado = _obtener_ruta_archivado(ruta)
                    os.rename(ruta, ruta_archivado)

                os.replace(ruta_temporal, ruta)
            except OSError as error:
                if ruta_archivado and not os.path.exists(ruta):
                    try:
                        os.rename(ruta_archivado, ruta)
                    except OSError as restauracion:
                        raise OSError(
                            f"No se pudo activar el nuevo historial ni "
                            f"restaurar el anterior. El anterior se conserva "
                            f"en: {ruta_archivado}. Error: {restauracion}"
                        ) from error
                raise
            finally:
                if os.path.exists(ruta_temporal):
                    os.remove(ruta_temporal)

            return ruta, ruta_archivado

        if os.path.exists(ruta):
            libro = load_workbook(ruta)
        else:
            libro = Workbook()
            libro.active.title = "Historial"

        hoja = preparar_hoja_historial(libro)
        hoja.append([
            registro.get(encabezado, "")
            for encabezado in ENCABEZADOS
        ])

        aplicar_formato_excel(hoja)
        libro.save(ruta)
        libro.close()
        return ruta, None

    except PermissionError:
        print(
            "\nNo se pudo guardar el archivo. "
            "Cierra Historial_Eficiencia.xlsx en Excel "
            "y vuelve a ejecutar el programa."
        )
        raise


def main():
    try:
        datos = ingresar_datos()
        indicadores = calcular_indicadores(datos)

        clasificacion = clasificar_eficiencia(
            indicadores["Indice global (%)"]
        )

        recomendaciones = generar_recomendaciones(indicadores)

        mostrar_resultados(
            indicadores,
            clasificacion,
            recomendaciones
        )

        registro = datos.copy()
        registro.update(indicadores)
        registro["Clasificacion"] = clasificacion
        registro["Recomendaciones"] = " | ".join(recomendaciones)

        ruta, ruta_archivado = guardar_historial_xlsx(registro)
        if ruta_archivado:
            print("\nHistorial anterior archivado en:")
            print(ruta_archivado)

        print("\nHistorial guardado correctamente en:")
        print(ruta)

    except KeyboardInterrupt:
        print("\nEjecucion cancelada por el usuario.")
    except PermissionError:
        pass
    except Exception as error:
        print(f"\nSe produjo un error: {error}")


if __name__ == "__main__":
    main()
