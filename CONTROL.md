""Genera un libro de Excel para registrar ingresos y gastos en soles."""
from datetime import date
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

MONEDA = '"S/ " #,##0.00;[Red]("S/ " #,##0.00);"S/ "0.00'
MESES = "Enero Febrero Marzo Abril Mayo Junio Julio Agosto Septiembre Octubre Noviembre Diciembre".split()


def encabezado(hoja, fila):
    for celda in hoja[fila]:
        celda.fill = PatternFill("solid", fgColor="195653")
        celda.font = Font(name="Calibri", bold=True, color="FFFFFF")
        celda.alignment = Alignment(vertical="center")
    hoja.row_dimensions[fila].height = 26


def anchos(hoja, valores):
    for indice, ancho in enumerate(valores, 1):
        hoja.column_dimensions[get_column_letter(indice)].width = ancho


def validar(hoja, rango, mensaje, **regla):
    control = DataValidation(**regla, allow_blank=True, showErrorMessage=True)
    control.errorTitle = "Revisa el dato"
    control.error = mensaje
    hoja.add_data_validation(control)
    control.add(rango)


def crear(ruta):
    libro = Workbook()
    datos = libro.active
    datos.title = "Movimientos"
    datos.append(["Fecha", "Tipo", "Categoría", "Descripción", "Medio de pago", "Monto"])
    encabezado(datos, 1)
    anchos(datos, [17, 16, 24, 46, 21, 20])
    datos.freeze_panes = "C2"
    datos.column_dimensions["A"].number_format = "dd/mm/yyyy"
    datos.column_dimensions["F"].number_format = MONEDA
    for fila in range(2, 102):
        datos.cell(fila, 1).number_format = "dd/mm/yyyy"
        datos.cell(fila, 6).number_format = MONEDA
    tabla = Table(displayName="Registro", ref="A1:F101")
    tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    datos.add_table(tabla)
    validar(datos, "B2:B1048576", "Elige Ingreso o Gasto.", type="list",
            formula1='"Ingreso,Gasto"')
    validar(datos, "E2:E1048576", "Selecciona un medio de pago.", type="list",
            formula1='"Efectivo,Yape,Plin,Transferencia,Tarjeta,Otro"')
    validar(datos, "F2:F1048576", "Escribe un importe mayor que cero con hasta dos decimales.",
            type="custom", formula1="AND(ISNUMBER(F2),F2>0,ROUND(F2,2)=F2)")
    validar(datos, "A2:A1048576", "Ingresa una fecha válida.", type="date",
            operator="between", formula1="DATE(1900,1,1)", formula2="DATE(9999,12,31)")

    resumen = libro.create_sheet("Resumen")
    resumen.append(["Resumen anual"])
    resumen["A1"].font = Font(size=20, bold=True, color="195653")
    resumen.append(["Año a consultar", date.today().year])
    resumen["B2"].fill = PatternFill("solid", fgColor="FFF1BC")
    validar(resumen, "B2", "Escribe un año entre 1900 y 9998.", type="whole",
            operator="between", formula1=1900, formula2=9998)
    resumen.append(["Mes", "Ingresos", "Gastos", "Saldo del mes"])
    encabezado(resumen, 3)
    for mes, nombre in enumerate(MESES, 1):
        fila = mes + 3
        resumen.cell(fila, 1, nombre)
        for columna, tipo in [(2, "Ingreso"), (3, "Gasto")]:
            formula = (
                f'=SUMIFS(Movimientos!F:F,Movimientos!B:B,"{tipo}",'
                f'Movimientos!A:A,">="&DATE($B$2,{mes},1),'
                f'Movimientos!A:A,"<"&DATE($B$2,{mes + 1},1))'
            )
            resumen.cell(fila, columna, formula)
        resumen.cell(fila, 4, f"=B{fila}-C{fila}")
    resumen.append(["Total anual", "=SUM(B4:B15)", "=SUM(C4:C15)", "=B16-C16"])
    encabezado(resumen, 16)
    for fila in resumen.iter_rows(min_row=4, max_row=16, min_col=2, max_col=4):
        for celda in fila:
            celda.number_format = MONEDA
    anchos(resumen, [27, 21, 21, 21])
    resumen.freeze_panes = "B4"
    resumen.sheet_view.showGridLines = False
    # El modo exclusivo protege los registros si ya existe un libro con ese nombre.
    with Path(ruta).open("xb") as archivo:
        libro.save(archivo)


if __name__ == "__main__":
    destino = Path("Mis_finanzas.xlsx")
    try:
        crear(destino)
        print(f"Libro creado en {destino.resolve()}")
    except FileExistsError:
        print("Ya existe Mis_finanzas.xlsx. Elige otra carpeta para generar un libro nuevo.")
