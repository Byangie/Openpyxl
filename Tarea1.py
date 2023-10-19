import openpyxl

def crear_archivo_excel():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Gastos"
    sheet.append(["Descripción", "Monto", "Fecha"])
    workbook.save("informe_gastos.xlsx")

def ingresar_gastos():
    gastos = []

    while True:
        descripcion = input("Ingrese la descripción del gasto: ")
        monto = float(input("Ingrese el monto del gasto: "))
        fecha = input("Ingrese la fecha del gasto: ")

        gastos.append([descripcion, monto, fecha])

        continuar = input("¿Desea seguir agregando gastos? (escriba 'salir' para finalizar, o enter para seguir): ")
        if continuar.lower() == "salir":
            break

    return gastos

def calcular_gastos(gastos):
    num_gastos = len(gastos)
    if num_gastos == 0:
        return None

    total_gastos = sum(gasto[1] for gasto in gastos)
    gasto_mayor = max(gastos, key=lambda x: x[1])
    gasto_menor = min(gastos, key=lambda x: x[1])

    return num_gastos, gasto_mayor, gasto_menor, total_gastos

def datos_excel(gastos):
    archivo_excel = "informe_gastos.xlsx"
    workbook = openpyxl.load_workbook(archivo_excel)
    sheet = workbook["Gastos"]
    for gasto in gastos:
        sheet.append(gasto)
    workbook.save(archivo_excel)

def intro():
    crear_archivo_excel()
    print("Ingrese los datos de sus gastos:")
    gastos = ingresar_gastos()

    if not gastos:
        print("No ingresó gastos. Completado.")
        return


    resumen = calcular_gastos(gastos)

    if resumen:
        num_gastos, gasto_mayor, gasto_menor, total_gastos = resumen
        print("\n")
        print("----------------------------------------------------------------------------------")
        print("Resumen de sus gastos")
        print(f"Número total de gastos: {num_gastos}")
        print(f"Su gasto más caro fue: {gasto_mayor[0]} - {gasto_mayor[1]} - ${gasto_mayor[2]}")
        print(f"Su gasto más barato fue: {gasto_menor[0]} - {gasto_menor[1]} - ${gasto_menor[2]}")
        print(f"Monto total de gastos: ${total_gastos}")
        print("----------------------------------------------------------------------------------")
    datos_excel(gastos)
    print("Sus gastos quedaron almacenados en el archivo de Excel. Tarea completada")

if __name__ == "__main__":
    intro()
