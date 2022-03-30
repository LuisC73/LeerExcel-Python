from fileinput import filename
# Importamos la libreria pandas para leer archivos excel
import pandas as pd
# Implementamos un try-catch para capturar los posibles errores en el sistema
try:
     # Declaramos una funcion main principal de la cual se crearan y empezaran a funcionar las diferentes funciones
    def main():
        df = leer_archivo()
        df = agregar_filtros(df)

        exportar_datos(df)
    # Funcion que permite indicar el archivo excel que queremos escanear para buscar datos
    def leer_archivo():
        print("---------------BIENVENIDO---------------")
        # Importamos libreria que permite unir strings
        import os
        l = 0
         # creamos un array vacio para luego por consola ingresarle las columnas deseadas
        input_cols = []
        filename =  input("Por favor Ingrese el nombre del archivo:  ") + ".xlsx"
        if filename == "":
            print("ERROR, Ingrese un archivo valido...")
        print("Escaneando archivo.....")
        # le proporcionamos la ubicacion de la carpeta al sistema
        path = r"D:\Users\practicante.geserv1\OneDrive - Centro de Servicios Mundial SAS\Imágenes\lm\Python\LeerExcel-Python\input\."
        fullpath = os.path.join(path,filename)
        cols = 0
        # Utilizamos un ciclo while para llenar el array de columnas desde la consola
        while cols!= -1:
            print("--------- Para acabar escriba -1 ---------")
            cols = int(input("Ingrese la columna a buscar: "))
            input_cols.append(cols)
            print(input_cols)
        # Leemos el archivo de excel indicado
        df = pd.read_excel(fullpath,
                        sheet_name="Hoja1",
                        header= 0,
                        usecols=input_cols,
                        converters={"CUPO_CREDITO":int})
        print("Archivo Escaneado")
        print(df.shape)
        return df

    # Funcion que permite aplicar filtros a nuestra hoja de excel para encontrar los clientes que necesitemos
    def agregar_filtros(df):
        cantidad = int(input("Por favor Ingrese la cantidad de clientes a filtrar: "))
        clientes = []
        i = 0
        # utilizamos un ciclo while para incresar varios clientes
        while i < cantidad:
            print("Con que metodo quiere realizar la busqueda : \n 1.ID \n 2.NOMBRE_CLIENTE")
            buscar = int(input("Ingrese el valor: "))
            if buscar == 1:
                idCliente = input("Por favor ingrese el id del cliente.....")
                clientes.append(idCliente)
                if len(clientes)==cantidad:
                    print("Aplicando filtros.......")
                    df = df[df["NUMERO_DE_CUENTA"].isin(clientes)]
                    print("Filtro aplicado con exito")
            elif buscar == 2:
                nombreCliente = input("Por favor ingrese el nombre del cliente....")
                clientes.append(nombreCliente)
                if len(clientes)> 1:
                    print("Aplicando filtros.......")
                    df = df[df["NOMBRE_CLIENTE"].isin(clientes)]
                    print("Filtro aplicado con exito")
            else:
                print("ERROR.....Ingrese un valor valido")
            i+=1
        return df

    # Funcion que permite exportar los datos en un nuevo archivo excel
    def exportar_datos(df):
        print("Exportando archivo procesado")
        import os
        path = r"D:\Users\practicante.geserv1\OneDrive - Centro de Servicios Mundial SAS\Imágenes\lm\Python\DecidirCols-Python\output\."
        name = input("Ingrese el nombre con el cual guardar el archivo: ") + ".xlsx"
        fullname = os.path.join(path,name)
        df.to_excel(fullname,
            header= True,index = False)
        print("Archivo exportado con exito")

    if __name__ == "__main__":
        main()
        input("\tPROCESO TERMINADO, presione enter para salir...")

except Exception as e:
    print("ERROR EN EL SISTEMA: " + e)
