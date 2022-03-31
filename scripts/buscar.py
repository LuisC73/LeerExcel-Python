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
        # mostrar_datos(df)

    # Funcion que permite indicar el archivo excel que queremos escanear para buscar datos
    def leer_archivo():
        print("---------------BIENVENIDO---------------")
        # Importamos libreria que permite unir strings
        import os
        # le indicamos al sistema que columnas del archivo excel deseamos
        print("-------------MENU-------------")
        menu = ("---UNIDADES OPERATIVAS---")
        menu+= ("\n 1. CGP")
        menu+= ("\n 2. KCR")
        menu+= ("\n 3. KCG")
        menu+= ("\n 4. KDH")
        menu+= ("\n 5. KIN")
        menu+= ("\n 6. KIS")
        menu+= ("\n 7. PIP")
        menu+= ("\n 8. KET")
        menu+= ("\n 9. PIE")
        print(menu)
        ou = int(input("Por favor ingrese la unidad operativa: "))
        menu2 = ("---TIPO---")
        menu2+= ("\n 1. FACTURACION")
        menu2+= ("\n 2. ENVIO")
        menu2+= ("\n 2. INACTIVOS")
        print(menu2)
        tipo = int(input("Por favor ingrese el tipo de maestro: "))
        input_cols = []
        if tipo==1:
            if ou==1:
                input_cols = [10,3,30,53,25,54,61]
            elif ou==2:
                input_cols = [4,3,13,36,8,37,44]
            elif ou==3:
                input_cols = [4,3,13,36,8,37,44]
            elif ou==4:
                input_cols = [10,3,30,53,25,54,61]
            elif ou==5:
                input_cols = [10,3,30,53,25,54,61]
            elif ou==6:
                input_cols = [10,3,30,53,25,54,61]
            elif ou==7:
                input_cols = [4,3,12,32,7,33,40]
            elif ou==8:
                input_cols = [10,3,30,53,25,54,61]
            elif ou==9:
                input_cols = [10,3,30,53,25,54,61]
            else:
                print("INGRESE UN VALOR VALIDO")
        elif tipo==2:
            if ou==1:            
                input_cols = [10,3,29,53,54,35]
            elif ou==2:
                input_cols = [4,3,13,36,8,37,44]
            elif ou==3:
                input_cols = [4,3,13,36,8,37,44]
            elif ou==4:
                input_cols = [10,3,30,53,25,54,61]
            elif ou==5:
                input_cols = [10,3,30,53,25,54,61]
            elif ou==6:
                input_cols = [10,3,30,53,25,54,61]
            elif ou==7:
                input_cols = [4,3,12,32,7,33,40]
            elif ou==8:
                input_cols = [10,3,30,53,25,54,61]
            elif ou==9:
                input_cols = [10,3,30,53,25,54,61]
            else:
                print("INGRESE UN VALOR VALIDO")
        else:
            print("ERROR EN EL SISTEMA")
            quit()                    
                 
        filename =  input("Por favor Ingrese el nombre del archivo:  ") + ".xlsx"
        if filename == "":
            print("ERROR, Ingrese un archivo valido...")
        print("Escaneando archivo.....")
        # le proporcionamos la ubicacion de la carpeta al sistema
        
        
        path = r"D:\Users\practicante.geserv1\OneDrive - Centro de Servicios Mundial SAS\Imágenes\lm\Python\Maestros-INFO\."
        fullpath = os.path.join(path,filename)
        df = pd.read_excel(fullpath,
                        sheet_name="Hoja1",
                        header= 0,
                        usecols=input_cols,
                        converters={"NUMERO_DE_CUENTA":str})
        print("Archivo Escaneado")
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
                idCliente = input("Por favor ingrese el id del cliente: ")
                clientes.append(idCliente)
                print(clientes)
                if len(clientes)==cantidad:
                    print("Aplicando filtros.......")
                    df = df[df["NUMERO_DE_CUENTA"].isin(clientes)]
                    print(df)
                    print("Filtro aplicado con exito")
            elif buscar == 2:
                if type(clientes)==int:
                    print("Error no puede tener dos tipos de datos diferentes en la busqueda")
                    exit()     
                nombreCliente = input("Por favor ingrese el nombre del cliente: ")
                clientes.append(nombreCliente)
                print(clientes)
                if len(clientes)> 1:
                    print("Aplicando filtros.......")
                    df = df[df["NOMBRE_CLIENTE"].isin(clientes)]
                    print("Filtro aplicado con exito")
            else:
                print("ERROR.....Ingrese un valor valido")
            i+=1
        return df

    # def mostrar_datos(df):
    #     print(df.shape)
    #     df_cols = df.columns

    #     for col in df_cols:
    #         print(df[col].head(5))

    # Funcion que permite exportar los datos en un nuevo archivo excel
    def exportar_datos(df):
        print("Exportando archivo procesado")
        import os
        path = r"D:\Users\practicante.geserv1\OneDrive - Centro de Servicios Mundial SAS\Imágenes\lm\Python\Buscador-Python\output\."
        name = input("Ingrese el nombre con el cual guardar el archivo: ") + ".xlsx"
        fullname = os.path.join(path,name)
        df.to_excel(fullname,
            header= True,index = False)
        print("Archivo exportado con exito")

    if __name__ == "__main__":
        main()
        input("\tPROCESO TERMINADO, presione enter para salir...")

except TypeError:
    print("Error en el tipo de dato.")

  