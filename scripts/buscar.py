# Tema del proyecto: Buscador de clientes 

# Creador: Luis Miguel Castro Curequia

# Funciones:
# Este programa cumple las funciones de abrir y exportar los datos necesarios para validar clientes, 
# esto para diferentes unidades operativas y tipo de maestro,
# se puede filtrar la cantidad deseada de clientes

from fileinput import filename
# Importamos la libreria pandas, la cual permite trabajar con dataframes, leer e exportar datos y manejar excel desde codigo python 
import pandas as pd

# Implementamos un try-except para capturar los posibles errores en el sistema
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
        # Diseñamos un menu en la consola para el facil entendimiento del programa
        print("-------------MENU-------------")
        menu1 = ("---TIPO---")
        menu1+= ("\n 1. FACTURACION")
        menu1+= ("\n 2. ENVIO")
        menu1+= ("\n 2. INACTIVOS")
        print(menu1)
        # capturamos con un input el tipo de maestro ingresado por el usuario
        tipo = int(input("Por favor ingrese el tipo de maestro: "))
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
        # capturamos con un input la unidad operativa ingresada por el usuario
        ou = int(input("Por favor ingrese la unidad operativa: "))
        # creamos un arreglo vacio para ingresarle luego dependiendo del tipo de maestro, las columnas del archivo excel
        # las columnas en excel se manejan de igual forma que un arreglo, iniciando desde la posicion 0
        # en este caso utilizaremos las columnas de numero de cuenta, razon social, direccion, regiona, cl_perfil, canal y vendedor
        # estos datos en especifico para poder realizar la validacion de los clientes creados
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
        # Le pedimos con un input al usuario el nombre del archivo excel
        # PD : se puede trabajar con cual tipo de documento, puede ser csv, word y etc...
        # para ello es solo cambiar la extension indicada al final del input          
        filename =  input("Por favor Ingrese el nombre del archivo:  ") + ".xlsx"
        if filename == "":
            print("ERROR, Ingrese un archivo valido...")
        print("Escaneando archivo.....")
        # le proporcionamos la ubicacion de la carpeta donde se encuentran almacenados los diferentes maestros de las unidades operativas
        path = r"D:\Users\practicante.geserv1\OneDrive - Centro de Servicios Mundial SAS\Imágenes\lm\Python\Maestros-INFO\."
        # hacemos uso del metodo join de python el cual permite unir strings
        fullpath = os.path.join(path,filename)
        # hacemos uso de metodo read_excel para leer el archivo excel que le indicamos, este recibe como parametros el nombre de la hoja
        # donde se encuentran los datos. Las columnas que deseamos, ademas convertimos a string la columna numero de cuenta,
        # esto debido a que muchos id contienen guiones o en caso de panama el DV, en caso de no parsear la columna, no encontrariamos los datos
        # PD : en caso de querer abrir otro tipo de archivo, cambiaremos el read_excel por el read_ y el tipo de archivo
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
        # se crea arreglo vacio para luego ir ingresando la cantidad de clientes solicitada por el usuario
        clientes = []
        i = 0
        # utilizamos un ciclo while para incresar varios clientes
        while i < cantidad:
            print("Con que metodo quiere realizar la busqueda : \n 1.ID \n 2.NOMBRE_CLIENTE")
            buscar = int(input("Ingrese el valor: "))
            if buscar == 1:
                idCliente = input("Por favor ingrese el id del cliente: ")
                # utilizamos el metodo append para llenar el arreglo en el ciclo while con los datos dados por el usario
                clientes.append(idCliente)
                print(clientes)
                # Hacemos uso de un if que pregunte si el tamaño del arreglo es igual a la cantidad de clientes solicitado por el usuario
                # aplique varios filtros al archivo excel, esto utilizando el metodo isin, el cual es un metodo que ofrece Pandas.
                if len(clientes)==cantidad:
                    print("Aplicando filtros.......")
                    df = df[df["NUMERO_DE_CUENTA"].isin(clientes)]
                    print(df)
                    print("Filtro aplicado con exito")
            elif buscar == 2:    
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

    # Funcion que permite exportar los datos en un nuevo archivo excel
    def exportar_datos(df):
        print("Exportando archivo procesado")
        import os
        path = r"D:\Users\practicante.geserv1\OneDrive - Centro de Servicios Mundial SAS\Imágenes\lm\Python\Buscador-Python\output\."
        name = input("Ingrese el nombre con el cual guardar el archivo: ") + ".xlsx"
        fullname = os.path.join(path,name)
        # exportamos el resultado en un nuevo archivo excel
        df.to_excel(fullname,
            header= True,index = False)
        print("Archivo exportado con exito")

    if __name__ == "__main__":
        main()
        input("\tPROCESO TERMINADO, presione enter para salir...")

except TypeError:
    print("Error en el tipo de dato")

  
