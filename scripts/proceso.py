from fileinput import filename
import pandas as pd
from array import array
try:
    def main():
        df = leer_archivo()
        df = agregar_filtros(df)

        exportar_datos(df)
        # mostrar_datos(df)

    def leer_archivo():
        print("---------------BIENVENIDO---------------")
        import os
        input_cols = [10,3,30,53,25,54,61]
        filename =  input("Por favor Ingrese el nombre del archivo:  ") + ".xlsx"
        if filename == "":
            print("ERROR, Ingrese un archivo valido...")
        print("Escaneando archivo.....")
        path = r"D:\Users\practicante.geserv1\OneDrive - Centro de Servicios Mundial SAS\Imágenes\lm\Python\LeerExcel-Python\input\."
        fullpath = os.path.join(path,filename)
        df = pd.read_excel(fullpath,
                        sheet_name="Hoja1",
                        header= 0,
                        usecols=input_cols,
                        converters={"CUPO_CREDITO":int})
        print("Archivo Escaneado")
        return df


    def agregar_filtros(df):
        cantidad = int(input("Por favor Ingrese la cantidad de clientes a filtrar: "))
        clientes = []
        i = 0
        while i < cantidad:
            print("Con que metodo quiere realizar la busqueda : \n 1.ID \n 2.NOMBRE_CLIENTE")
            buscar = int(input("Ingrese el valor: "))
            if buscar == 1:
                idCliente = input("Por favor ingrese el id del cliente.....")
                clientes.append(idCliente)
                print(clientes)
                if len(clientes)==cantidad:
                    print("Aplicando filtros.......")
                    df = df[df["NUMERO_DE_CUENTA"].isin(clientes)]
                    print(df)
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

    # def mostrar_datos(df):
    #     print(df.shape)
    #     df_cols = df.columns

    #     for col in df_cols:
    #         print(df[col].head(5))

    def exportar_datos(df):
        print("Exportando archivo procesado")
        import os
        path = r"D:\Users\practicante.geserv1\OneDrive - Centro de Servicios Mundial SAS\Imágenes\lm\Python\LeerExcel-Python\output\."
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