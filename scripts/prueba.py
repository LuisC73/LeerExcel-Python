from fileinput import filename
import pandas as pd

def main():
    df = leer_archivo()
    df = agregar_filtros(df)

    exportar_datos(df)

def leer_archivo():
    print("Leyendo archivo.....")
    import os
    input_cols = [0,1,2]
    path = r"C:\Users\Alejo\Documents\Luis\solo_python\prueba\input\."
    filename =  input("Ingrese el nombre del archivo:  ") + ".xlsx"
    fullpath = os.path.join(path,filename)
    df = pd.read_excel(fullpath,
                        sheet_name="Hoja1",
                        header= 0,
                        usecols=input_cols)

    return df


def agregar_filtros(df):
    print("Aplicando filtros.......")
    df = df[df["NOMBRE"]=="MIGUEL"]

    return df


# print(df.shape)
# df_cols = df.columns

# for col in df_cols:
#     print(df[col].head(5))

def exportar_datos(df):
    print("Exportando archivo procesado")
    import os
    path = r"C:\Users\Alejo\Documents\Luis\solo_python\prueba\output\."
    name = input("Ingrese el nombre con el cual guardar el archivo: ") + ".xlsx"
    fullname = os.path.join(path,name)
    df.to_excel(fullname,
            header= True,index = False)


if __name__ == "__main__":
    main()
    input("\tPROCESO TERMINADO, presione enter para salir...")

