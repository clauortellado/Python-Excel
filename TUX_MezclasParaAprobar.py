# Report in Excel about MEZCLAS of preview month to Esteban
# Export TXT from app: FOM
# Import TXT and transform data
# Export data to XLS
# Export Pivot Sheet to XLS
# Send mail to ePala + Tere + Guille
# Check out Task in app: Asana

import pandas as pd

print("App para generar Reporte MEZCLAS Consumidas Mes anerior - y enviar a ePala - para Aprobar")


def main():
    #ImportTXT()
    CompleteXLS()


def ImportTXT():
    # Import TXT and transform data
    # Export data to XLS

    # file="K:/Sistema Costos/PreciosMEZCLA/FOM ConsumosNew/2021-01.txt"
    file = "2021-01.txt"

    df = pd.read_csv(file, encoding='ISO-8859-1') #por el caracter Ñ en el Header
    # "F. Desmotadora","F. Fabrica","Campaña","Receta","Clave","Lote","Num Fardo","Mezcla","Material","Ga Far"]]

    # Extract the first 10 chars + convert to date
    df["F.DMT"] = pd.to_datetime(df["F. Desmotadora"].str.slice(start=0, stop=10, step=None))
    df["F.FAB"] = pd.to_datetime(df["F. Fabrica"].str.slice(start=0, stop=10, step=None))
    # Aux MEZ
    df["MEZ"] = df["Mezcla"].str.slice(start=0, stop=6, step=None)
    # Convert in year/month/day
    df["AÑO"] = pd.DatetimeIndex(df["F.FAB"]).year
    df["MES"] = pd.DatetimeIndex(df["F.FAB"]).month
    df["DIA"] = pd.DatetimeIndex(df["F.FAB"]).day
    # Empty values
    #df["META_MEZCLA"]=""
    #df["META_MATERIAL"]=""
    #df["GRADO"]=""


    # Create a new dataframe with the selected cols 
    df_mezcla= df[["F.DMT","F.FAB",
            "Campana","Receta","Clave","Lote",
            "Num Fardo","Mezcla","Material","Ga Far", 
            "MEZ", 
            "AÑO", "MES", "DIA"]]

    # Export data to XLS - without Index Col
    df_mezcla.to_excel("2021-01.xlsx", index = False)
    print("txt exportado a xlsx")

    #['F.DMT','F.FAB','CAMPAÑA','RECETA','CLAVE','LOTE','NUM.FARDO','MEZCLA','MATERIAL','MA.FAR','GA.FAR'])




def CompleteXLS():
    # Complete cells META_MEZCLA + META_MATERIAL + GRADO
    file_Mes = "2021-01.xlsx"
    file_Aux = "Aux_FOM.xlsx"

    df_Mes = pd.read_excel(file_Mes)
    df_Mezcla = pd.read_excel(file_Aux, sheet_name="Mezcla")
    df_MetaMezcla = pd.read_excel(file_Aux, sheet_name="MetaMezcla")
    df_Grado = pd.read_excel(file_Aux, sheet_name="Grado")
  
    df_MesTotal = df_Mes.merge(df_Grado, on= "Ga Far")
    df_MesTotal = df_MesTotal.merge(df_Mezcla, on= "MEZ")
    df_MesTotal = df_MesTotal.merge(df_MetaMezcla, on= "Material")

    print(df_MesTotal)


if __name__ == "__main__":
    main()


