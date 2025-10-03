import Funciones_Formato
import CalculoManual
import CalculoCanalesPostres
import CalculoDemograficos
import CalculoDemograficosPostres
import CalculoMarcasPostres
import CalculoMarcasYOQE
import CalculoRegionCanales
import CalculoSegmentoPostres
import os
import pandas as pd
import openpyxl

def procesar_datos(hoja_marcas_yogures_manual, hoja_marcas_queso_manual, hoja_marcas_postres_manual):
    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    direccion_base1 = os.path.join(directorio_actual, "Datos.xlsx")
    nombre_excel = "Reporte Ejecutivo.xlsx"
    direccion_base2 = os.path.join(directorio_actual, nombre_excel)

    # Crear y guardar el archivo Excel
    wb = openpyxl.Workbook()
    wb.save(direccion_base2)

    lista_base = Funciones_Formato.obtener_nombres_hojas(direccion_base1)

    for i in lista_base: 
        df = pd.read_excel(direccion_base1, sheet_name=i)
        subcadenas = i.split("_")

        if subcadenas[0] == "Demo" and subcadenas[1] != "PO": 
            CalculoDemograficos.DatosDemograficos(df, nombre_excel, i).procesar()

        elif subcadenas[0] == "Marcas" and subcadenas[1] == "yog":
            if not hoja_marcas_yogures_manual:
                CalculoMarcasYOQE.DatosMarcasYOQE(df, nombre_excel, i).procesar()
            else:
                CalculoManual.Realizar_hoja_formato_manual(df, nombre_excel, i).procesar()

        elif subcadenas[0] == "Marcas" and subcadenas[1] == "QE":
            if not hoja_marcas_queso_manual:
                CalculoMarcasYOQE.DatosMarcasYOQE(df, nombre_excel, i).procesar()
            else:
                CalculoManual.Realizar_hoja_formato_manual(df, nombre_excel, i).procesar()

        elif subcadenas[0] == "Demo" and subcadenas[1] == "PO":
            CalculoDemograficosPostres.DatosDemograficoPostres(df, nombre_excel, i).procesar()

        elif subcadenas[0] == "Regiones" and subcadenas[1] == "PO":
            CalculoCanalesPostres.DatosCanalesPostres(df, nombre_excel, i).procesar()

        elif subcadenas[0] == "Canales" and subcadenas[1] == "PO":
            # Changed this to use CalculoSegmentoPostres as per your requirement
            CalculoCanalesPostres.DatosCanalesPostres(df, nombre_excel, i).procesar()

        elif subcadenas[0] == "Marcas" and subcadenas[1] == "PO": 
            if not hoja_marcas_postres_manual:
                CalculoMarcasPostres.DatosMarcasPostres(df, nombre_excel, i).procesar()
            else:
                CalculoManual.Realizar_hoja_formato_manual(df, nombre_excel, i).procesar()

        elif subcadenas[0] == "Seg" and subcadenas[1] == "PO":   
            CalculoSegmentoPostres.DatosSegmentoPostre(df, nombre_excel, i).procesar()

        elif subcadenas[0] in ["Regiones", "Canales"] and subcadenas[1] != "PO":
            CalculoRegionCanales.DatosCanalesRegion(df, nombre_excel, i).procesar()

        else:
            CalculoManual.Realizar_hoja_formato_manual(df, nombre_excel, i).procesar()

    Funciones_Formato.summary_funcion(nombre_excel)
    print("Reporte Ejecutivo Realizado")
