import pandas as pd
import numpy as np
from tkinter.messagebox import showinfo

def Procesar_CSV(Archivos : list):
    '''
    Procesar archivos de CSV en base a una lista de ubicaciones


    '''
    # Reemplazar de la lista todos los '\\' por '/'
    Archivos = [i.replace('\\' , '/') for i in Archivos]

    # Eliminar los archivos que no sean CSV
    Archivos = [i for i in Archivos if i.split(".")[-1] == "csv"]

    for i in Archivos:

        # Leer tipos de comprobantes
        Tipos_Comprobantes = pd.read_excel("TABLACOMPROBANTES.xls")

        # Leer Proveedores
        Proveedores = pd.read_csv("Proveedores.csv" , sep=";")

        # Leer hojas de Excel para agregarlas al archivo final
        Provincias = pd.read_excel("Modelo-Holistor-Ventas.xls" , sheet_name="Provincias")
        Tipo_DOC = pd.read_excel("Modelo-Holistor-Ventas.xls" , sheet_name="Tipo Doc.")

        # Renombrar 'Código' por 'Código Provincia' en 'Provincias'
        Provincias.rename(columns={'Código ': 'Código Provincia'} , inplace=True)

        # Hacer un trim en la columna de 'Provincias'
        Provincias['Provincias'] = Provincias['Provincias'].str.strip()

        # hacer un marge de Proveedores y Provincias
        Proveedores = pd.merge(Proveedores , 
                            Provincias , 
                            how='left' , 
                            left_on='Provincia' , 
                            right_on='Provincias')
        Proveedores.drop('Provincias' , axis=1 , inplace=True)

        Proveedores.rename(columns={'Tipo': 'Tipo Responsable'} , inplace=True)


        # Leer el archivo CSV
        Ventas = pd.read_csv(i , sep=";" , decimal=",")


        # Converir la columna 'Fecha de Emisión' a datetime de formato %Y-%m-%d
        Ventas['Fecha de Emisión'] = pd.to_datetime(Ventas['Fecha de Emisión'] , format='%Y-%m-%d')

        # Mostrar como %d/%m/%Y la columna 'Fecha de Emisión'
        Ventas['Fecha de Emisión'] = Ventas['Fecha de Emisión'].dt.strftime('%d/%m/%Y')

        # Columnas a sumar a 'Otros Tributos'
        Otros = ['Importe de Per. o Pagos a Cta. de Otros Imp. Nac.' , 'Importe de Impuestos Municipales' , 'Importe de Impuestos Internos' , 'Importe Otros Tributos']
        Ventas['Otros Tributos'] = Ventas[Otros].sum(axis=1)
        Ventas.drop(Otros , axis=1 , inplace=True)
        del Otros

        # Eliminar columnas
        Columnas_Eliminar = ['Total Neto Gravado' , 'Total IVA' , ]
        Ventas.drop(Columnas_Eliminar , axis=1 , inplace=True)
        del Columnas_Eliminar

        # Hacer un melt de las columnas de 'Neto Gravado IVA 0%' , 'Neto Gravado IVA 2,5%' , 'Importe IVA 2,5%' , 'Neto Gravado IVA 5%' , 'Importe IVA 5%' , 'Neto Gravado IVA 10,5%' , 'Importe IVA 10,5%' , 'Neto Gravado IVA 21%' , 'Importe IVA 21%' , 'Neto Gravado IVA 27%' , 'Importe IVA 27%' en 3 columnas: 'Neto Gravado' , 'Importe IVA' , 'Alicuota IVA'
        Melt = ['Fecha de Emisión' , 'Tipo de Comprobante' , 'Punto de Venta' , 'Número de Comprobante' , 'Número de Comprobante Hasta' , 'Tipo Doc. Comprador' , 'Nro. Doc. Comprador' , 'Denominación Comprador' , 'Fecha de Vencimiento del Pago' , 'Importe Total' , 'Moneda Original' , 'Tipo de Cambio' , 'Importe No Gravado' , 'Importe Exento' , 'Importe de Percepciones de Ingresos Brutos' , 'Percepción a No Categorizados' , 'Otros Tributos']
        Ventas = pd.melt(Ventas, id_vars=Melt ,  value_vars=['Neto Gravado IVA 0%' , 'Neto Gravado IVA 2,5%' , 'Importe IVA 2,5%' , 'Neto Gravado IVA 5%' , 'Importe IVA 5%' , 'Neto Gravado IVA 10,5%' , 'Importe IVA 10,5%' , 'Neto Gravado IVA 21%' , 'Importe IVA 21%' , 'Neto Gravado IVA 27%' , 'Importe IVA 27%'] , var_name='Tipo IVA' , value_name='Monto IVA')
        del Melt

        # Filtrar los que 'Monto IVA' sea distinto de np.nan
        Ventas = Ventas[Ventas['Monto IVA'].notna()]

        # Separar la columna 'Tipo IVA' en 2 columnas: 'Alicuota IVA' y 'Tipo IVA' donde 'Alicuota IVA' es el ultimo caracter de 'Tipo IVA'
        Ventas['Alicuota IVA'] = Ventas['Tipo IVA'].str.split(' ').str[-1]
        Ventas['Tipo IVA'] = Ventas['Tipo IVA'].str.split(' ').str[0:-2].str.join(' ')

        # Filtar los que 'Tipo IVA' sea 'Importe'
        Ventas = Ventas[Ventas['Tipo IVA'] != 'Importe']
        Ventas.drop('Tipo IVA' , axis=1 , inplace=True)

        Ventas['Alicuota IVA'] = Ventas['Alicuota IVA'].str.replace('%' , '').str.replace("," , ".").astype(float) /100
        Ventas['Importe IVA'] = round( (Ventas['Monto IVA'] * Ventas['Alicuota IVA']) , 2)

        # Renombrar columna 'Monto IVA' a 'Neto Gravado'
        Ventas.rename(columns={'Monto IVA': 'Neto Gravado'} , inplace=True)

        # Merge de 'Compras' y 'Tipos_Comprobantes'
        Ventas = pd.merge(Ventas ,
                            Tipos_Comprobantes[['Código CBTE' ,'Letra CBTE' , 'Tipo CBTE' ]] ,
                            how='left' ,
                            left_on='Tipo de Comprobante' ,
                            right_on='Código CBTE')
        Ventas.drop('Código CBTE' , axis=1 , inplace=True)
        del Tipos_Comprobantes

        # Merge de 'Ventas' y 'Proveedores'
        Ventas = pd.merge(Ventas ,
                            Proveedores[['CUIT' , 'Tipo Responsable' , 'Domicilio' , 'Código Provincia']] ,
                            how='left' ,
                            left_on=['Nro. Doc. Comprador'] ,
                            right_on=['CUIT'])
        Ventas.drop('CUIT' , axis=1 , inplace=True)
        del Proveedores

        Ventas['C.P'] = np.nan
        Ventas['Cód. Neto'] = "V01"
        Ventas ['CF Computable'] = Ventas['Importe IVA']
        Ventas['Cód. NG/EX'] = "NGC"
        Ventas['IVA Débito'] = Ventas['Importe IVA']
        Ventas['NG + E'] = Ventas['Importe No Gravado'] + Ventas['Importe Exento'] + Ventas['Otros Tributos']
        Ventas['Cód. P/R'] = np.nan
        Ventas['RETPER'] = Ventas['Importe de Percepciones de Ingresos Brutos'] + Ventas['Percepción a No Categorizados']
        Ventas['Pcia P/R'] = Ventas['Código Provincia']

        # Reordenar columnas en otro DataFrame
        Columnas_Exportar = ['Fecha de Emisión'  , 'Tipo CBTE' , 'Letra CBTE' , 'Punto de Venta' , 'Número de Comprobante' , 'Denominación Comprador' , 'Tipo Doc. Comprador' , 'Nro. Doc. Comprador' , 'Domicilio' , 'C.P' , 'Código Provincia' , 'Tipo Responsable' , 'Cód. Neto' , 'Neto Gravado' , 'Alicuota IVA' , 'Importe IVA' , 'IVA Débito' , 'Cód. NG/EX' , 'NG + E' , 'Cód. P/R' , 'RETPER' , 'Pcia P/R' , 'Importe Total'  ]
        Ventas_Ordenado = Ventas[Columnas_Exportar]

        # Encuentra las filas duplicadas basadas en las columnas especificadas
        duplicates = Ventas_Ordenado.duplicated(subset=['Tipo CBTE', 'Punto de Venta', 'Número de Comprobante', 'Nro. Doc. Comprador'])

        # Utiliza loc para asignar 0 a las filas duplicadas en 'NG + E' y 'RETPER'
        Ventas_Ordenado.loc[duplicates, ['NG + E', 'RETPER']] = 0

        # Ordenar por 'Nro. Doc. Comprador', 'Punto de Venta' y 'Número de Comprobante'
        Ventas_Ordenado.sort_values(by=['Nro. Doc. Comprador' , 'Punto de Venta' , 'Número de Comprobante'] , inplace=True)

        # Renomrar las columnas de Provincias y Tipo_DOC a sus nombres originales
        Provincias.rename(columns={'Código Provincia': 'Código '} , inplace=True)

        # Renombrar columnas de Compras_Ordenado a: 'Fecha Emisión ' , 'Fecha Recepción' , 'Cpbte' , 'Tipo' , 'Suc.' , 'Número' , 'Razón Social/Denominación Proveedor' , 'Tipo Doc.' , 'CUIT' , 'Domicilio' , 'C.P.' , 'Pcia' , 'Cond Fisc' , 'Cód. Neto' , 'Neto Gravado' , 'Alíc.' , 'IVA Liquidado' , 'IVA Crédito' , 'Cód. NG/EX' , 'Conceptos NG/EX' , 'Cód. P/R' , 'Perc./Ret.' , 'Pcia P/R' , 'Total'
        Ventas_Ordenado.rename(columns={'Fecha de Emisión': 'Fecha dd/mm/aaaa' , 'Tipo CBTE': 'Cpbte' , 'Letra CBTE': 'Tipo' , 'Punto de Venta': 'Suc.' , 'Número de Comprobante': 'Número' , 'Denominación Comprador': 'Razón Social o Denominación Cliente ' , 'Tipo Doc. Comprador': 'Tipo Doc.' , 'Nro. Doc. Comprador': 'CUIT' , 'Código Provincia': 'Pcia' , 'Tipo Responsable': 'Cond Fisc' , 'Importe IVA': 'IVA Liquidado' , 'CF Computable': 'IVA Crédito' , 'NG + E': 'Conceptos NG/EX' , 'RETPER': 'Perc./Ret.' , 'Importe Total': 'Total' , 'C.P':'C.P.' , 'Alicuota IVA':'Alíc.'} , inplace=True)

        # Obtener la ubicación de i
        Nombre_Archivo = str(i).split("/")[-1]
        Path = str(i).replace(Nombre_Archivo , "")

        # Exportar Compras_Ordenado, Provincias y Tipo_Doc a Excel XLS
        with pd.ExcelWriter(f'{Path}{Nombre_Archivo.replace(".CSV" , "")} - Procesado para Holisor.xls') as writer:
            Ventas_Ordenado.to_excel(writer, sheet_name='HWCompra-modelo', index=False)
            Provincias.to_excel(writer, sheet_name='Provincias', index=False)
            Tipo_DOC.to_excel(writer, sheet_name='Tipo Doc.', index=False)

    # Mostrar mensaje de finalización
    showinfo("Finalizado" , "Se han procesado todos los archivos y se han guardado en la misma ubicación que los archivos originales con el nombre de 'NombreArchivo - Procesado para Holistor.xls'")

if __name__ == "__main__":
    Archivos = ["F:/Proyectos Python/Scripts/CSV a Holistor/CSV-Ventas/comprobantes_periodo_202307_ventas_20230902_2347 (montos expresados en pesos).csv"]
    Procesar_CSV(Archivos)
