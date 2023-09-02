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
        Provincias = pd.read_excel("Modelo-Holistor-Compras.xls" , sheet_name="Provincias")
        Tipo_DOC = pd.read_excel("Modelo-Holistor-Compras.xls" , sheet_name="Tipo Doc.")

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
        Compras = pd.read_csv(i , sep=";" , decimal=",")


        # Converir la columna 'Fecha de Emisión' a datetime de formato %Y-%m-%d
        Compras['Fecha de Emisión'] = pd.to_datetime(Compras['Fecha de Emisión'] , format='%Y-%m-%d')

        # Mostrar como %d/%m/%Y la columna 'Fecha de Emisión'
        Compras['Fecha de Emisión'] = Compras['Fecha de Emisión'].dt.strftime('%d/%m/%Y')

        # Columnas a sumar a 'Otros Tributos'
        Otros = ['Importe de Per. o Pagos a Cta. de Otros Imp. Nac.' , 'Importe de Impuestos Municipales' , 'Importe de Impuestos Internos' , 'Importe Otros Tributos']
        Compras['Otros Tributos'] = Compras[Otros].sum(axis=1)
        Compras.drop(Otros , axis=1 , inplace=True)
        del Otros

        # Eliminar columnas
        Columnas_Eliminar = ['Total Neto Gravado' , 'Total IVA' , 'Crédito Fiscal Computable', ]
        Compras.drop(Columnas_Eliminar , axis=1 , inplace=True)
        del Columnas_Eliminar

        # Hacer un melt de las columnas de 'Neto Gravado IVA 0%' , 'Neto Gravado IVA 2,5%' , 'Importe IVA 2,5%' , 'Neto Gravado IVA 5%' , 'Importe IVA 5%' , 'Neto Gravado IVA 10,5%' , 'Importe IVA 10,5%' , 'Neto Gravado IVA 21%' , 'Importe IVA 21%' , 'Neto Gravado IVA 27%' , 'Importe IVA 27%' en 3 columnas: 'Neto Gravado' , 'Importe IVA' , 'Alicuota IVA'
        Compras = pd.melt(Compras, id_vars=['Fecha de Emisión' , 'Tipo de Comprobante' , 'Punto de Venta' , 'Número de Comprobante' , 'Tipo Doc. Vendedor' , 'Nro. Doc. Vendedor' , 'Denominación Vendedor' , 'Importe Total' , 'Moneda Original' , 'Tipo de Cambio' , 'Importe No Gravado' , 'Importe Exento' , 'Importe de Percepciones de Ingresos Brutos' , 'Importe de Percepciones o Pagos a Cuenta de IVA' , 'Otros Tributos' ] ,  value_vars=['Neto Gravado IVA 0%' , 'Neto Gravado IVA 2,5%' , 'Importe IVA 2,5%' , 'Neto Gravado IVA 5%' , 'Importe IVA 5%' , 'Neto Gravado IVA 10,5%' , 'Importe IVA 10,5%' , 'Neto Gravado IVA 21%' , 'Importe IVA 21%' , 'Neto Gravado IVA 27%' , 'Importe IVA 27%'] , var_name='Tipo IVA' , value_name='Monto IVA')

        # Filtrar los que 'Monto IVA' sea distinto de np.nan
        Compras = Compras[Compras['Monto IVA'].notna()]

        # Separar la columna 'Tipo IVA' en 2 columnas: 'Alicuota IVA' y 'Tipo IVA' donde 'Alicuota IVA' es el ultimo caracter de 'Tipo IVA'
        Compras['Alicuota IVA'] = Compras['Tipo IVA'].str.split(' ').str[-1]
        Compras['Tipo IVA'] = Compras['Tipo IVA'].str.split(' ').str[0:-2].str.join(' ')

        # Filtar los que 'Tipo IVA' sea 'Importe'
        Compras = Compras[Compras['Tipo IVA'] != 'Importe']
        Compras.drop('Tipo IVA' , axis=1 , inplace=True)

        Compras['Alicuota IVA'] = Compras['Alicuota IVA'].str.replace('%' , '').str.replace("," , ".").astype(float) /100
        Compras['Importe IVA'] = round( (Compras['Monto IVA'] * Compras['Alicuota IVA']) , 2)

        # Renombrar columna 'Monto IVA' a 'Neto Gravado'
        Compras.rename(columns={'Monto IVA': 'Neto Gravado'} , inplace=True)

        # Merge de 'Compras' y 'Tipos_Comprobantes'
        Compras = pd.merge(Compras ,
                            Tipos_Comprobantes[['Código CBTE' ,'Letra CBTE' , 'Tipo CBTE' ]] ,
                            how='left' ,
                            left_on='Tipo de Comprobante' ,
                            right_on='Código CBTE')
        Compras.drop('Código CBTE' , axis=1 , inplace=True)
        del Tipos_Comprobantes

        # Merge de 'Compras' y 'Proveedores'
        Compras = pd.merge(Compras ,
                            Proveedores[['CUIT' , 'Tipo Responsable' , 'Domicilio' , 'Código Provincia']] ,
                            how='left' ,
                            left_on=['Nro. Doc. Vendedor'] ,
                            right_on=['CUIT'])
        Compras.drop('CUIT' , axis=1 , inplace=True)
        del Proveedores

        Compras['Fecha de Recepción'] = Compras['Fecha de Emisión']
        Compras['C.P'] = np.nan
        Compras['Cód. Neto'] = "CMG"
        Compras ['CF Computable'] = Compras['Importe IVA']
        Compras['Cód. NG/EX'] = "NGC"
        Compras['NG + E'] = Compras['Importe No Gravado'] + Compras['Importe Exento'] + Compras['Otros Tributos']
        Compras['Cód. P/R'] = np.nan
        Compras['RETPER'] = Compras['Importe de Percepciones de Ingresos Brutos'] + Compras['Importe de Percepciones o Pagos a Cuenta de IVA']
        Compras['Pcia P/R'] = Compras['Código Provincia']

        # Reordenar columnas en otro DataFrame
        Columnas_Exportar = ['Fecha de Emisión' , 'Fecha de Recepción' , 'Tipo CBTE' , 'Letra CBTE' , 'Punto de Venta' , 'Número de Comprobante' , 'Denominación Vendedor' , 'Tipo Doc. Vendedor' , 'Nro. Doc. Vendedor' , 'Domicilio' , 'C.P' , 'Código Provincia' , 'Tipo Responsable' , 'Cód. Neto' , 'Neto Gravado' , 'Alicuota IVA' , 'Importe IVA' , 'CF Computable' , 'Cód. NG/EX' , 'NG + E' , 'Cód. P/R' , 'RETPER' , 'Pcia P/R' , 'Importe Total'  ]
        Compras_Ordenado = Compras[Columnas_Exportar]

        # Encuentra las filas duplicadas basadas en las columnas especificadas
        duplicates = Compras_Ordenado.duplicated(subset=['Tipo CBTE', 'Punto de Venta', 'Número de Comprobante', 'Nro. Doc. Vendedor'])

        # Utiliza loc para asignar 0 a las filas duplicadas en 'NG + E' y 'RETPER'
        Compras_Ordenado.loc[duplicates, ['NG + E', 'RETPER']] = 0

        # Ordenar por 'Nro. Doc. Vendedor', 'Punto de Venta' y 'Número de Comprobante'
        Compras_Ordenado.sort_values(by=['Nro. Doc. Vendedor' , 'Punto de Venta' , 'Número de Comprobante'] , inplace=True)

        # Renomrar las columnas de Provincias y Tipo_DOC a sus nombres originales
        Provincias.rename(columns={'Código Provincia': 'Código '} , inplace=True)

        # Si la 'Letra CBTE' es 'B'o 'C' entonces la 'Alicuota IVA' es 0
        Compras_Ordenado.loc[Compras_Ordenado['Letra CBTE'].isin(['B' , 'C']) , 'Alicuota IVA'] = 0
        
        # Si la 'Letra CBTE' es 'B' o 'C' entonces el 'Neto Gravado' es igual al 'Importe Total'
        Compras_Ordenado.loc[Compras_Ordenado['Letra CBTE'].isin(['B' , 'C']) , 'Neto Gravado'] = Compras_Ordenado['Importe Total']

        # Renombrar columnas de Compras_Ordenado a: 'Fecha Emisión ' , 'Fecha Recepción' , 'Cpbte' , 'Tipo' , 'Suc.' , 'Número' , 'Razón Social/Denominación Proveedor' , 'Tipo Doc.' , 'CUIT' , 'Domicilio' , 'C.P.' , 'Pcia' , 'Cond Fisc' , 'Cód. Neto' , 'Neto Gravado' , 'Alíc.' , 'IVA Liquidado' , 'IVA Crédito' , 'Cód. NG/EX' , 'Conceptos NG/EX' , 'Cód. P/R' , 'Perc./Ret.' , 'Pcia P/R' , 'Total'
        Compras_Ordenado.rename(columns={'Fecha de Emisión': 'Fecha Emisión ' , 'Fecha de Recepción': 'Fecha Recepción' , 'Tipo CBTE': 'Cpbte' , 'Letra CBTE': 'Tipo' , 'Punto de Venta': 'Suc.' , 'Número de Comprobante': 'Número' , 'Denominación Vendedor': 'Razón Social/Denominación Proveedor' , 'Tipo Doc. Vendedor': 'Tipo Doc.' , 'Nro. Doc. Vendedor': 'CUIT' , 'Código Provincia': 'Pcia' , 'Tipo Responsable': 'Cond Fisc' , 'Importe IVA': 'IVA Liquidado' , 'CF Computable': 'IVA Crédito' , 'NG + E': 'Conceptos NG/EX' , 'RETPER': 'Perc./Ret.' , 'Importe Total': 'Total' , 'C.P':'C.P.' , 'Alicuota IVA':'Alíc.'} , inplace=True)

        # Obtener la ubicación de i
        Nombre_Archivo = str(i).split("/")[-1]
        Path = str(i).replace(Nombre_Archivo , "")

        # Exportar Compras_Ordenado, Provincias y Tipo_Doc a Excel XLS
        with pd.ExcelWriter(f'{Path}{Nombre_Archivo.replace(".CSV" , "")} - Procesado para Holisor.xls') as writer:
            Compras_Ordenado.to_excel(writer, sheet_name='HWCompra-modelo', index=False)
            Provincias.to_excel(writer, sheet_name='Provincias', index=False)
            Tipo_DOC.to_excel(writer, sheet_name='Tipo Doc.', index=False)

    # Mostrar mensaje de finalización
    showinfo("Finalizado" , "Se han procesado todos los archivos y se han guardado en la misma ubicación que los archivos originales con el nombre de 'NombreArchivo - Procesado para Holistor.xls'")
