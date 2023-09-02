import pandas as pd
import numpy as np

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
Compras = pd.read_csv("CSV-Compras/comprobantes_periodo_202307_compras_20230901_1950 (montos expresados en pesos).csv" , sep=";" , decimal=",")


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


# Exportar a Resultado.CSV
#Compras.to_csv("Resultado.CSV" , index=False , sep=";" , decimal=",")

