import pandas as pd
import re
from tkinter import filedialog
import os
import numpy as np

Qs= { "1": ["Enero" , "Febrero" , "Marzo"] , "2": ["Abril" , "Mayo" , "Junio"] , "3": ["Julio", "Agosto" , "Septiembre"] , "4": ["Octubre" , "Noviembre" , "Diciembre"]}
categorias= {"low_tier (XT 185 ,RVA 50,EP 350,VX 80, DTR 720 , VZ 30)": ["XT 185" ,"RVA 50","EP 350","VX 80", "DTR 720","VZ 30"] , 
             "mid_tier (DEP 250, DEP 450,DEP 550E,DEP 570E,SL 500E,DEM 300,DEM 400,DEM 500)" : ["DEP 250","DEP 450","DEP 550E","DEP 570E","SL 500E","DEM 300","DEM 400","DEM 500"] ,
             "high_tier (DGP 5050E,DGP 5550E,DGP 8050E,DGP 8550E,DGM 5000E,DGM 5500E,DGM 8000E,DGM 8500E,SL 8550E)": ["DGP 5050E","DGP 5550E","DGP 8050E","DGP 8550E","DGM 5000E","DGM 5500E","DGM 8000E","DGM 8500E","SL 8550E" ] ,
             "repetidoras (SLR 5100,SLR 1000)" : ["SLR 5100","SLR 1000"] ,
             "T_4X0 (T 400CO, T 470CO)": ["T 400CO", "T 470CO"] , 
             "otras_referencias": [] , 
             "Total General" : []
             }
MesesNum= {'Enero':1 , 'Febrero':2 , 'Marzo':3 , 'Abril':4 , 'Mayo':5 , 'Junio':6 , 'Julio':7 , 'Agosto':8 , 'Septiembre':9 , 'Octubre':10, 'Noviembre':11 , 'Diciembre':12}
MesesText= {1: 'Enero' , 2:'Febrero' , 3: 'Marzo' , 4: 'Abril' , 5: 'Mayo' , 6: 'Junio' , 7: 'Julio' , 8: 'Agosto' , 9: 'Septiembre' , 10: 'Octubre', 11: 'Noviembre' , 12: 'Diciembre'}


#DE AMBOS PORTAFOLIOS
def agregar_fecha(df , mes1 , mes2 , dia1 , dia2):
    if mes2=='diario':
        #print(mes1)
        #Fechas Diario
        df= df[df['Factura (Mes natural)']==mes1]
        df= df[df['Factura (Fecha de factura)'].str.startswith(dia1)]
        return df
        
    elif mes2 == 'acumulado':
        #INFORMES ACUMULADO MES
        df= df[df['Factura (Mes natural)']==mes1]
        return df

    elif dia1 != '' and dia2 != '':
        #fechas de Semanal
        df_filt= pd.DataFrame()
        if mes1 == mes2:
            #Si el mes es el mismo se filtra el mes elegido
            df= df[df['Factura (Mes natural)']==mes1]
            if int(dia1) < int(dia2):
                for i in range(int(dia1),int(dia2)+1 ,1):
                    if i <= 9:
                        fecha = "0" + str(i) + "."
                    else:
                        fecha= str(i) + "."
                    df_filt = df_filt.append(df[df['Factura (Fecha de factura)'].str.startswith(fecha)], ignore_index=True)
                
                return df_filt
        
        elif mes1 != mes2:
            #Si el mes es distinto, se filtra primero el mes inicio y despues el mes final
            #print(dia1,dia2,mes1,mes2)
            mesIni= MesesNum[mes1]
            mesFin= MesesNum[mes2]
            #Df donde se acuularan todos los resultados de cada mes
            df_filt1= pd.DataFrame()
            for i in range(mesIni , (mesFin+1),1):
                df_filt= df[df['Factura (Mes natural)']==MesesText[i]]
                for i in range(int(dia1),32,1):
                    if i <= 9:
                        fecha = "0" + str(i) + "."
                    else:
                        fecha= str(i) + "."
                
                    df_filt1 = df_filt1.append(df_filt[df_filt['Factura (Fecha de factura)'].str.startswith(fecha)], ignore_index=True)
                    
            return df_filt1

def guardar_reporte(df , df2, df3 ,indicativo_Hojas):
    file = filedialog.asksaveasfilename(title="Guardar Reporte", filetypes=(("Ficheros de Excel" , "*.xlsx"),("Ficheros de texto", "*.txt"),("Todos los ficheros","*.*")) , defaultextension='.xlsx')
    
    if indicativo_Hojas == 1:
        df.to_excel(file ,header= True , index = False, sheet_name="REPORTE")
    
    elif indicativo_Hojas ==2:
        #EXPORTAR A EXCEL FILTRADA
        with pd.ExcelWriter(file) as writer:
            df.to_excel(writer, header= True , index = False , sheet_name= "REPORTE_1")
            df2.to_excel(writer ,header= True , index = False, sheet_name="REPORTE_2")
    elif indicativo_Hojas ==3:
        #EXPORTAR A EXCEL FILTRADA
        with pd.ExcelWriter(file) as writer:
            df.to_excel(writer, header= True , index = False , sheet_name= "REPORTE_1")
            df2.to_excel(writer ,header= True , index = False, sheet_name="REPORTE_2")
            df3.to_excel(writer ,header= True , index = False, sheet_name="REPORTE_3")
    #Lanzar Excel
    os.startfile(file)

def cargar_Informe_SAP():
    '''
        title --> Es el titulo que se le va a dar a la ventana de seleccion
        initialdir --> Para que muestre desde qué nombre de directorio se va a abrir la ventana de seleccion
        filetypes --> Para seleccionar la extension de los tipos de archivos a visualizar
        '''
        
    file= filedialog.askopenfilename(title="Abrir"  , filetypes=(("Ficheros de Excel" , "*.xlsx"),("Ficheros de texto", "*.txt"),("Todos los ficheros","*.*")) )
    return file    

def calcular_dia_final(df):
    dias=[]
    for i in range(len(df['Factura (Fecha de factura)'])):
        fecha= df['Factura (Fecha de factura)'].iloc[i]
        diaNum= int(fecha[0:2])
        if  int(diaNum) not in dias:
            dias.append(int(diaNum))
    
    return max(dias)

def limpieza_datos(filename, ano):
    df= pd.read_excel(filename  , header=1 )
    if type(ano) == list:
        ano=list(map(int , ano))
        if ano[0]!= ano[1]:
            df= df[df['Factura (Año natural)'].isin(ano)]
        else:
            df= df[df['Factura (Año natural)']== ano[0]]
    else:
        
        df= df[df['Factura (Año natural)']== int(ano)]

    df= df[df['Factura (Estado del ciclo de vida de la factura)']=='Liberado']
    df= df[df['Cantidad de factura'] >= 1]
    df= df[~df['Factura'].str.startswith('J')]
    return df
        
def Anio(filename, ano):
    df= pd.read_excel(filename  , header=1 )
    if type(ano) == list:
        ano=list(map(int , ano))
        if ano[0]!= ano[1]:
            df= df[df['Factura (Año natural)'].isin(ano)]
        else:
            df= df[df['Factura (Año natural)']== ano[0]]
    else:
        
        df= df[df['Factura (Año natural)']== int(ano)]
    return df


#MOTOROLA
def Filtro_Mot(df):
    #LIMPIEZA DE LOS DATOS
    df = df[df['Categoría de producto'].isin(['RADIOS MOVILES MOTOROLA','RADIOS PORTATILES MOTOROLA','RADIOS MOVILES VERTEX','RADIOS PORTATILES VERTEX','REPETIDORAS MOTOROLA'])]
    return  df

def Filtro_Baterias(df):
    df= df[df['Categoría de producto'].isin(['BATERIAS IMPRES MOTOROLA','BATERIAS NORMALES MOTOROLA','BATERIAS VERTEX'])]
    return df

def Filtro_AyE(df):
    '''
    <FILTROS>
    categoria de producto
	MOTOROLA
	VERTEX
	
	<QUITAR>
	BATERIAS
	FACT PROY MOTOROLA
	RADIOS
	ALQUILER
    '''

    df= df[df['Categoría de producto'].str.contains("MOTOROLA|VERTEX")]
    df= df[~df['Categoría de producto'].isin(['BATERIAS IMPRES MOTOROLA','BATERIAS NORMALES MOTOROLA','BATERIAS VERTEX','RADIOS MOVILES MOTOROLA','RADIOS PORTATILES MOTOROLA','RADIOS MOVILES VERTEX','RADIOS PORTATILES VERTEX','REPETIDORAS MOTOROLA','FACTURACIÓN PROYECTOS MOTOROLA','SERVICIO DE ALQUILER MOTOROLA','LICENCIAS-SOFTWARE-MANUALES MOTOROLA','VARIOS SOLUCIONES MOTOROLA'])]
    return df


def AbreviacionMes(mes,ano):
    mesesIngles={'Enero' : 'jan' , 'Febrero' : 'feb' , 'Marzo':'mar' , 'Abril': 'apr' , 'Mayo' : 'may' ,'Junio': 'jun' , 'Julio': 'jul' , 'Agosto': 'aug' , 'Septiembre':'sept' , 'Octubre':'oct', 'Noviembre':'nov' , 'Diciembre':'dec' }
    mesENCURSO= mesesIngles[mes]
    anoENCURSO= ano[2:]
    mesAB= mesENCURSO + '-' + anoENCURSO
    return mesAB
    

def GenerarReporteSIMS(radios,baterias, AyE, mes ,ano):
    TodaLaInfo= pd.DataFrame()
    TodaLaInfo= radios.copy()
    TodaLaInfo = pd.concat([TodaLaInfo, baterias , AyE])
    #SE GENERA LA TABLA FINAL
    TablaSIMS= pd.DataFrame(columns=['DISTRIBUTOR ERP NUMBER','DISTRIBUTOR NAME','MONTH OF ACTIVITY','DISTRIBUTOR INVOICE NUMBER','DISTRIBUTOR INVOICE DATE','RESELLER  ERP NUMBER','RESELLER TRADING NAME*','RESELLER COMPANY NAME','RESELLER COUNTRY','RESELLER STATE','Part NUMBER (SKU)','Model NUMBER','QUANTITY','SERIAL NUMBER*','Sales Price*','PE or Promotion number (if applicable)*','End Customer Name','End Customer Industry Vertical'])
    #VALORES DINAMICOS
    TablaSIMS['DISTRIBUTOR INVOICE DATE']= TodaLaInfo['Factura (Fecha de factura)'].copy()
    TablaSIMS['RESELLER TRADING NAME*'] = TodaLaInfo['Factura (Cliente)'].copy()
    TablaSIMS['Part NUMBER (SKU)']= TodaLaInfo['Producto'].copy()
    TablaSIMS['Model NUMBER']= TodaLaInfo['Producto (Texto)'].copy()
    TablaSIMS['QUANTITY']= TodaLaInfo['Cantidad de factura'].copy()
    TablaSIMS['Sales Price*']=TodaLaInfo['Valor neto facturado'].copy()
    TablaSIMS['PE or Promotion number (if applicable)*']=TodaLaInfo['Pedido de cliente (Promoción)']
    TablaSIMS['RESELLER  ERP NUMBER'] = 0 #SITUACION INICIAL , POR DEFECTO CUANDO EL CLIENTE NO ESTÁ REGISTRADO EN EL SALES VIEW
    TablaSIMS['RESELLER STATE']= 'Bogotá' #SITUACION INICIAL , POR DEFECTO CUANDO EL CLIENTE NO ESTÁ REGISTRADO EN EL SALES VIEW
    #INSERTAR EL CODIGO SIMS
    #CARGAR TABLA DE SIMS DADA POR MOTOROLA
    tablaSALESVIEW= pd.read_excel("SALES VIEW - NOLA RESELLER COUNTRIES & STATES (COLOMBIA).xlsx" , header=3)
    tablaSALESVIEW = tablaSALESVIEW.fillna(0)
    tablaSeleccionada= tablaSALESVIEW[['PARTNER NAME', 'ERP NUMBER', 'RESELLER STATE']]
    tablaSeleccionada['ERP NUMBER'] = tablaSeleccionada['ERP NUMBER'].astype(int)

    for i in range(len(tablaSeleccionada)):
        ind = list(np.where(TablaSIMS['RESELLER TRADING NAME*'].str.contains(tablaSeleccionada['PARTNER NAME'].iloc[i])))
        indice= ind[0]
        for j in indice:
            TablaSIMS['RESELLER  ERP NUMBER'].iloc[j] = tablaSeleccionada['ERP NUMBER'].iloc[i]
            TablaSIMS['RESELLER STATE'].iloc[j]= tablaSeleccionada['RESELLER STATE'].iloc[i]

    #VALORES ESTATICOS
    TablaSIMS['MONTH OF ACTIVITY']= AbreviacionMes(mes,ano)
    TablaSIMS['RESELLER COUNTRY']='Colombia'
    TablaSIMS['DISTRIBUTOR ERP NUMBER']='168442'
    TablaSIMS['DISTRIBUTOR NAME']='Meltec'
    TablaSIMS['End Customer Name']='UNKNOWN'
    TablaSIMS['End Customer Industry Vertical']='Unknown'

    
    #Reemplaza los valores de NaN por 0 o UNKNOWN
    TablaSIMS['QUANTITY'] = TablaSIMS['QUANTITY'].fillna(0)
    TablaSIMS['Sales Price*'] = TablaSIMS['Sales Price*'].fillna(0)
    TablaSIMS['PE or Promotion number (if applicable)*'] = TablaSIMS['PE or Promotion number (if applicable)*'].replace('#','UNKNOWN')
    
    return TablaSIMS
    




def cambiar_descripciones():
    #BASE DE DATOS REFERENCIAS DE RADIOS
    dB= pd.read_excel('./DatabaseReferencias.xlsx' , sheet_name= "Database" , header = 0)
    #REPORTE FILTRADO
    df_RP= pd.read_excel("ArchivoTemp.xlsx" ,sheet_name="SELL-THROUGH")
    longitudDB= len(dB.index)
    longitud_df_RP= len(df_RP)
    
    for i in range(longitud_df_RP):
        patron= df_RP['Producto'][i]
        for j in range(longitudDB):
            if (re.fullmatch(patron,dB['ID de producto'][j]) != None ):
                df_RP['Producto (Texto)'].replace([df_RP['Producto (Texto)'][i]] , dB['Descripción de producto'][j] , inplace= True )
                break
    return df_RP

def ContarCantidadesMotorola(df):
    radios=0
    t470co=0
    xt185=0
    repetidoras=0
    cantTotal=0

    for i in df['Cantidad de factura']:
        cantTotal+=i

    for i in range(len(df['Producto (Texto)'])):
        if df['Producto (Texto)'][i] == "T 470CO":
            t470co+= df['Cantidad de factura'][i]
        elif df['Producto (Texto)'][i] == "XT 185":
            xt185+= df['Cantidad de factura'][i]
        elif df['Categoría de producto'][i] == "REPETIDORAS MOTOROLA":
            repetidoras+= df['Cantidad de factura'][i]
        else:
            radios+=df['Cantidad de factura'][i]
    return  radios, t470co , xt185 , repetidoras , cantTotal
    
def textoCorreoMot(dia1 , dia2 , mes1  , mes2, anio , radios, t470co , xt185 , repetidoras , cantTotal):
    if dia2==0 and mes2==0:
    #INFORMES DIARIOS
        if cantTotal != 0:
            textoCorreo= '''
Buenos días, cordial saludo

Envío SELL THROUGH RADIOS del '''+dia1+''' de '''+mes1.lower()+''' del '''+anio+''' , con:
        
'''
        
            if radios != 0 : textoCorreo+= "Radios = " + str(radios) + "\n"
            if t470co != 0 : textoCorreo+= "T470 = " + str(t470co) + "\n"
            if xt185 != 0 : textoCorreo+= "XT 185 = " + str(xt185) + "\n"
            if repetidoras != 0 : textoCorreo+= "Repetidoras = " + str(repetidoras) + "\n"

            textoCorreo+="\nPara un total general del día de " + str(cantTotal) + " dispositivos.\n"
            textoCorreo+='''
PEGAR RECORTE

Quedo atento a tus comentarios, saludos.'''

            return textoCorreo
        else:
            textoCorreo='''
Buenos días, cordial saludo

Informo que no se reportan radios ni repetidoras facturados el '''+dia1+''' de '''+mes1.lower()+''' del '''+anio+'''

Quedo atento a tus comentarios, saludos.
        '''
            return textoCorreo
    
    #INFORME SEMANAL
    elif dia1 != '' and dia2 != '':
        if mes1 == mes2:
            #MES IGUAL
            if cantTotal != 0:
                textoCorreo= '''
Buenos días, cordial saludo

Envío SELL THROUGH RADIOS del '''+dia1+''' al '''+dia2+''' de '''+mes1.lower()+''' del '''+anio+''' , con:
        
'''
        
                if radios != 0 : textoCorreo+= "Radios = " + str(radios) + "\n"
                if t470co != 0 : textoCorreo+= "T470 = " + str(t470co) + "\n"
                if xt185 != 0 : textoCorreo+= "XT 185 = " + str(xt185) + "\n"
                if repetidoras != 0 : textoCorreo+= "Repetidoras = " + str(repetidoras) + "\n"

                textoCorreo+="\nPara un total general semanal de " + str(cantTotal) + " dispositivos.\n"
                textoCorreo+='''
PEGAR RECORTE

Quedo atento a tus comentarios, saludos.'''

                return textoCorreo
            else:
                
                textoCorreo='''
Buenos días, cordial saludo

Informo que no se reportan radios ni repetidoras facturados del '''+dia1+''' al '''+dia2+''' de '''+mes1.lower()+''' del '''+anio+'''

Quedo atento a tus comentarios, saludos.
        '''
                return textoCorreo
        else:
            #MESES DISTINTOS
            if cantTotal != 0:
                textoCorreo= '''
Buenos días, cordial saludo

Envío SELL THROUGH RADIOS del '''+dia1+''' de '''+mes1.lower()+''' al '''+dia2+''' de '''+mes2.lower()+''' del '''+anio+''' , con:
        
'''
        
                if radios != 0 : textoCorreo+= "Radios = " + str(radios) + "\n"
                if t470co != 0 : textoCorreo+= "T470 = " + str(t470co) + "\n"
                if xt185 != 0 : textoCorreo+= "XT 185 = " + str(xt185) + "\n"
                if repetidoras != 0 : textoCorreo+= "Repetidoras = " + str(repetidoras) + "\n"

                textoCorreo+="\nPara un total general semanal de " + str(cantTotal) + " dispositivos.\n"
                textoCorreo+='''
PEGAR RECORTE

Quedo atento a tus comentarios, saludos.'''

                return textoCorreo
            else:
                textoCorreo='''
Buenos días, cordial saludo

Informo que no se reportan radios ni repetidoras facturados del '''+dia1+''' al '''+dia2+''' de '''+mes1.lower()+''' del '''+anio+'''

Quedo atento a tus comentarios, saludos.
        '''
                return textoCorreo

'''REPORTE POR Q'''
def calcular_cantidades_Q(data):
    #TABLA PARA MOSTRAR Qs
    tabla= pd.DataFrame(columns=['CATEGORIA', 'Q1', 'Q2','Q3','Q4'])
    otrasRef = pd.DataFrame()
    #print(list(data.columns.values))
    tabla["CATEGORIA"] = categorias.keys()

    otrasRef= data[~data['Producto (Texto)'].isin(categorias["low_tier (XT 185 ,RVA 50,EP 350,VX 80, DTR 720 , VZ 30)"] + categorias["mid_tier (DEP 250, DEP 450,DEP 550E,DEP 570E,SL 500E,DEM 300,DEM 400,DEM 500)"] + categorias["high_tier (DGP 5050E,DGP 5550E,DGP 8050E,DGP 8550E,DGM 5000E,DGM 5500E,DGM 8000E,DGM 8500E,SL 8550E)"] + categorias["repetidoras (SLR 5100,SLR 1000)"] + categorias["T_4X0 (T 400CO, T 470CO)"])]
    for i in range(1,5):
        df_Q= data[data['Factura (Mes natural)'].isin(Qs[str(i)])]
        
        for llave in categorias:
            df= pd.DataFrame()
            if llave != "otras_referencias" and llave != "Total General":
                df= df_Q[df_Q['Producto (Texto)'].isin(categorias[llave])]
                indice= tabla.index[tabla["CATEGORIA"]== llave][0]
                tabla["Q"+ str(i)].iloc[indice] = df['Cantidad de factura'].sum()
            elif llave == "Total General":
                indice= tabla.index[tabla["CATEGORIA"]== llave][0]
                tabla["Q"+ str(i)].iloc[indice] = tabla["Q"+ str(i)].sum()
            else:
                df_otros= df_Q[~df_Q['Producto (Texto)'].isin(categorias["low_tier (XT 185 ,RVA 50,EP 350,VX 80, DTR 720 , VZ 30)"] + categorias["mid_tier (DEP 250, DEP 450,DEP 550E,DEP 570E,SL 500E,DEM 300,DEM 400,DEM 500)"] + categorias["high_tier (DGP 5050E,DGP 5550E,DGP 8050E,DGP 8550E,DGM 5000E,DGM 5500E,DGM 8000E,DGM 8500E,SL 8550E)"] + categorias["repetidoras (SLR 5100,SLR 1000)"] + categorias["T_4X0 (T 400CO, T 470CO)"])]
                try:
                    otrasRef.append(df_otros)
                except:
                    pass
                indice= tabla.index[tabla["CATEGORIA"]== llave][0]
                tabla["Q"+ str(i)].iloc[indice] = df_otros['Cantidad de factura'].sum()

    return tabla , otrasRef


def calcular_cantidades_mes_a_mes(data):
    tabla= pd.DataFrame(columns=['CATEGORIA','Enero', 'Febrero', 'Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'])
    tabla["CATEGORIA"] = categorias.keys()
    
    for mes in range(1,13):
        df_mes= data[data['Factura (Mes natural)']== tabla.columns[mes]]
        
        for llave in categorias:
            df= pd.DataFrame()
            if llave != "otras_referencias" and llave != "Total General":
                df= df_mes[df_mes['Producto (Texto)'].isin(categorias[llave])]
                indice= tabla.index[tabla["CATEGORIA"]== llave][0]
                tabla[tabla.columns[mes]].iloc[indice] = df['Cantidad de factura'].sum()
            elif llave == "Total General":
                indice= tabla.index[tabla["CATEGORIA"]== llave][0]
                tabla[tabla.columns[mes]].iloc[indice] = tabla[tabla.columns[mes]].sum()
            else:
                df= df_mes[~df_mes['Producto (Texto)'].isin(categorias["low_tier (XT 185 ,RVA 50,EP 350,VX 80, DTR 720 , VZ 30)"] + categorias["mid_tier (DEP 250, DEP 450,DEP 550E,DEP 570E,SL 500E,DEM 300,DEM 400,DEM 500)"] + categorias["high_tier (DGP 5050E,DGP 5550E,DGP 8050E,DGP 8550E,DGM 5000E,DGM 5500E,DGM 8000E,DGM 8500E,SL 8550E)"] + categorias["repetidoras (SLR 5100,SLR 1000)"] + categorias["T_4X0 (T 400CO, T 470CO)"])]
                indice= tabla.index[tabla["CATEGORIA"]== llave][0]
                tabla[tabla.columns[mes]].iloc[indice] = df['Cantidad de factura'].sum()
    return tabla


def tabla_ref_radios(df):
    df=pd.DataFrame(df)

    #CREACION DE LA TABLA 
    tabla= pd.DataFrame(columns=['REFERENCIA', 'PRODUCTO'])
    
    df2= df.loc[:,['Producto','Producto (Texto)']]
    df2= df2.drop_duplicates('Producto')
    df2= df2.reset_index() #Se resetea el index debido a que al eliminar los duplicados los index quedan variados y en la asignacion de los valores a las columnas por mes no funciona
    tabla['REFERENCIA']= df2['Producto']
    tabla['PRODUCTO'] = df2['Producto (Texto)']
    return tabla

def FiltroMesesConAños(df , anios , meses):
    #Df donde se acumulará todos los resultados de cada mes
    df_filt1= pd.DataFrame()
    
    mesIni= MesesNum[meses[0]]
    mesFin= MesesNum[meses[1]]
    #MISMO AÑO
    if anios[0] == anios[1]:
        
        for i in range(mesIni , (mesFin+1),1):
            df_filt= df[df['Factura (Mes natural)']==MesesText[i]]
            df_filt1 = df_filt1.append(df_filt , ignore_index=True)
                
        
    #DISTINTO AÑO
    elif anios[0] != anios[1]:
        for anio in range(int(anios[0]),  (int(anios[1]))+1  ):
            #SI NO CORRESPONDE AL ULTIMO AÑO DE ANALISIS
            if str(anio) != anios[1]:

                df= df[df['Factura (Año natural)']== anio]
                for i in range(mesIni , 13 , 1):
                    df_filt= df[df['Factura (Mes natural)']==MesesText[i]] 
                    df_filt1 = df_filt1.append(df_filt, ignore_index=True)
            #CORRESPONDE AL ULTIMO AÑO DE ANALISIS
            elif str(anio) == anios[1]:
                if mesFin == 1:
                    df_filt= df[df['Factura (Mes natural)']==MesesText[1]]
                    df_filt1 = df_filt1.append(df_filt, ignore_index=True)
                
                elif mesFin != 1:
                    df= df[df['Factura (Año natural)']== anio]
                    for i in range(1 , mesFin+1 , 1):
                        df_filt= df[df['Factura (Mes natural)']==MesesText[i]]
                        df_filt1 = df_filt1.append(df_filt , ignore_index=True)
    
    return df_filt1

def tabla_minimos(tabla , anios , meses):
    mesIni= MesesNum[meses[0]]
    mesFin= MesesNum[meses[1]]

    #MISMO AÑO
    if anios[0] == anios[1]:
        df= pd.DataFrame()
        df= pd.read_excel("ArchivoTemp.xlsx")
        df_mismo_anio= df[df['Factura (Año natural)']== int(anios[0])]
        for i in range(mesIni , (mesFin+1),1):
            tabla[MesesText[i] + " "+ anios[0]] = 0
            for referencia in tabla['REFERENCIA']:
                df_mes= df_mismo_anio[df_mismo_anio['Factura (Mes natural)']==MesesText[i]]
                df2= df_mes[df_mes['Producto'] == referencia]
                
                indice= tabla.index[tabla["REFERENCIA"] == referencia][0]
                tabla[MesesText[i] + " "+anios[0]].iloc[indice] = df2['Cantidad de factura'].sum()
                
    #DISTINTO AÑO
    elif anios[0] != anios[1]:
        for anio in range(int(anios[0]),  (int(anios[1])) + 1 ):
            df= pd.DataFrame()
            df= pd.read_excel("ArchivoTemp.xlsx")
            df_mes= pd.DataFrame()
            #SI NO CORRESPONDE AL ULTIMO AÑO DE ANALISIS
            
            if anio != int(anios[1]):
                #AÑO A EVALUAR
                df_anio_evaluar= pd.DataFrame()
                df_anio_evaluar= df[df['Factura (Año natural)']== anio ]
                
                for i in range(mesIni , 13 , 1):
                    tabla[MesesText[i] + " "+ str(anio)] = 0
                    for referencia in tabla['REFERENCIA']:

                        df_mes= df_anio_evaluar[df_anio_evaluar['Factura (Mes natural)']== MesesText[i] ]
                        df2= df_mes[df_mes['Producto'] == referencia]
                
                        indice= tabla.index[tabla["REFERENCIA"] == referencia][0]
                        tabla[MesesText[i] + " "+ str(anio) ].iloc[indice] = df2['Cantidad de factura'].sum()

            #CORRESPONDE AL ULTIMO AÑO DE ANALISIS
            elif anio == int(anios[1]):
                #AÑO A EVALUAR
                df_anio_final= pd.DataFrame()
                df_anio_final= df[df['Factura (Año natural)']== int(anio) ]
                tabla[MesesText[1] + " "+ str(anio)] = 0
                
                #EL MES DE ENERO ES EL MES FINAL
                if mesFin == 1:
                    df_mes= df_anio_final[df_anio_final['Factura (Mes natural)']==MesesText[1]]
                    print(df_mes)
                    for referencia in tabla['REFERENCIA']:
                        df2= df_mes[df_mes['Producto'] == referencia]
                        indice= tabla.index[tabla["REFERENCIA"] == referencia][0]
                        tabla[MesesText[1] + " "+ str(anio) ].iloc[indice] = df2['Cantidad de factura'].sum()
                #EL MES FINAL ES OTRO DISTINTO A ENERO
                elif mesFin != 1:
                
                    for i in range(1 , mesFin+1 , 1):
                        tabla[MesesText[i] + " "+ str(anio)] = 0
                        
                        df_mes= df_anio_final[df_anio_final['Factura (Mes natural)']==MesesText[i]]
                        for referencia in tabla['REFERENCIA']:
                            df2= df_mes[df_mes['Producto'] == referencia]
                            indice= tabla.index[tabla["REFERENCIA"] == referencia][0]
                            tabla[MesesText[i] + " "+ str(anio) ].iloc[indice] = df2['Cantidad de factura'].sum()
                else:
                    print("ERROR")
    return tabla

def hallar_media(df):
    prom_meses1= 6
    prom_meses2= 12


    df=pd.DataFrame(df)
    columnas= df.columns
    #MEDIA DE ULTIMOS SEIS MESES
    longitudCols= len(df.columns)
    longitudCols-=2 #que no cuente las columnas de REF y DESC
    seis_meses= longitudCols- prom_meses1
    doce_meses= longitudCols-prom_meses2

    df6_meses= df.loc[:,list(columnas[seis_meses:longitudCols])]
    df12_meses= df.loc[:,list(columnas[doce_meses:longitudCols])]
    
    
    #MEDIA DE TODOS LOS MESES SELECCIONADOS
    df['Promedio General']=df.mean(axis=1)
    df.sort_values('Promedio General' , inplace=True ,ascending=False )
    
    
    #MEDIA DE LOS ULTIMOS 6 MESES SI CUBRE LOS 6 MESES
    if len(columnas) >= 8:
        df['Promedio Ult ' + str(prom_meses1) +' Meses']=df6_meses.mean(axis=1)
        df['Desv. Est. Ult ' + str(prom_meses1) +' Meses'] = df6_meses.std(axis= 1)
        df = df.assign(coef_var = lambda x:(x['Desv. Est. Ult ' + str(prom_meses1) +' Meses'] / x['Promedio Ult ' + str(prom_meses1) +' Meses'])) 
 
    if len(columnas) >=14:
        df['Promedio Ult ' + str(prom_meses2) +' Meses']=df12_meses.mean(axis=1)
        df['Desv. Est. Ult ' + str(prom_meses2) +' Meses'] = df12_meses.std(axis= 1)
        df = df.assign(coef_var2 = lambda x:(x['Desv. Est. Ult ' + str(prom_meses2) +' Meses'] / x['Promedio Ult ' + str(prom_meses2) +' Meses'])) 
 

    return df

'TEXTO CORREO SIMS'
def textoCorreoSIMS(mes , ano , radios, t470co , xt185 , repetidoras ,baterias , TotalSIMS , factAyE):
    textoCorreo= '''
Buenas Tardes Ana / Johanna , cordial saludo.

Anexo informes SIMS del mes de '''+mes.upper()+''' del '''+ano+''', correspondiente a baterías, radios y A&E, con totales de '''+str(baterias)+''' baterías, '''+str(radios)+''' radios, más '''+str(t470co)+''' T470 CO, '''+str(xt185)+''' XT185 y  '''+str(repetidoras)+''' repetidoras para un total general mes de '''+str(TotalSIMS)+''' y un total de A&E de   COP $'''+str(factAyE)+'''.

Anexo Informe provisto por Motorola.

Atento a cualquier inquietud y/o requerimiento adicional,
    '''
    return textoCorreo

"-------------------------------------------------------------------------------------------------------------------"
#CAMBIUM
def Filtro_CAMBIUM(df):
    df= df[df['Categoría de producto'].str.contains("CAMBIUM|TP-LINK|SIKLU|EERO MESH|ANTENAS SANNY TELECOM")]
    return  df

def textoCorreoCambium(dia1 , dia2 , mes1, mes2, anio , longitud):
    if dia2==0 and mes2==0:
    #INFORMES DIARIOS
        if longitud > 0:
            textoCorreo='''
Buenos dias, cordial saludo

Envío SELL-THROUGH CAMBIUM , TP-LINK , EERO Y SIKLU del '''+dia1+''' de '''+mes1+''' del '''+anio+'''
        
PEGAR RECORTE

Quedo atento a tus comentarios, saludos.
'''
        else:
            textoCorreo='''
Buenos dias, cordial saludo

informo que no se reportan ventas de equipos CAMBIUM , TP-LINK , EERO Y SIKLU el '''+dia1+''' de '''+mes1+''' del '''+anio+'''

Quedo atento a tus comentarios, saludos.
'''
        return textoCorreo
    
    #INFORME SEMANAL Y ACUMULADO
    elif dia1 != '' and dia2 != '':
        if mes1 == mes2:
            #MES IGUAL
            textoCorreo='''
Buenos dias, cordial saludo

Envío SELL-THROUGH CAMBIUM , TP-LINK , EERO Y SIKLU del '''+dia1+''' al '''+dia2+''' de '''+mes1+''' del '''+anio+'''
        
PEGAR RECORTE

Quedo atento a tus comentarios, saludos.
'''
            return textoCorreo
        else:
            #MESES DISTINTOS
            textoCorreo='''
Buenos dias, cordial saludo

Envío SELL-THROUGH CAMBIUM , TP-LINK , EERO Y SIKLU del '''+dia1+''' de '''+mes1+''' al '''+dia2+''' de '''+mes2+''' del '''+anio+'''
        
PEGAR RECORTE

Quedo atento a tus comentarios, saludos.
'''
            return textoCorreo



'---------------------------------------------------------------------------------------------'
#HUAWEI
def Filtro_Onts_HUAWEI(df):
    #LIMPIEZA DE LOS DATOS
    df= df[df['Categoría de producto'].str.contains("ONTS HUAWEI")]
    df= df[~df['Producto'].str.contains("EG8245W5-6T-50084187|EG8245W5-6T-50084676")]
    return  df

def GenerarTablaCantsOnts(df_filt , diccionario , contador):

    df_auxiliar= pd.pivot_table(df_filt, index=['CLIENTE'] , values=["Cantidad de factura"], columns=['Producto'] , aggfunc=['sum'],margins= True )
    df_auxiliar.sort_values(by=('sum', 'Cantidad de factura' , 'All'), ascending=False,inplace=True)
    #print(df_auxiliar)

    #Reemplaza los valores de NaN por 0
    df_auxiliar = df_auxiliar.fillna(0)
    df_auxiliar.rename(index={"All":"Total General."} ,  inplace=True) #REEMPLAZA EL NOMBRE ALL POR DEFECTO POR TOTAL GENERAL

    #GENERAR TABLA FINAL DE CANTIDADES
    #OBTIENE LAS COLUMNAS
    columnastabla1= list(df_auxiliar.columns.get_level_values(2))
    #print(columnastabla1)
    #OBTIENE LAS FILAS
    filastabla1= df_auxiliar.values
    tabla1= pd.DataFrame(filastabla1 , columns=  columnastabla1)
    tabla1.insert(0, 'CLIENTE', list(df_auxiliar.index) )
    totales= tabla1[tabla1["CLIENTE"]=='Total General.']
    indice= tabla1[tabla1["CLIENTE"]=='Total General.'].index
    tabla1= tabla1.drop(indice)
    tabla1.sort_values('All' , ascending=False , inplace=True)
    tabla1= tabla1.append(totales , ignore_index=True)

    #AGREGAR LOS VALORES TOTALES DIARIOS Y ACUMULADOS PARA EL TEXTO DE VISUALIZACION
    valores = df_auxiliar.loc['Total General.'].values.copy()
    #print(valores)  
    for i in range(len(columnastabla1)):
        if contador == 1:
            dato= {'diario': valores[i] }
        elif contador == 2:
            dato= {'semanal': valores[i] }
        
        diccionario[columnastabla1[i]].append(dato)
    #print(diccionario)
    return tabla1

def GenerarAcumuladoMes_HUAWEI(df_filt2 , mesTemp , ano , diccionario):
    df_acumulado= pd.pivot_table(df_filt2, index=['Empleado responsable'] , values=["Cantidad de factura"], aggfunc=['sum'])

    #GENERAR TABLA FINAL DE CANTIDADES POR EMPLEADO
    df_ac_Final= pd.DataFrame(columns=['EMPLEADO RESPONSABLE' , 'TOTAL ONTs ' + mesTemp.upper() + ' ' + ano])

    df_ac_Final['EMPLEADO RESPONSABLE']= df_acumulado.index
    df_ac_Final['TOTAL ONTs ' + mesTemp.upper() + ' ' + ano] = df_acumulado.values
    #ORDENAR DE MAYOR A MENOR
    df_ac_Final.sort_values('TOTAL ONTs ' + mesTemp.upper() + ' ' + ano , ascending=False , inplace=True)

    total= df_ac_Final['TOTAL ONTs ' + mesTemp.upper() + ' ' + ano].sum()
    tot_gen= {'EMPLEADO RESPONSABLE': 'Total General.' , 'TOTAL ONTs ' + mesTemp.upper() + ' ' + ano : total}
    df_ac_Final= df_ac_Final.append(tot_gen , ignore_index=True)
        
    #GENERAR TABLA AC. MES POR CLIENTE
    df_Cliente = GenerarTablaCantsOnts(df_filt2 , diccionario , contador=2)
    
    return df_ac_Final , df_Cliente

def TextoCorreoONTs_HUAWEI(dia , mes , ano , diccionario):
    texto=''' 
Buenos días, cordial saludo.

El '''+dia+''' de '''+mes+''' del '''+ano+''' se vendieron '''+ str(diccionario['All'][0]['diario']) +''' unds de ONTs.'''
    #print(texto)
    del diccionario['All']

    texto2=''

    for ref in diccionario:
        
        if len(diccionario[ref]) == 2:
            texto2= texto2 + '\n' +'''-'''+ref+''' : '''+str(diccionario[ref][0]['diario'])+''' unidades, para un total de '''+str(diccionario[ref][1]['semanal'])+''' unidades para el mes de '''+mes.capitalize()+''' '''
        elif len(diccionario[ref]) == 1:
            texto2=texto2 + '\n' +'''-'''+ref+''' : 0 unidades, para un total de '''+str(diccionario[ref][0]['semanal'])+''' unidades para el mes de '''+mes.capitalize()+''' '''

    return (texto + '\n' + texto2)


#AIDC
def filtro_AIDC(df, filtro , categoriasAIDC):
    if filtro == 'TODAS LAS CATEGORIAS':
        df = df[df['Categoría de producto'].str.contains('|'.join(categoriasAIDC))]
    else:
        df = df[df['Categoría de producto'].str.contains(filtro)]
    return  df

def Tabla_cants_AIDC(df):
    #TABLA POR PRODUCTO(TEXTO)
    tablaProd= pd.pivot_table(df, columns=['Factura (Mes natural)' ] , index= [ 'Producto' , 'Producto (Texto)' , 'Categoría de producto'] , values='Cantidad de factura' , aggfunc='sum' , margins=True )
    #Reemplaza los valores de NaN por 0
    tablaProd = tablaProd.fillna(0)
    tablaProd.rename(index={"All":"Total General"} ,  inplace=True) #REEMPLAZA EL NOMBRE ALL POR DEFECTO POR TOTAL GENERAL
    tablaProd.rename(columns={'All':'Total General.'}, inplace=True)
    
    #Genera tabla ordenada
    #OBTIENE LAS COLUMNAS
    columnastabla1= list(tablaProd.columns)
    tabla3= pd.DataFrame( tablaProd.values , columns= columnastabla1 )
    tabla3.insert(0, 'DESCRIPCION', list(tablaProd.index.get_level_values(1)) )
    tabla3.insert(0, 'PRODUCTO', list(tablaProd.index.get_level_values(0)) )
    tabla3.insert(0, 'CATEGORIA DE PRODUCTO', list(tablaProd.index.get_level_values(2)) )
    totales3= tabla3[tabla3["PRODUCTO"]=='Total General']
    indice3= tabla3[tabla3["PRODUCTO"]=='Total General'].index
    tabla3= tabla3.drop(indice3)
    tabla3.sort_values('Total General.' ,ascending=False , inplace=True)
    tabla3= tabla3.append(totales3 , ignore_index=True)

    return tabla3