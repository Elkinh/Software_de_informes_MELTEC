import pandas as pd
import sys

pd.set_option('mode.chained_assignment', None)
#Los meses que cubre cada Q
Qs= { "1": ["Enero" , "Febrero" , "Marzo"] , "2": ["Abril" , "Mayo" , "Junio"] , "3": ["Julio", "Agosto" , "Septiembre"] , "4": ["Octubre" , "Noviembre" , "Diciembre"]}

def Seleccion_Anio(filename, ano):
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

def cargar_metas_y_comerciales_motorola():
    #LEER EXCEL
    archivo_Excel= pd.read_excel('./Metas Motorola.xlsx' , sheet_name= "METAS" , header = 0)
    return archivo_Excel

def seleccion_Q_actual(mes , Qs):
    for i in range(1,5):
        if mes in Qs[i]:
            return str(i)
        

def agregar_filtros( df , df_metas , Q_number , mes_Actual , actual_Q):
    comerciales=[]
    for asesor in df_metas["ASESOR"]:
        if asesor != "TOTAL GENERAL":
            comerciales.append(asesor)

    df_filt=pd.DataFrame()
    df_filt = df[df['Empleado responsable'].isin(comerciales)]

    #Creacion de la tabla a Exportar a Excel
    fin_titulo= mes_Actual.upper() + ' 2022'
    df_tabla= pd.DataFrame(columns=['ASESOR',('VALOR VENTA ' + fin_titulo),('META MES ' + fin_titulo), ('PORCENTAJE CUMPLIDO ' + fin_titulo), ('VENTA Q' + Q_number +' AÑO 2022') , ('META Q' + Q_number +' AÑO 2022' ), ('PORCENTAJE DEL Q'+ Q_number+' CUMPLIDO AÑO 2022'), ('VENTA ACUMULADA AÑO 2022'), ('META ANUAL') , ('PORCENTAJE CUMPLIDO META ANUAL') ])
    df_tabla['ASESOR']= comerciales

    '''FILTRO COMERCIALES FOCO MOTOROLA'''
    #Filtro de mes actual
    mes_Act= pd.DataFrame(df_filt[df_filt['Factura (Mes natural)']== mes_Actual])
    df_auxiliar= pd.pivot_table(mes_Act, index=["Empleado responsable"], values=["Valor neto facturado"], aggfunc=['sum'])
    
    for i in range(len(df_auxiliar)):
        for j in range(len(df_tabla)):
            if df_auxiliar.index[i] == df_tabla["ASESOR"].iloc[j] and df_auxiliar.index[i] != "DORA PATRICIA AYALA SUAREZ" and df_auxiliar.index[i] != "ANDRES FELIPE MUÑOZ PORTILLO" and df_auxiliar.index[i] != "NATALIA ALVAREZ PINO":
                df_tabla['VALOR VENTA ' + fin_titulo].iloc[j] = df_auxiliar.iloc[i,0]
    
    #Filtro Q en curso
    Q_actual = pd.DataFrame(df[df['Factura (Mes natural)'].isin(actual_Q)])
    df_auxiliar= pd.pivot_table(Q_actual, index=["Empleado responsable"], values=["Valor neto facturado"], aggfunc=['sum'])
    
    for i in range(len(df_auxiliar)):
        for j in range(len(df_tabla)):
            if df_auxiliar.index[i] == df_tabla["ASESOR"].iloc[j] and df_auxiliar.index[i] != "DORA PATRICIA AYALA SUAREZ" and df_auxiliar.index[i] != "ANDRES FELIPE MUÑOZ PORTILLO" and df_auxiliar.index[i] != "NATALIA ALVAREZ PINO":
                df_tabla['VENTA Q' + Q_number +' AÑO 2022'].iloc[j] = df_auxiliar.iloc[i,0]
    
    #Filtro Acumulado AÑO
    df_auxiliar_Año= pd.pivot_table(df_filt, index=["Empleado responsable"], values=["Valor neto facturado"], aggfunc=['sum'])
    for i in range(len(df_auxiliar_Año)):
        for j in range(len(df_tabla)):
            if df_auxiliar_Año.index[i] == df_tabla["ASESOR"].iloc[j] and df_auxiliar_Año.index[i] != "DORA PATRICIA AYALA SUAREZ" and df_auxiliar_Año.index[i] != "ANDRES FELIPE MUÑOZ PORTILLO" and df_auxiliar_Año.index[i] != "NATALIA ALVAREZ PINO":
                df_tabla['VENTA ACUMULADA AÑO 2022'].iloc[j] = df_auxiliar_Año.iloc[i,0]
    
    '''FILTRO COMERCIALES SUCURSALES'''
    #Filtro mes actual
    mes_Act= mes_Act[mes_Act['Categoría de producto'].str.contains("MOTOROLA|VERTEX")]
    df_auxiliar= pd.pivot_table(mes_Act, index=["Empleado responsable"], values=["Valor neto facturado"], aggfunc=['sum'])
    for i in range(len(df_auxiliar)):
        for j in range(len(df_tabla)):
            if df_auxiliar.index[i] == df_tabla["ASESOR"].iloc[j] and df_auxiliar.index[i] == "DORA PATRICIA AYALA SUAREZ":
                df_tabla['VALOR VENTA ' + fin_titulo].iloc[j] = df_auxiliar.iloc[i,0]
            elif df_auxiliar.index[i] == df_tabla["ASESOR"].iloc[j] and df_auxiliar.index[i] == "ANDRES FELIPE MUÑOZ PORTILLO":
                df_tabla['VALOR VENTA ' + fin_titulo].iloc[j] = df_auxiliar.iloc[i,0]
            elif df_auxiliar.index[i] == df_tabla["ASESOR"].iloc[j] and df_auxiliar.index[i] == "NATALIA ALVAREZ PINO":
                df_tabla['VALOR VENTA ' + fin_titulo].iloc[j] = df_auxiliar.iloc[i,0]
    
    #Filtro Q en curso
    Q_actual = Q_actual[Q_actual['Categoría de producto'].str.contains("MOTOROLA|VERTEX")]
    df_auxiliar= pd.pivot_table(Q_actual, index=["Empleado responsable"], values=["Valor neto facturado"], aggfunc=['sum'])
    
    for i in range(len(df_auxiliar)):
        for j in range(len(df_tabla)):
            if df_auxiliar.index[i] == df_tabla["ASESOR"].iloc[j] and df_auxiliar.index[i] == "DORA PATRICIA AYALA SUAREZ":
                df_tabla['VENTA Q' + Q_number +' AÑO 2022'].iloc[j] = df_auxiliar.iloc[i,0]
            elif df_auxiliar.index[i] == df_tabla["ASESOR"].iloc[j] and df_auxiliar.index[i] == "ANDRES FELIPE MUÑOZ PORTILLO":
                df_tabla['VENTA Q' + Q_number +' AÑO 2022'].iloc[j] = df_auxiliar.iloc[i,0]
            elif df_auxiliar.index[i] == df_tabla["ASESOR"].iloc[j] and df_auxiliar.index[i] == "NATALIA ALVAREZ PINO":
                df_tabla['VENTA Q' + Q_number +' AÑO 2022'].iloc[j] = df_auxiliar.iloc[i,0]

    #Filtro Acumulado AÑO
    df_filt= df_filt[df_filt['Categoría de producto'].str.contains("MOTOROLA|VERTEX")]
    df_filt= pd.pivot_table(df_filt, index=["Empleado responsable"], values=["Valor neto facturado"], aggfunc=['sum'])
    for i in range(len(df_filt)):
        for j in range(len(df_tabla)):
            if df_filt.index[i] == df_tabla["ASESOR"].iloc[j] and df_filt.index[i] == "DORA PATRICIA AYALA SUAREZ":
                df_tabla['VENTA ACUMULADA AÑO 2022'].iloc[j] = df_filt.iloc[i,0]
            elif df_filt.index[i] == df_tabla["ASESOR"].iloc[j] and df_filt.index[i] == "ANDRES FELIPE MUÑOZ PORTILLO":
                df_tabla['VENTA ACUMULADA AÑO 2022'].iloc[j] = df_filt.iloc[i,0]
            elif df_filt.index[i] == df_tabla["ASESOR"].iloc[j] and df_filt.index[i] == "NATALIA ALVAREZ PINO":
                df_tabla['VENTA ACUMULADA AÑO 2022'].iloc[j] = df_filt.iloc[i,0]
    
    #Reemplaza los valores de NaN por 0
    df_tabla = df_tabla.fillna(0)
    #Se ordena la tabla por Mes actual de mayor venta de los comerciales
    df_tabla= df_tabla.sort_values('VALOR VENTA ' + fin_titulo , ascending=False)
            
    #Sumatoria
    total_Gen_mes= df_tabla['VALOR VENTA ' + fin_titulo].sum()
    total_Gen_Q= df_tabla['VENTA Q' + Q_number +' AÑO 2022'].sum()
    total_Gen_Año= df_tabla['VENTA ACUMULADA AÑO 2022'].sum()

    #Se agregan los totales generales de cada columna
    totales= {'ASESOR': "Total General" , 'VALOR VENTA ' + fin_titulo : total_Gen_mes ,'VENTA Q' + Q_number +' AÑO 2022' : total_Gen_Q ,'VENTA ACUMULADA AÑO 2022': total_Gen_Año }
    df_tabla= df_tabla.append(totales , ignore_index=True)

    return df_tabla , fin_titulo

def agregar_Metas(df_tabla, df_metas , Number_Q , fin_titulo):
    for i in range(len(df_tabla)):
        for j in range(len(df_metas)):
            if df_tabla["ASESOR"].iloc[i] != "TOTAL GENERAL":
                if df_tabla["ASESOR"].iloc[i] == df_metas["ASESOR"].iloc[j]:
                    df_tabla['META MES ' + fin_titulo].iloc[i] = df_metas["META MES A MES Q" + Number_Q].iloc[j].copy()
                    df_tabla['META Q' + Number_Q +' AÑO 2022'].iloc[i] = df_metas["META Q" + Number_Q].iloc[j]
                    df_tabla['META ANUAL'].iloc[i] = df_metas["META 2022"].iloc[j]
    
    #Sumatoria
    total_Gen_meta_mes= df_tabla['META MES ' + fin_titulo].sum()
    total_Gen_meta_Q= df_tabla['META Q' + Number_Q +' AÑO 2022'].sum()
    total_Gen_meta_Año= df_tabla['META ANUAL'].sum()

    df_tabla['META MES ' + fin_titulo].iloc[len(df_tabla)-1] = total_Gen_meta_mes
    df_tabla['META Q' + Number_Q +' AÑO 2022'].iloc[len(df_tabla)-1]=total_Gen_meta_Q
    df_tabla['META ANUAL'].iloc[len(df_tabla)-1]=total_Gen_meta_Año

    return df_tabla

def agregar_Porcentajes(df_tabla, Q_number, fin_titulo):
    for i in range(len(df_tabla)):
        #Porcentaje cumplido MES AÑO 2022
        df_tabla['PORCENTAJE CUMPLIDO ' + fin_titulo].iloc[i] = df_tabla['VALOR VENTA ' + fin_titulo].iloc[i] / df_tabla['META MES ' + fin_titulo].iloc[i]
        #Porcentaje cumplido Q EN CURSO AÑO 2022
        df_tabla['PORCENTAJE DEL Q'+ Q_number+' CUMPLIDO AÑO 2022'].iloc[i] = df_tabla['VENTA Q' + Q_number +' AÑO 2022'].iloc[i] / df_tabla['META Q' + Q_number +' AÑO 2022'].iloc[i]
        #Porcentaje cumplido AÑO 2022
        df_tabla['PORCENTAJE CUMPLIDO META ANUAL'].iloc[i] = df_tabla['VENTA ACUMULADA AÑO 2022'].iloc[i] / df_tabla['META ANUAL'].iloc[i]
    return df_tabla

def Ventas_Motorola_Vertex(data , df_metas):
    #FILTROS DEL INFORME
    data= data[data['Categoría de producto'].str.contains("MOTOROLA|VERTEX")]
    data= pd.pivot_table(data, index=["Empleado responsable"], values=["Valor neto facturado"], aggfunc=['sum'])

    df_tabla2= pd.DataFrame(columns=['ASESOR','VENTA AÑO 2022 MOTOROLA', 'META MOTOROLA ANUAL' , 'PORCENTAJE CUMPLIDO'])
    df_tabla2["ASESOR"] = data.index
    df_tabla2["VENTA AÑO 2022 MOTOROLA"] = data.values

    #Ingresa la Meta de cada comercial de motorola y sucursales
    for i in range(len(df_tabla2)):
        for j in range(len(df_metas)):
            if df_tabla2["ASESOR"].iloc[i] != "TOTAL GENERAL":
                if df_tabla2["ASESOR"].iloc[i] == df_metas["ASESOR"].iloc[j]:
                    df_tabla2['META MOTOROLA ANUAL'].iloc[i] = df_metas["META 2022"].iloc[j]
     
    df_tabla2 = df_tabla2.sort_values('VENTA AÑO 2022 MOTOROLA',ascending=False)
    total_general= {"ASESOR": "Total general" , "VENTA AÑO 2022 MOTOROLA" : df_tabla2["VENTA AÑO 2022 MOTOROLA"].sum() ,"META MOTOROLA ANUAL" :df_tabla2["META MOTOROLA ANUAL"].sum() }
    df_tabla2= df_tabla2.append(total_general, ignore_index=True)

    #Ingresa el porcentaje cumplido
    for i in range(len(df_tabla2)):
        df_tabla2['PORCENTAJE CUMPLIDO'].iloc[i] = df_tabla2["VENTA AÑO 2022 MOTOROLA"].iloc[i] / df_tabla2['META MOTOROLA ANUAL'].iloc[i]
       
    return df_tabla2


'-------------------------------------------------------------------------------------------------------'
#AIDC
def tabla_AIDC(df):
    
    #TABLA POR EMPLEADO RESPONSABLE
    tabla_aux= pd.pivot_table(df, columns=['Factura (Mes natural)'] , index= 'Empleado responsable' , values='Valor neto facturado' , aggfunc='sum' , margins=True )
    #Reemplaza los valores de NaN por 0
    tabla_aux = tabla_aux.fillna(0)
    tabla_aux.rename(index={"All":"Total General"} ,  inplace=True) #REEMPLAZA EL NOMBRE ALL POR DEFECTO POR TOTAL GENERAL
    tabla_aux.rename(columns={'All':'Total General.'}, inplace=True)
    
    #Genera tabla ordenada
    tabla= pd.DataFrame( tabla_aux.values , columns= tabla_aux.columns )
    tabla.insert(0 , 'ASESOR', list(tabla_aux.index))
    totales= tabla[tabla["ASESOR"]=='Total General']
    indice= tabla[tabla["ASESOR"]=='Total General'].index
    tabla= tabla.drop(indice)
    tabla.sort_values('Total General.' ,ascending=False , inplace=True)
    tabla= tabla.append(totales , ignore_index=True)
    
    
    #TABLA POR CAT PRODUCTO
    tablaCat= pd.pivot_table(df, columns='Factura (Mes natural)' , index= 'Categoría de producto' , values='Valor neto facturado' , aggfunc='sum' , margins=True )
    #Reemplaza los valores de NaN por 0
    tablaCat = tablaCat.fillna(0)
    tablaCat.rename(index={"All":"Total General"} ,  inplace=True) #REEMPLAZA EL NOMBRE ALL POR DEFECTO POR TOTAL GENERAL
    tablaCat.rename(columns={'All':'Total General.'}, inplace=True)
    
    #Genera tabla ordenada
    tabla2= pd.DataFrame( tablaCat.values , columns= tablaCat.columns )
    tabla2.insert(0 , 'CATEGORIA DE PRODUCTO', list(tablaCat.index))
    totales2= tabla2[tabla2["CATEGORIA DE PRODUCTO"]=='Total General']
    indice2= tabla2[tabla2["CATEGORIA DE PRODUCTO"]=='Total General'].index
    tabla2= tabla2.drop(indice2)
    tabla2.sort_values('Total General.' ,ascending=False , inplace=True)
    tabla2= tabla2.append(totales2 , ignore_index=True)

    #TABLA POR CLIENTE
    tablaCliente= pd.pivot_table(df, columns='Factura (Mes natural)' , index= 'Factura (Cliente)' , values='Valor neto facturado' , aggfunc='sum' , margins=True )
    #Reemplaza los valores de NaN por 0
    tablaCliente = tablaCliente.fillna(0)
    tablaCliente.rename(index={"All":"Total General"} ,  inplace=True) #REEMPLAZA EL NOMBRE ALL POR DEFECTO POR TOTAL GENERAL
    tablaCliente.rename(columns={'All':'Total General.'}, inplace=True)
    
    #Genera tabla ordenada
    tabla3= pd.DataFrame( tablaCliente.values , columns= tablaCliente.columns )
    tabla3.insert(0 , 'CLIENTE', list(tablaCliente.index))
    totales3= tabla3[tabla3["CLIENTE"]=='Total General']
    indice3= tabla3[tabla3["CLIENTE"]=='Total General'].index
    tabla3= tabla3.drop(indice3)
    tabla3.sort_values('Total General.' ,ascending=False , inplace=True)
    tabla3= tabla3.append(totales3 , ignore_index=True)

    return tabla , tabla2 , tabla3

