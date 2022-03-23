import Func_Informes_Meltec_GUI_POO_CNTS as func
import Func_Informes_Meltec_GUI_POO_FACT as fact
import pandas as pd
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from collections import defaultdict

Qs= { 1: ["Enero" , "Febrero" , "Marzo"] , 2: ["Abril" , "Mayo" , "Junio"] , 3: ["Julio", "Agosto" , "Septiembre"] , 4: ["Octubre" , "Noviembre" , "Diciembre"]}
categoriasAIDC= ['HONEYWELL','ZEBRA','ULEFONE','BIXOLON','EQUIPOS EN ALQUILER','PROYECTOS MOVILIDAD','CINTAS CONSUMIBLES','CABLES CONSUMIBLES','ESTUCHES CONSUMIBLES','ETIQUETAS CONSUMIBLES','SOTI','IGS','SEWOO','LIMASOFT''PANASONIC','RUGGEAR','CATERPILAR','BEMATECH','CIPHERLAB','CYRUS','KYOCERA','NEWLAND','OPTICON','HYT','SONY','TODAS LAS CATEGORIAS']
    
class APP(Tk):
    def __init__(self,*args,**kwargs):
        super().__init__(*args,**kwargs)
        self.title("Informes Meltec")
        
        #MENU SUPERIOR
        self.MenuSuperior()
        #FRAME DEL LOGO
        self.F_1 = FrameLogo(self)
        self.F_1.config(width= "300" , height="170") 
        self.F_1.pack(fill="both" , expand="True" )#Ajustar el Frame dentro de la raiz
        
        self.contenedor_principal= Frame()
        self.contenedor_principal.pack()

        self.testigo=0
        self.seleccion=''

        #Diccionario donde se guardaran las 2 vistas
        self.todos_los_frames = dict()

        for F in (Frame_Inicio, Informes_Diarios, Informes_Semanales, Informes_Acumulados, Informe_Facturacion , Informe_Q_cants , Informe_SIMS , Informe_Mes_Mes_Motorola , INFORMES_AIDC):
            frame = F(self.contenedor_principal , self)
            self.todos_los_frames[F] = frame
            frame.grid(row = 0, column = 0, sticky = "nsew")
        self.show_frame(Frame_Inicio)
        
    def show_frame(self,contenedor_llamado):
        frame = self.todos_los_frames[contenedor_llamado]
        frame.tkraise()

    def MenuSuperior(self):
        self.barraMenu= Menu(self)# Creacion
        self.config(menu= self.barraMenu)#Asignacion
        #PESTAÑAS
        self.archivoMenu= Menu(self.barraMenu , tearoff=0) #tearoff= 0 --> Elimina un espacio en el submenu que no se ve bien
        self.barraMenu.add_cascade(label="Archivo", menu=self.archivoMenu)
        #AGREGAR SUBMENUS
        self.archivoMenu.add_command(label="Cargar Archivo" , command= self.abreFichero)
        self.archivoMenu.add_command(label="Salir" , command= self.salir)

    def abreFichero(self):
        #LEER EXCEL SAP
        self.fichero= func.cargar_Informe_SAP()
        
        
    def salir(self):
        valor= messagebox.askquestion("Salir","Desea salir de la aplicacion?") #devuelve yes / no
        if valor == 'yes':
            self.destroy() #Cierra el programa

    def getFichero(self):
        try:
            return self.fichero
        except AttributeError:
            messagebox.showerror("Error","No se ha cargado el archivo, cargarlo por favor")
            self.fichero= filedialog.askopenfilename(title="Abrir" , filetypes=(("Ficheros de Excel" , "*.xlsx"),("Ficheros de texto", "*.txt"),("Todos los ficheros","*.*")) )
            return self.fichero
    def setTestigo(self , testigo):
        self.testigo= testigo
    
    def getTestigo(self):
        return self.testigo

    def setSeleccion(self , seleccion): 
        self.seleccion = seleccion
    
    def getSeleccion(self):
        return self.seleccion

class FrameLogo(Frame):
    '''IMPLEMENTACION DE LOGO MELTEC'''
    def __init__(self, container,*args, **kwargs):
        #Para inicializar Frame y que se entienda la clase como un Frame, tambien se inicializa desde la herencia:
        #NOTA: NO olvidar pasar el parametro tambien, de lo contrario, no se comprende donde localizar el Frame creado
        #Luego de hacer esto, recordar que "self" es equivalente a decir el Frame_1, que es en donde agregaremos widgets
        super().__init__(container, *args, **kwargs)
        self.miImagen= PhotoImage(file= "./imagenes/logo_Meltec.png")
        self.miImagen= self.miImagen.subsample(4)
        L_1=Label(self, image=self.miImagen) #Imagen
        L_1.place(x=90 , y= 20)
        

class FrameMen(Frame):
    '''NUEVO FRAME PARA HACER MENÚ DESPLEGABLE'''
    varEnviar= 0
    def __init__(self, container , controller,*args, **kwargs): 
        
        super().__init__(container, *args, **kwargs)
        self.varOpcion= IntVar()
        self.Seleccion= StringVar()
        self.Seleccion2=StringVar()
        self.Seleccion3=StringVar()
        self.Seleccion4=StringVar()

        self.OpcionesMotorola= ['Informe Diario Cants','Informe Semanal Cants','Informe Ac/Men Cants' ,'Informe Facturacion','Informe por Q Cants', 'Informes por Mes Cants','Informe SIMS']
        self.OpcionesCambium=['Informe Diario Cants','Informe Semanal Cants','Informe Ac/Men Cants']
        self.OpcionesHuawei= ['Informe Diario Cants']
        self.OpcionesAIDC= ['Informe Cants' , 'Informe Facturacion']

        self.Motorola= Radiobutton(self , text= "MOTOROLA" , variable= self.varOpcion, value=1, command= self.MenuMotorola)
        self.Motorola.grid(row=0 , column=0 ,padx=10 , pady=5, sticky='w')
        self.Motorola.config(justify='right')

        self.Cambium= Radiobutton(self , text= "CAMBIUM" , variable=self.varOpcion, value=2 , command=self.MenuCambium)
        self.Cambium.grid(row=1 , column=0 ,padx=10 , pady=5, sticky='w')
        self.Cambium.config(justify='right')

        self.Huawei= Radiobutton(self , text= "HUAWEI" , variable=self.varOpcion, value=3 , command=self.MenuHuawei)
        self.Huawei.grid(row=2 , column=0 ,padx=10 , pady=5, sticky='w')
        self.Huawei.config(justify='right')

        self.AIDC= Radiobutton(self , text= "AIDC" , variable=self.varOpcion, value=4 , command=self.MenuAIDC)
        self.AIDC.grid(row=3 , column=0 ,padx=10 , pady=5, sticky='w')
        self.AIDC.config(justify='right')

    #Funciones de activacion del menu desplegable
    def MenuMotorola(self):
        global opcionMot
        self.Seleccion.set("")
        self.Seleccion2.set("")
        self.Seleccion3.set("")
        self.Seleccion4.set("")
        opcionMot= OptionMenu(self , self.Seleccion , *self.OpcionesMotorola).grid(row=0, column=1  , padx=5 , pady=5) 
        
    def MenuCambium(self):
        global opcionCambium
        self.Seleccion.set("")
        self.Seleccion2.set("")
        self.Seleccion3.set("")
        self.Seleccion4.set("")
        opcionCambium= OptionMenu(self , self.Seleccion2 , *self.OpcionesCambium).grid( row=1 ,column=1 , padx=5 , pady=5)

    def MenuHuawei(self):
        global opcionHuawei
        self.Seleccion.set("")
        self.Seleccion2.set("")
        self.Seleccion3.set("")
        self.Seleccion4.set("")
        opcionHuawei= OptionMenu(self , self.Seleccion3 , *self.OpcionesHuawei).grid( row=2 ,column=1 , padx=5 , pady=5)

    def MenuAIDC(self):
        global opcionAIDC
        self.Seleccion.set("")
        self.Seleccion2.set("")
        self.Seleccion3.set("")
        self.Seleccion4.set("")
        opcionAIDC= OptionMenu(self , self.Seleccion4 , *self.OpcionesAIDC).grid( row=3 ,column=1 , padx=5 , pady=5)

    def getVaroption(self):
        return self.varOpcion.get()

class Frame_Botones(Frame):
    def __init__(self, container , controller , valor1, seleccion1 , seleccion2 , seleccion3 , seleccion4 ,*args, **kwargs): 
        super().__init__(container, *args, **kwargs)
        self.botonenvio = Button(self , text= "Aceptar" , command=lambda: seleccionInforme()) 
        self.botonenvio.grid(row=0 ,column=0,padx=10 , pady=5)

        self.botonCargar = Button(self , text= "Cargar Archivo", command=lambda: controller.abreFichero() )#, command=abreFichero) 
        self.botonCargar.grid(row=0 ,column=1,padx=10 , pady=5)

        self.botonSalir= Button(self , text="Salir" , command=lambda: controller.salir())#, command=salir)
        self.botonSalir.grid(row=0 ,column=2,padx=10 , pady=5)

        def seleccionInforme():
            'Informe Diario Cants','Informe Semanal Cants','Informe Ac/Men Cants' ,'Informe Facturacion'
            #MOTOROLA
            if valor1.get() == 1:
                controller.setTestigo(1)
                if seleccion1.get() == 'Informe Diario Cants':
                    controller.show_frame(Informes_Diarios)
                elif seleccion1.get() == 'Informe Semanal Cants':
                    controller.show_frame(Informes_Semanales)
                elif seleccion1.get() == 'Informe Ac/Men Cants':
                    controller.show_frame(Informes_Acumulados)
                elif seleccion1.get() == 'Informe Facturacion':
                    controller.show_frame(Informe_Facturacion)
                elif seleccion1.get() == 'Informes por Mes Cants':
                    controller.show_frame(Informe_Mes_Mes_Motorola)
                elif seleccion1.get() == 'Informe por Q Cants':
                    controller.show_frame(Informe_Q_cants)
                elif seleccion1.get() == 'Informe SIMS':
                    controller.show_frame(Informe_SIMS)
                
            #CAMBIUM
            elif valor1.get() ==2:
                controller.setTestigo(2)
                if seleccion2.get() == 'Informe Diario Cants':
                    controller.show_frame(Informes_Diarios)
                elif seleccion2.get() == 'Informe Semanal Cants':
                    controller.show_frame(Informes_Semanales)
                elif seleccion2.get() == 'Informe Ac/Men Cants':
                    controller.show_frame(Informes_Acumulados)

            #HUAWEI
            elif valor1.get() ==3:
                controller.setTestigo(3)
                if seleccion3.get() == 'Informe Diario Cants':
                    controller.show_frame(Informes_Diarios)
            
            #AIDC
            elif valor1.get() == 4:
                controller.setTestigo(4)
                if seleccion4.get() == 'Informe Cants':
                    controller.setSeleccion('Cants')
                    controller.show_frame(INFORMES_AIDC)
                elif seleccion4.get() == 'Informe Facturacion':
                    controller.setSeleccion('Fact')
                    controller.show_frame(INFORMES_AIDC)
            
'''
CLASES CORRESPONDIENTES A LAS VISTAS
'''

class Frame_Inicio(Frame):
    def __init__(self, container, controller ,*args, **kwargs):
        #Para inicializar Frame y que se entienda la clase como un Frame, tambien se inicializa desde la herencia:
        #NOTA: NO olvidar pasar el parametro tambien, de lo contrario, no se comprende donde localizar el Frame creado
        #Luego de hacer esto, recordar que "self" es equivalente a decir el Frame_1, que es en donde agregaremos widgets
        super().__init__( container, *args, **kwargs)
        self.contenedor_principal= FrameMen(self , controller)
        self.contenedor_principal.pack()
        self.contenedor_secundario= Frame_Botones(self , controller , self.contenedor_principal.varOpcion , self.contenedor_principal.Seleccion ,self.contenedor_principal.Seleccion2 , self.contenedor_principal.Seleccion3 , self.contenedor_principal.Seleccion4 )
        self.contenedor_secundario.pack()




class Informes_Diarios(Frame):
    def __init__(self, container , controller , *args, **kwargs):
        super().__init__( container , *args, **kwargs)
        
        #LABELS
        self.labeldia = Label(self, text = "Ingrese el dia en numeros: ").grid(row=0 , column=0 , sticky='e' , padx=5 , pady= 5)
        self.labelmes = Label(self, text = "Ingrese el mes: ").grid(row=1 , column=0 , sticky='e' , padx=5 , pady= 5)
        self.labelAno = Label(self, text = "Año: ").grid(row=2 , column=0 , sticky='e' , padx=5 , pady= 5)

        #DIA
        self.Textodia= Entry(self , width=15)
        self.Textodia.grid(row=0 , column=1 , padx=2 , pady= 2)
        self.Textodia.config(justify= "center")
        #MES
        self.Meses= ['Enero' , 'Febrero' , 'Marzo' , 'Abril' , 'Mayo' , 'Junio' , 'Julio' , 'Agosto' , 'Septiembre' , 'Octubre' , 'Noviembre' , 'Diciembre']
        self.Textomes= StringVar()
        menuMes=OptionMenu(self , self.Textomes , *self.Meses)
        menuMes.grid( row=1 ,column=1 , padx=2 , pady=2 , sticky='e' )
        menuMes.config(width=9 ,justify="center")

        #AÑO
        self.Ano= ['2019' , '2020' , '2021' , '2022']
        self.TextoAno= StringVar()
        menuAno=OptionMenu(self , self.TextoAno , *self.Ano)
        menuAno.grid( row=2 ,column=1 , padx=2 , pady=2 , sticky='e' )
        menuAno.config(width=9 ,justify="center")

        #BOTONES
        self.buttonExample = Button(self, text = "Generar Reporte"  , command= lambda: [self.generar_diario(controller.getFichero() , controller.getTestigo() , self.Textodia.get() , self.Textomes.get() , self.TextoAno.get() ) , self.TextoCorreo()]).grid(row=4 , column=1  , padx=5 , pady= 5 )
        self.buttonAtras = Button(self, text = "Atras" , command= lambda: controller.show_frame(Frame_Inicio)).grid(row=5 , column=0 , sticky='w' , padx=10 , pady= 5)
            
    
    def generar_diario(self, filename,testigo , diaTemp , mesTemp , ano):
        self.pasar=0
        if diaTemp =="" or mesTemp=="" or int(diaTemp)<=0 or int(diaTemp)>31 or ano=="": messagebox.showwarning("Aviso" , "Error en ingreso de datos en los campos designados")

        else:
            self.pasar=1
            if len(diaTemp) == 1:
                if int(diaTemp) > 0 and int(diaTemp) <= 9:
                    diaTemp= "0" + diaTemp
            
            if testigo == 1:
                #Testigo = 1 Indica que corresponde a MOTOROLA
                df_filt= func.limpieza_datos(filename , ano)
                df_filt = func.Filtro_Mot(df_filt)
                df_filt= func.agregar_fecha(df_filt , mesTemp , "diario" , diaTemp , 0)
                #se genera archivo temporal
                df_filt.to_excel("ArchivoTemp.xlsx" ,header= True , index = False, sheet_name="SELL-THROUGH" )
                #se cambian las descripciones
                df_filt2= func.cambiar_descripciones()
                df_filt2= pd.DataFrame(df_filt2)
                
                #Redaccion del Correo
                radios, t470co , xt185 , repetidoras , cantTotal =func.ContarCantidadesMotorola(df_filt2)
                self.texto= func.textoCorreoMot(diaTemp, 0 , mesTemp, 0 , ano , radios, t470co , xt185 , repetidoras , cantTotal)
                
                #DIRECTORIO DE DESCARGA
                os.remove('ArchivoTemp.xlsx') #Elimina el archivo el cual no es necesario
                func.guardar_reporte(df_filt2 ,False , False , 1)
                
            elif testigo ==2:
                #Testigo = 2 Indica que corresponde a CAMBIUM
                df_filt= func.limpieza_datos(filename , ano)
                df_filt= func.Filtro_CAMBIUM(df_filt)
                df_filt= func.agregar_fecha(df_filt , mesTemp , 'diario' , diaTemp , 0)

                #Redaccion del Correo
                self.texto=func.textoCorreoCambium(diaTemp,0, mesTemp,0, ano,  len(df_filt.index))

                #DIRECTORIO DE DESCARGA
                func.guardar_reporte(df_filt , False , False ,1)

            elif testigo==3:
                #Diccionario para registrar los valores para redactar el correo
                DicSumario = defaultdict(list)

                #Testigo = 3 Indica que corresponde a HUAWEI
                df_filt= func.limpieza_datos(filename , ano)
                df_filt= func.Filtro_Onts_HUAWEI(df_filt)
                
                #1 DF
                df_filt1= func.agregar_fecha(df_filt , mesTemp , 'diario' , diaTemp , 0)
                
                if df_filt1.shape[0] == 0 : messagebox.showwarning("Información" , "No existen ventas de ONTs en la fecha escrita")

                else:
                    df_filt1 = df_filt1.rename(columns={'Factura (Cliente)':'CLIENTE'}) #REEMPLAZA EL NOMBRE ORIGINAL POR SOLO CLIENTE
                    
                    #Hoja 1 DIARIO
                    tabla1 = func.GenerarTablaCantsOnts(df_filt1, DicSumario ,contador=1)
                    
                    #INFORME AC/MES AL DIA DEL INFORME
                    df_filt2= func.agregar_fecha(df_filt , mesTemp , mesTemp ,'01',diaTemp)
                    df_filt2 = df_filt2.rename(columns={'Factura (Cliente)':'CLIENTE'}) #REEMPLAZA EL NOMBRE ORIGINAL POR SOLO CLIENTE
                    tabla2 , tabla3= func.GenerarAcumuladoMes_HUAWEI(df_filt2 , mesTemp , ano  , DicSumario)

                    #Redaccion del correo
                    self.texto=func.TextoCorreoONTs_HUAWEI(diaTemp , mesTemp , ano , DicSumario)

                    #DIRECTORIO DE DESCARGA
                    func.guardar_reporte(tabla1 , tabla2 , tabla3 ,3)

    def  TextoCorreo(self):
        if self.pasar==1:
            newWindows = Toplevel(self)
            #TEXTO A COPIAR
            ComentariosLabel= Label(newWindows , text="REDACCION DEL CORREO").grid(row=0 , column=0 , padx=5 , pady=5 )
            textoComentario= Text(newWindows , width= 50 , height= 20) #Diseña cuadro de texto
            textoComentario.grid( row=1, column=0, padx=5 , pady= 5)
            textoComentario.insert(1.0 , self.texto)
            #SCROLLBAR
            scroll= Scrollbar(newWindows, command=textoComentario.yview) #Se crea scroll
            scroll.grid(row=1, column=1, padx=5 , pady= 5 ,  sticky="nsew" ) #Posicionar scroll
            textoComentario.config(yscrollcommand= scroll.set) #Para que el scroll se mueva con el desplazamiento hacia abajo
        

class Informes_Semanales(Frame):
    def __init__(self, container , controller , *args, **kwargs):
        super().__init__( container , *args, **kwargs)
        #DIA 1
        self.labeldia1 = Label(self, text = "Dia inicio:" , width= 10).grid(row=0 , column=0 , sticky='e' , padx=5 , pady= 5)
        self.Textodia1= Entry(self, width= 10)
        self.Textodia1.grid(row=0 , column=1  , padx=5 , pady= 5 )
        self.Textodia1.config(justify = "center")
        #MES 1
        self.labelmes1 = Label(self, text = "Mes: ").grid(row=0 , column=2 , sticky='e' , padx=5 , pady= 5)
        self.Meses= ['Enero' , 'Febrero' , 'Marzo' , 'Abril' , 'Mayo' , 'Junio' , 'Julio' , 'Agosto' , 'Septiembre' , 'Octubre' , 'Noviembre' , 'Diciembre']
        self.Textomes1= StringVar()
        menuMes1=OptionMenu(self , self.Textomes1 , *self.Meses)
        menuMes1.grid( row=0 ,column=3 , padx=2 , pady=2 , sticky='e' )
        menuMes1.config(width=9 ,justify="center")

        #DIA 2
        self.labeldia2 = Label(self, text = "Dia final:", width= 10).grid(row=1 , column=0 , sticky='e' , padx=5 , pady= 5)
        self.Textodia2= Entry(self , width=10)
        self.Textodia2.grid(row=1 , column=1 , padx=5 , pady= 5)
        self.Textodia2.config(justify = "center")
        #MES 2
        self.labelmes2 = Label(self, text = "Mes: ").grid(row=1 , column=2 , sticky='e' , padx=5 , pady= 5)
        self.Textomes2= StringVar()
        menuMes2=OptionMenu(self , self.Textomes2 , *self.Meses)
        menuMes2.grid( row=1 ,column=3 , padx=2 , pady=2 , sticky='e' )
        menuMes2.config(width=9 ,justify="center")
        
        #AÑO EN EJECUCION
        self.labelAno = Label(self, text = "Año: ").grid(row=2 , column=2 , sticky='e' , padx=5 , pady= 5)
        self.Ano= ['2019' , '2020' , '2021' , '2022']
        self.TextoAno= StringVar()
        menuAno=OptionMenu(self , self.TextoAno , *self.Ano)
        menuAno.grid( row=2 ,column=3 , padx=2 , pady=2 , sticky='e' )
        menuAno.config(width=9 ,justify="center")
    
        #BOTONES
        self.buttonExample = Button(self, text = "Generar Reporte" , command= lambda: [ self.generar_semanal(controller.getFichero(), controller.getTestigo() , (self.Textodia1.get() , self.Textodia2.get()) , (self.Textomes1.get() , self.Textomes2.get() ) , self.TextoAno.get()) , self.TextoCorreo()] ).grid(row=3 , column=2  , padx=5 , pady= 5 , columnspan=2 , sticky="w")
        self.buttonAtras = Button(self, text = "Atras" , command= lambda: controller.show_frame(Frame_Inicio)).grid(row=4 , column=0 , sticky='w' , padx=10 , pady= 5)
        
    def generar_semanal(self, filename,testigo , rangoDias , rangoMes , ano):
        self.pasar=0
        error= False
        if rangoDias[0] =="" or rangoDias[1]=="" or rangoMes[0]=="" or rangoMes[1]=="" or int(rangoDias[0])<=0 or int(rangoDias[1])>31 or ano=='' : messagebox.showwarning("Aviso" , "Error en ingreso de datos en los campos designados")
        
        else:
            
            if rangoMes[0]== rangoMes[1]:
                if int(rangoDias[0]) > int(rangoDias[1]):
                    messagebox.showerror("Error", "El dia final es mayor que el dia de inicio")
                    self.pasar=0
                    error=True

            if testigo == 1 and  error == False:
                self.pasar=1
                #Testigo = 1 Indica que corresponde a MOTOROLA
                df_filt= func.limpieza_datos(filename , ano)
                df_filt = func.Filtro_Mot(df_filt)
                df_filt= func.agregar_fecha(df_filt , rangoMes[0] , rangoMes[1] , rangoDias[0] , rangoDias[1])
                #se genera archivo temporal
                df_filt.to_excel("ArchivoTemp.xlsx" ,header= True , index = False, sheet_name="SELL-THROUGH" )
                
                #se cambian las descripciones
                df_filt2= func.cambiar_descripciones()
                df_filt2= pd.DataFrame(df_filt2)
                
                #Redaccion del Correo
                radios, t470co , xt185 , repetidoras , cantTotal =func.ContarCantidadesMotorola(df_filt2)
                self.texto= func.textoCorreoMot(rangoDias[0] , rangoDias[1] , rangoMes[0], rangoMes[1] , ano , radios, t470co , xt185 , repetidoras , cantTotal)
                
                #DIRECTORIO DE DESCARGA
                os.remove('ArchivoTemp.xlsx') #Elimina el archivo el cual no es necesario
                func.guardar_reporte(df_filt2 , False , False , 1)
            
            elif testigo==2 and error == False:
                #Testigo = 2 Indica que corresponde a CAMBIUM
                self.pasar=1
                df_filt= func.limpieza_datos(filename , ano)
                df_filt= func.Filtro_CAMBIUM(df_filt)
                df_filt= func.agregar_fecha(df_filt , rangoMes[0] , rangoMes[1] , rangoDias[0] , rangoDias[1])
                #Redaccion del Correo
                self.texto=func.textoCorreoCambium(rangoDias[0] , rangoDias[1] , rangoMes[0] , rangoMes[1] , ano,  len(df_filt.index))

                #DIRECTORIO DE DESCARGA
                func.guardar_reporte(df_filt , False , False , 1)
            elif testigo==4 and error== False:
                pass

    def  TextoCorreo(self):
        if self.pasar==1:
            newWindows = Toplevel(self)
            #TEXTO A COPIAR
            ComentariosLabel= Label(newWindows , text="REDACCION DEL CORREO").grid(row=0 , column=0 , padx=5 , pady=5 )
            textoComentario= Text(newWindows , width= 50 , height= 20) #Diseña cuadro de texto
            textoComentario.grid( row=1, column=0, padx=5 , pady= 5)
            textoComentario.insert(1.0 , self.texto)
            #SCROLLBAR
            scroll= Scrollbar(newWindows, command=textoComentario.yview) #Se crea scroll
            scroll.grid(row=1, column=1, padx=5 , pady= 5 ,  sticky="nsew" ) #Posicionar scroll
            textoComentario.config(yscrollcommand= scroll.set) #Para que el scroll se mueva con el desplazamiento hacia abajo

class Informes_Acumulados(Frame):
    def __init__(self, container , controller , *args, **kwargs):
        super().__init__( container , *args, **kwargs)
    
        #MES EN EJECUCION
        self.labelmes = Label(self, text = "Mes: ").grid(row=3 , column=5 , sticky='e' , padx=5 , pady= 5)
        self.Meses= ['Enero' , 'Febrero' , 'Marzo' , 'Abril' , 'Mayo' , 'Junio' , 'Julio' , 'Agosto' , 'Septiembre' , 'Octubre' , 'Noviembre' , 'Diciembre']
        self.Textomes= StringVar()
        menuMes=OptionMenu(self , self.Textomes , *self.Meses)
        menuMes.grid( row=3 ,column=6 , padx=2 , pady=2 , sticky='e' )
        menuMes.config(width=9 ,justify="center")

        #AÑO EN EJECUCION
        self.labelAno = Label(self, text = "Año: ").grid(row=4 , column=5 , sticky='e' , padx=5 , pady= 5)
        self.Ano= ['2019' , '2020' , '2021' , '2022']
        self.TextoAno= StringVar()
        menuAno=OptionMenu(self , self.TextoAno , *self.Ano)
        menuAno.grid( row=4 ,column=6 , padx=2 , pady=2 , sticky='e' )
        menuAno.config(width=9 ,justify="center")

        #BOTONES
        self.buttonExample = Button(self, text = "Generar Reporte" , command= lambda: [ self.generar_acumulado(controller.getFichero(), controller.getTestigo() , self.Textomes.get() , self.TextoAno.get() ) ,self.TextoCorreo() ] ).grid(row=9 , column=6  , padx=5 , pady= 5 , columnspan=2 , sticky="w")
        self.buttonAtras = Button(self, text = "Atras" , command= lambda: controller.show_frame(Frame_Inicio)).grid(row=10, column=0 , sticky='w' , padx=10 , pady= 5)
        
    def generar_acumulado(self, filename,testigo, mes , ano):
        self.pasar=0

        if mes == "" or ano=="" : messagebox.showwarning("Aviso", "No se ha seleccionado ningun mes o año")

        else:
            if testigo==1:
                self.pasar=1
                #Testigo=1 Indica que corresponde a Motorola
                df_filt= func.limpieza_datos(filename , ano)
                df_filt = func.Filtro_Mot(df_filt)
                df_filt= func.agregar_fecha(df_filt , mes , "acumulado" ,0,0)
                #se genera archivo temporal
                df_filt.to_excel("ArchivoTemp.xlsx" ,header= True , index = False, sheet_name="SELL-THROUGH" )
                
                #se cambian las descripciones
                df_filt2= func.cambiar_descripciones()
                df_filt2= pd.DataFrame(df_filt2)

                #CALCULAR DIA FINAL
                dia_final= func.calcular_dia_final(df_filt2)
                

                #Redaccion del Correo
                radios, t470co , xt185 , repetidoras , cantTotal =func.ContarCantidadesMotorola(df_filt2)
                self.texto= func.textoCorreoMot("01" , str(dia_final) , mes, mes , ano , radios, t470co , xt185 , repetidoras , cantTotal)
                
                #EXPORTAR REPORTE
                os.remove('ArchivoTemp.xlsx') #Elimina el archivo el cual no es necesario
                func.guardar_reporte(df_filt2, False , False, 1)

            elif testigo==2:
                #Testigo = 2 Indica que corresponde a CAMBIUM
                self.pasar=1
                df_filt= func.limpieza_datos(filename , ano)
                df_filt= func.Filtro_CAMBIUM(df_filt)
                df_filt= func.agregar_fecha(df_filt , mes , "acumulado" ,0,0)#Redaccion del Correo
                #CALCULAR DIA FINAL
                dia_final= func.calcular_dia_final(df_filt)
                self.texto=func.textoCorreoCambium("01" , str(dia_final) , mes , mes , ano,  len(df_filt.index))

                #DIRECTORIO DE DESCARGA
                func.guardar_reporte(df_filt , False ,False, 1)

    def  TextoCorreo(self):
        if self.pasar==1:
            newWindows = Toplevel(self)
            #TEXTO A COPIAR
            ComentariosLabel= Label(newWindows , text="REDACCION DEL CORREO").grid(row=0 , column=0 , padx=5 , pady=5 )
            textoComentario= Text(newWindows , width= 50 , height= 20) #Diseña cuadro de texto
            textoComentario.grid( row=1, column=0, padx=5 , pady= 5)
            textoComentario.insert(1.0 , self.texto)
            #SCROLLBAR
            scroll= Scrollbar(newWindows, command=textoComentario.yview) #Se crea scroll
            scroll.grid(row=1, column=1, padx=5 , pady= 5 ,  sticky="nsew" ) #Posicionar scroll
            textoComentario.config(yscrollcommand= scroll.set) #Para que el scroll se mueva con el desplazamiento hacia abajo
            
                    
class Informe_Facturacion(Frame):
    def __init__(self, container , controller , *args, **kwargs):
        super().__init__( container , *args, **kwargs)
    
        #MES EN EJECUCION
        self.labelmes = Label(self, text = "Mes: ").grid(row=3 , column=5 , sticky='e' , padx=5 , pady= 5)
        self.Meses= ['Enero' , 'Febrero' , 'Marzo' , 'Abril' , 'Mayo' , 'Junio' , 'Julio' , 'Agosto' , 'Septiembre' , 'Octubre' , 'Noviembre' , 'Diciembre']
        self.Textomes= StringVar()
        menuMes=OptionMenu(self , self.Textomes , *self.Meses)
        menuMes.grid( row=3 ,column=6 , padx=2 , pady=2 , sticky='e' )
        menuMes.config(width=9 ,justify="center")

        #AÑO EN EJECUCION
        self.labelAno = Label(self, text = "Año: ").grid(row=4 , column=5 , sticky='e' , padx=5 , pady= 5)
        self.Ano= ['2019' , '2020' , '2021' , '2022']
        self.TextoAno= StringVar()
        menuAno=OptionMenu(self , self.TextoAno , *self.Ano)
        menuAno.grid( row=4 ,column=6 , padx=2 , pady=2 , sticky='e' )
        menuAno.config(width=9 ,justify="center")

        #BOTONES
        self.buttonExample = Button(self, text = "Generar Reporte" , command= lambda: self.GenerarFacturacion(controller.getFichero(), controller.getTestigo() , self.Textomes.get() , self.TextoAno.get() )).grid(row=9 , column=6  , padx=5 , pady= 5 , columnspan=2 , sticky="w")
        self.buttonAtras = Button(self, text = "Atras" , command= lambda: controller.show_frame(Frame_Inicio)).grid(row=10, column=0 , sticky='w' , padx=10 , pady= 5)
        
    def GenerarFacturacion(self, filename,testigo, mes , ano):
        if  ano=='' or mes=='' : messagebox.showwarning("Aviso" , "No se ha ingresado ningun año")

        else:
            if testigo==1:
                #El testigo se deja para proximas funcionalidades y diversificaciones de la idea
                df=pd.read_excel(filename , header=1)
                df= df[df['Factura (Año natural)']== int(ano)]
                df_metas=fact.cargar_metas_y_comerciales_motorola()
                number_Q= fact.seleccion_Q_actual(mes , Qs)
                df_exportar  , fin_titulo = fact.agregar_filtros(df, df_metas , number_Q , mes , Qs[int(number_Q)])
                df_exportar= fact.agregar_Metas(df_exportar, df_metas , number_Q , fin_titulo)
                df_exportar= fact.agregar_Porcentajes(df_exportar,number_Q , fin_titulo)
                df2= fact.Ventas_Motorola_Vertex(df , df_metas)

                #self.texto = fact.textoCorreoFact()
                #EXPORTAR REPORTE
                func.guardar_reporte(df_exportar , df2, False ,2)

class Informe_Q_cants(Frame):
    def __init__(self, container , controller , *args, **kwargs):
        super().__init__( container , *args, **kwargs)
    
        #AÑO EN EJECUCION
        self.labelAno = Label(self, text = "Año: ").grid(row=3 , column=5 , sticky='e' , padx=5 , pady= 5)
        self.Ano= ['2019' , '2020' , '2021' , '2022']
        self.TextoAno= StringVar()
        menuAno=OptionMenu(self , self.TextoAno , *self.Ano)
        menuAno.grid( row=3 ,column=6 , padx=2 , pady=2 , sticky='e' )
        menuAno.config(width=9 ,justify="center")

        #BOTONES
        self.buttonExample = Button(self, text = "Generar Reporte" , command= lambda: self.GenerarQ(controller.getFichero(), controller.getTestigo() , self.TextoAno.get()) ).grid(row=9 , column=6  , padx=5 , pady= 5 , columnspan=2 , sticky="w")
        self.buttonAtras = Button(self, text = "Atras" , command= lambda: controller.show_frame(Frame_Inicio)).grid(row=10, column=0 , sticky='w' , padx=10 , pady= 5)
    
    def GenerarQ(self, filename,testigo , ano):
        if  ano=='' : messagebox.showwarning("Aviso" , "No se ha ingresado ningun año")

        else:
            if testigo==1:
                df_filt= func.limpieza_datos(filename , ano)
                df_filt = func.Filtro_Mot(df_filt)
                #se genera archivo temporal
                df_filt.to_excel("ArchivoTemp.xlsx" ,header= True , index = False, sheet_name="SELL-THROUGH" )
                #se cambian las descripciones
                df_filt2= func.cambiar_descripciones()
                df_filt2= pd.DataFrame(df_filt2)
                otras_Ref= pd.DataFrame()
                tabla_Q , otras_Ref = func.calcular_cantidades_Q(df_filt2)
                tabla_mes= func.calcular_cantidades_mes_a_mes(df_filt2)

                #EXPORTAR REPORTE
                os.remove('ArchivoTemp.xlsx') #Elimina el archivo el cual no es necesario
                func.guardar_reporte(tabla_Q , tabla_mes, otras_Ref ,3)


class Informe_Mes_Mes_Motorola(Frame):
    def __init__(self, container , controller , *args, **kwargs):
        super().__init__( container , *args, **kwargs)
        
        #MES 1
        self.labelmes1 = Label(self, text = "Mes: ").grid(row=0 , column=0 , sticky='e' , padx=5 , pady= 5)
        self.Meses= ['Enero' , 'Febrero' , 'Marzo' , 'Abril' , 'Mayo' , 'Junio' , 'Julio' , 'Agosto' , 'Septiembre' , 'Octubre' , 'Noviembre' , 'Diciembre']
        self.Textomes1= StringVar()
        menuMes1=OptionMenu(self , self.Textomes1 , *self.Meses)
        menuMes1.grid( row=0 ,column=1 , padx=2 , pady=2 , sticky='e' )
        menuMes1.config(width=9 ,justify="center")

        #MES 2
        self.labelmes2 = Label(self, text = "Mes: ").grid(row=1 , column=0 , sticky='e' , padx=5 , pady= 5)
        self.Textomes2= StringVar()
        menuMes2=OptionMenu(self , self.Textomes2 , *self.Meses)
        menuMes2.grid( row=1 ,column=1 , padx=2 , pady=2 , sticky='e' )
        menuMes2.config(width=9 ,justify="center")
        
        self.Anos= ['2019' , '2020' , '2021' , '2022']
        #AÑO
        self.labelAno = Label(self, text = "Año: ").grid(row=0 , column=2 , sticky='e' , padx=5 , pady= 5)
        self.TextoAno= StringVar()
        menuAno=OptionMenu(self , self.TextoAno , *self.Anos)
        menuAno.grid( row=0 ,column=3 , padx=2 , pady=2 , sticky='e' )
        menuAno.config(width=9 ,justify="center")

        #AÑO
        self.labelAno2 = Label(self, text = "Año: ").grid(row=1 , column=2 , sticky='e' , padx=5 , pady= 5)
        self.TextoAno2= StringVar()
        menuAno2=OptionMenu(self , self.TextoAno2 , *self.Anos)
        menuAno2.grid( row=1 ,column=3 , padx=2 , pady=2 , sticky='e' )
        menuAno2.config(width=9 ,justify="center")

        #BOTONES
        self.buttonExample = Button(self, text = "Generar Reporte" , command= lambda: self.generar_reporte(controller.getFichero(), controller.getTestigo() , (self.Textomes1.get() , self.Textomes2.get() ) ,[ self.TextoAno.get() , self.TextoAno2.get() ] ) ).grid(row=3 , column=3  , padx=5 , pady= 5 , columnspan=2 , sticky="w")
        self.buttonAtras = Button(self, text = "Atras" , command= lambda: controller.show_frame(Frame_Inicio)).grid(row=4 , column=0 , sticky='w' , padx=10 , pady= 5)
        
    def generar_reporte(self, filename,testigo , rangoMes , anios):
        self.pasar=0
        if  rangoMes[0]=="" or rangoMes[1]=="" or anios[0]=="" or anios[1]=="" : messagebox.showwarning("Aviso" , "Error en ingreso de datos en los campos designados")

        else:
            if testigo == 1:
                self.pasar=1
                #Testigo = 1 Indica que corresponde a MOTOROLA
                df_filt= func.limpieza_datos(filename , anios)
                df_filt = func.Filtro_Mot(df_filt)

                #se genera archivo temporal
                df_filt.to_excel("ArchivoTemp.xlsx" ,header= True , index = False, sheet_name="SELL-THROUGH" )

                #CREACION DE LA TABLA
                tabla = func.tabla_ref_radios(df_filt)

                df_filt= func.FiltroMesesConAños(df_filt , anios , rangoMes)
                
                tablaExport= func.tabla_minimos(tabla , anios , rangoMes)
                tablaExport= func.hallar_media(tablaExport)


                #DIRECTORIO DE DESCARGA
                os.remove('ArchivoTemp.xlsx') #Elimina el archivo el cual no es necesario
                func.guardar_reporte(tablaExport , False , False , 1)


class Informe_SIMS(Frame):
    def __init__(self, container , controller , *args, **kwargs):
        super().__init__( container , *args, **kwargs)
    
        #MES EN EJECUCION
        self.labelmes = Label(self, text = "Mes: ").grid(row=3 , column=5 , sticky='e' , padx=5 , pady= 5)
        self.Meses= ['Enero' , 'Febrero' , 'Marzo' , 'Abril' , 'Mayo' , 'Junio' , 'Julio' , 'Agosto' , 'Septiembre' , 'Octubre' , 'Noviembre' , 'Diciembre']
        self.Textomes= StringVar()
        menuMes=OptionMenu(self , self.Textomes , *self.Meses)
        menuMes.grid( row=3 ,column=6 , padx=2 , pady=2 , sticky='e' )
        menuMes.config(width=9 ,justify="center")

        #AÑO EN EJECUCION
        self.labelAno = Label(self, text = "Año: ").grid(row=4 , column=5 , sticky='e' , padx=5 , pady= 5)
        self.Ano= ['2019' , '2020' , '2021' , '2022']
        self.TextoAno= StringVar()
        menuAno=OptionMenu(self , self.TextoAno , *self.Ano)
        menuAno.grid( row=4 ,column=6 , padx=2 , pady=2 , sticky='e' )
        menuAno.config(width=9 ,justify="center")

        #BOTONES
        self.buttonExample = Button(self, text = "Generar Reporte" , command= lambda: [ self.GenerarSims(controller.getFichero(), controller.getTestigo() , self.Textomes.get() , self.TextoAno.get() ) , self.TextoCorreo() ] ).grid(row=9 , column=6  , padx=5 , pady= 5 , columnspan=2 , sticky="w")
        self.buttonAtras = Button(self, text = "Atras" , command= lambda: controller.show_frame(Frame_Inicio)).grid(row=10, column=0 , sticky='w' , padx=10 , pady= 5)

    def GenerarSims(self , filename, testigo , mes , ano):
        self.pasar=0

        if mes == "": messagebox.showwarning("Aviso", "No se ha seleccionado ningun mes")

        else:
            if testigo==1:
                self.pasar=1
                #Testigo=1 Indica que corresponde a Motorola
                df_filt= func.limpieza_datos(filename , ano)
                df_filt= func.agregar_fecha(df_filt , mes , "acumulado" ,0,0)
                
                #SIMS RADIOS
                df_filt_Radios = func.Filtro_Mot(df_filt)
                #Se usa esto para contar los radios y generar el texto del correo
                df_filt_Radios.to_excel("ArchivoTemp.xlsx" ,sheet_name="SELL-THROUGH")
                df_filt_Radios= df_filt_Radios.loc[:,['Factura (Cliente)','Producto' , 'Producto (Texto)' , 'Cantidad de factura']]
                df_filt_Radios= df_filt_Radios.sort_values('Cantidad de factura' , ascending=False)
                
                #SIMS BATERIAS
                df_filt_Baterias= func.Filtro_Baterias(df_filt)
                df_filt_Baterias= df_filt_Baterias.loc[:,['Factura (Cliente)','Producto' , 'Producto (Texto)' , 'Cantidad de factura']]
                df_filt_Baterias= df_filt_Baterias.sort_values('Cantidad de factura' , ascending=False)
                
                #SIMS A&E
                df=pd.read_excel(filename , sheet_name="Hoja1" , header=1)
                df= df[df['Factura (Año natural)']== int(ano)]
                df= func.agregar_fecha(df , mes , "acumulado" ,0,0)
                df_filt_AyE = func.Filtro_AyE(df)
                #Solo selecciona las columnas que necesito
                df_filt_AyE= df_filt_AyE.loc[:,['Factura (Cliente)','Producto' , 'Producto (Texto)' , 'Cantidad de factura','Valor neto facturado']]
                df_filt_AyE= df_filt_AyE.sort_values('Cantidad de factura' , ascending=False)

                #Totales sumarizados
                Total_Cant_Baterias= df_filt_Baterias['Cantidad de factura'].sum()
                Total_Fact_AyE= df_filt_AyE['Valor neto facturado'].sum()
                Total_Cant_AyE= df_filt_AyE['Cantidad de factura'].sum()
                 
                #Adicion de los totales a la tabla
                totalGeneral ={ 'Factura (Cliente)':'' , 'Producto':'' , 'Producto (Texto)':'TOTAL' ,'Cantidad de factura': Total_Cant_AyE , 'Valor neto facturado': Total_Fact_AyE }
                df_filt_AyE = df_filt_AyE.append(totalGeneral, ignore_index=True)               

                #se cambian las descripciones
                df_filt2= func.cambiar_descripciones()
                df_filt2= pd.DataFrame(df_filt2)
                #Redaccion del Correo
                radios, t470co , xt185 , repetidoras , cantTotal =func.ContarCantidadesMotorola(df_filt2)
                TotalSIMS= cantTotal + int(Total_Cant_Baterias)
                self.texto= func.textoCorreoSIMS(mes , ano , radios, t470co , xt185 , repetidoras , int(Total_Cant_Baterias) , TotalSIMS , round(Total_Fact_AyE, 2) )
                
                #EXPORTAR REPORTE
                os.remove('ArchivoTemp.xlsx') #Elimina el archivo el cual no es necesario
                func.guardar_reporte(df_filt_Radios , df_filt_Baterias , df_filt_AyE , 3)
    
    def  TextoCorreo(self):
        if self.pasar==1:
            newWindows = Toplevel(self)
            #TEXTO A COPIAR
            ComentariosLabel= Label(newWindows , text="REDACCION DEL CORREO").grid(row=0 , column=0 , padx=5 , pady=5 )
            textoComentario= Text(newWindows , width= 50 , height= 20) #Diseña cuadro de texto
            textoComentario.grid( row=1, column=0, padx=5 , pady= 5)
            textoComentario.insert(1.0 , self.texto)
            #SCROLLBAR
            scroll= Scrollbar(newWindows, command=textoComentario.yview) #Se crea scroll
            scroll.grid(row=1, column=1, padx=5 , pady= 5 ,  sticky="nsew" ) #Posicionar scroll
            textoComentario.config(yscrollcommand= scroll.set) #Para que el scroll se mueva con el desplazamiento hacia abajo

class INFORMES_AIDC(Frame):
    def __init__(self, container , controller , *args, **kwargs):
        super().__init__( container , *args, **kwargs)
        #DIA 1
        self.labeldia1 = Label(self, text = "Dia inicio:" , width= 10).grid(row=0 , column=0 , sticky='e' , padx=5 , pady= 5)
        self.Textodia1= Entry(self, width= 10)
        self.Textodia1.grid(row=0 , column=1  , padx=5 , pady= 5 )
        self.Textodia1.config(justify = "center")
        #MES 1
        self.labelmes1 = Label(self, text = "Mes: ").grid(row=0 , column=2 , sticky='e' , padx=5 , pady= 5)
        self.Meses= ['Enero' , 'Febrero' , 'Marzo' , 'Abril' , 'Mayo' , 'Junio' , 'Julio' , 'Agosto' , 'Septiembre' , 'Octubre' , 'Noviembre' , 'Diciembre']
        self.Textomes1= StringVar()
        menuMes1=OptionMenu(self , self.Textomes1 , *self.Meses)
        menuMes1.grid( row=0 ,column=3 , padx=2 , pady=2 , sticky='e' )
        menuMes1.config(width=9 ,justify="center")

        #DIA 2
        self.labeldia2 = Label(self, text = "Dia final:", width= 10).grid(row=1 , column=0 , sticky='e' , padx=5 , pady= 5)
        self.Textodia2= Entry(self , width=10)
        self.Textodia2.grid(row=1 , column=1 , padx=5 , pady= 5)
        self.Textodia2.config(justify = "center")
        #MES 2
        self.labelmes2 = Label(self, text = "Mes: ").grid(row=1 , column=2 , sticky='e' , padx=5 , pady= 5)
        self.Textomes2= StringVar()
        menuMes2=OptionMenu(self , self.Textomes2 , *self.Meses)
        menuMes2.grid( row=1 ,column=3 , padx=2 , pady=2 , sticky='e' )
        menuMes2.config(width=9 ,justify="center")
        
        #AÑO
        self.labelAno = Label(self, text = "Año: ").grid(row=2 , column=2 , sticky='e' , padx=5 , pady= 5)
        self.Ano= ['2019' , '2020' , '2021' , '2022']
        self.TextoAno= StringVar()
        menuAno=OptionMenu(self , self.TextoAno , *self.Ano)
        menuAno.grid( row=2 ,column=3 , padx=2 , pady=2 , sticky='e' )
        menuAno.config(width=9 ,justify="center")

        #MARCA
        self.labelMarca = Label(self, text = "Marca: ").grid(row=3 , column=0 , sticky='e' , padx=5 , pady= 5)
        self.TextoMarca= StringVar()
        menuMarca=OptionMenu(self , self.TextoMarca , *categoriasAIDC)
        menuMarca.grid( row=3 ,column=1 , padx=2 , pady=2 , sticky='e' )
        menuMarca.config(width=9 ,justify="center")

        #BOTONES
        self.buttonExample = Button(self, text = "Generar Reporte" , command= lambda: self.generar_reporte(controller.getFichero(), controller.getTestigo() , controller.getSeleccion() , (self.Textodia1.get() , self.Textodia2.get()), (self.Textomes1.get() , self.Textomes2.get() ) ,self.TextoAno.get() , self.TextoMarca.get()) ).grid(row=3 , column=3  , padx=5 , pady= 5 , columnspan=2 , sticky="w")
        self.buttonAtras = Button(self, text = "Atras" , command= lambda: controller.show_frame(Frame_Inicio)).grid(row=4 , column=0 , sticky='w' , padx=10 , pady= 5)
    
    def generar_reporte(self, filename,testigo , seleccion,rangoDias , rangoMes , anio , marca):
        self.pasar=0
        if  rangoMes[0]=="" or rangoMes[1]=="" or anio=="" : messagebox.showwarning("Aviso" , "Error en ingreso de datos en los campos designados")

        else:
            df= func.Anio(filename, anio)
            df= func.agregar_fecha(df , rangoMes[0] , rangoMes[1] ,rangoDias[0],rangoDias[1])
            df= func.filtro_AIDC(df,marca ,categoriasAIDC)
            if testigo == 4:
                if seleccion=='Cants':
                    df_cants= func.Tabla_cants_AIDC(df)
                    func.guardar_reporte(df_cants , df , None , 2)
                elif seleccion=='Fact':
                    df_export , df_export2 , df_export3= fact.tabla_AIDC(df)
                    func.guardar_reporte(df_export , df_export2 , df_export3 , 3)
                #edfdnvnefvdscvfredscfveds
                #sdvdfvfeds

root = APP()
root.mainloop()
