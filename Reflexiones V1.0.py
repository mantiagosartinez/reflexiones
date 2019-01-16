import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import xlwt
import time
import datetime

var = tk.Tk()
screen_width = var.winfo_screenwidth()
screen_height = var.winfo_screenheight()

class Ecograma(tk.Tk):
    def __init__(self,master):      
        master.title('Calculo de reflexiones')
        master.resizable(False, False)
        
        #los 2 principales Frames:
        self.medidas = tk.Frame(master, 
                                width=master.winfo_screenwidth(),
                                height = master.winfo_screenheight())
        self.medidas.grid()
        self.graficos = tk.Frame(master, 
                                width=master.winfo_screenwidth(),
                                height = master.winfo_screenheight())
        self.graficos.grid()
        
        #Configuracion del menu
        self.menu = tk.Menu(master)
        self.menu.add_command(label='Ayuda', command=self.ayuda)
        master.config(menu=self.menu)
        
        #Dentro de medidas:
        PX=20 #Parametros para ajustar el padding en todos
        PY=10
        tk.Label(self.medidas, text="X").grid(row=1, column=0, padx=PX)
        tk.Label(self.medidas, text="Y").grid(row=2, column=0, padx=PY)
        
        
        tk.Label(self.medidas,
                 text="Medidas de la sala").grid(row=0,
                                          column=1, padx = PX, pady=PY)
        
        self.SalaX = tk.Entry(self.medidas)
        self.SalaX.grid(row=1, column=1, pady=PY)
        
        self.SalaY = tk.Entry(self.medidas)
        self.SalaY.grid(row=2, column=1, pady=PY)
        
        tk.Label(self.medidas,
                 text="Posicion de la fuente").grid(row=0,
                                             column=2, padx=PX )
        
        self.FuenteX = tk.Entry(self.medidas)
        self.FuenteX.grid(row=1, column=2)
        
        self.FuenteY = tk.Entry(self.medidas)
        self.FuenteY.grid(row=2, column=2)
        
        
        tk.Label(self.medidas,
                 text="Posicion del microfono").grid(row=0,
                                              column=3, padx=PX)
        
        self.MicX = tk.Entry(self.medidas)
        self.MicX.grid(row=1, column=3)
        
        self.MicY = tk.Entry(self.medidas)
        self.MicY.grid(row=2, column=3)
        
        
        #Posicionamiento de los dos botones
        tk.Button(self.medidas,
                  text="Mas opciones",
                  command = self.opciones).grid(row=1, column=4)
        
        
        tk.Button(self.medidas,
                  text="Calcular",
                  command = self.calcular).grid(row=2,
                                         column=4, sticky = tk.S)
        
        #Seteo de los valores predeterminados
        self.nPot = 100.0 
        self.cel = 343.0
        self.a1 = 0.0
        self.a2 = 0.0
        self.a3 = 0.0
        self.a4 = 0.0
        
        ##Seteo del primer grafico:
        
        #Predeterminados
        self.times = np.zeros(5)
        self.niveles = np.zeros(5)
        self.prov = ["---","Inferior", "Derecha", "Superior", "Izquierda"]
        
        self.fig = Figure(figsize=(7,7), dpi=50)

        self.sub = self.fig.add_subplot(111)
        self.sub.stem(self.times,self.niveles)
        self.sub.set_title ("Ecograma", fontsize=16)
        self.sub.set_ylabel("Amplitud", fontsize=14)
        self.sub.set_xlabel("Tiempo[s]", fontsize=14)
        self.sub.tick_params(axis="x",rotation=45)
       
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graficos)
        self.canvas.get_tk_widget().grid(row=5,
                                 column=0, columnspan=3, sticky= tk.S)
        self.canvas.draw()
        
        self.Expo = tk.Button(self.graficos,
                              text="Exportar datos", command=self.exportar)
        self.Expo.grid(row=6, column=0)
        
        
        
    def calcular(self):
        """Funcion encargada de realizar el calculo"""
                
        try:
            Lw = self.nPot.get()
            c = self.cel.get()
            a = [0, float(self.a1.get()),
                 float(self.a2.get()),
                 float(self.a3.get()), float(self.a4.get())]
        except:
            Lw = self.nPot
            c = self.cel
            a = [0, float(self.a1),
                 float(self.a2), float(self.a3), float(self.a4)]
        
        try:
            sx = float(self.SalaX.get())
            sy = float(self.SalaY.get())
            fx = float(self.FuenteX.get())
            fy = float(self.FuenteY.get())
            mx = float(self.MicX.get())
            my = float(self.MicY.get())
        except ValueError:
            if (self.SalaX.get()=='' or self.SalaY.get()=='' 
               or self.FuenteX.get()=='' or self.FuenteY.get()=='' 
               or self.MicX.get()=='' or self.MicY.get()==''):
                error(0) #Error de no completar
                return
            else:
                error(1) #Error de ingresar valores no numericos
                return
        if sx<0 or sy<0 or fx<0 or fy<0 or mx<0 or my<0:
            error(1) #Error de ingresar numeros negativos
            return
        if fx>sx or fy>sy:
            error(2) #Error de que la fuente se halle fuera de la sala
            return
        if mx>sx or my>sy:
            error(3) #Error de que el microfono se halle fuera de la sala
            return
        
        #Calculo de la distancia de cada onda, la directa y cada reflexion
        d0 = (mx-fx)**2 + (fy-my)**2 
        d1 = (mx-fx)**2 + (fy+my)**2
        d2 = (2*sx-mx-fx)**2 + (fy-my)**2
        d3 = (mx-fx)**2 + (2*sy-my-fy)**2
        d4 = (mx+fx)**2 + (fy-my)**2   
        
        d = [d1, d2, d3, d4]
         
        W = 10 ** (Lw/10) #conversion de nivel de potencia a potencia
        L = np.zeros(4)
        n = ["Inferior", "Derecha", "Superior", "Izquierda"]
        
        #Estos auxiliares los coloco para garantizar que el sonido directo
        #este en primer lugar, ya que si la fuente se encuentra en una esquina, 
        #el tiempo de arribo sera igual al de una reflexion y el sonido
        #directo podria no encontrarse en el primer lugar
        if d0 != 0:   
            #Tomo como criterio que si la fuente y el microfono se encuentran 
            #en el mismo punto el nivel de presion sonora en ese punto sera
            #igual al nivel de potencia, esta situacion, aunque sea absurda en 
            #la realidad podria usarse para estudiar unicamente las reflexiones
            aux_l = Lw - 10*np.log10(4*np.pi*d0)
        else:
            aux_l = Lw
            
        
        aux_t = d0 / c
        aux_n = "---"
        
        for i in range(0,4):  
               if a[i+1] != 1:
                   L[i] = 10*np.log10(W*(1-a[i+1])) - 10*np.log10(4*np.pi*d[i])

                   
               else:
                   #Para los casos en que se coloque un coef de absorcion==1
                   #la ecuacion de arriba daria un error por logaritmo negativo
                   L[i] = 0 
                   d[i] = 2**100
                   
        t = [x / c for x in d]

        auxSort = np.argsort(t)
        #print(auxSort,n,t)        
        L = [L[auxSort[0]], L[auxSort[1]],
             L[auxSort[2]], L[auxSort[3]]] 
        
        t = [t[auxSort[0]], t[auxSort[1]],
             t[auxSort[2]], t[auxSort[3]]]
        
        n = [n[auxSort[0]], n[auxSort[1]],
             n[auxSort[2]], n[auxSort[3]]]
        
        L = np.append(aux_l,L)
        t = np.append(aux_t,t)
        n = np.append(aux_n,n)
        print(t)
        print(auxSort)
        print(L)
        print(n)
        
        #El proximo for sirve para borrar las reflexiones en donde el 
        #coeficiente de absorcion sea igual a 0
        
        for i in range(0,len(L)-np.count_nonzero(L)): 
            t = np.delete(t,-1)
            L = np.delete(L,-1)
            n = np.delete(n,-1)
                
        #Asignacion de las variables al objeto
        self.times = t
        self.niveles = L 
        self.prov = n
        self.sub.clear()

        self.sub.stem(self.times,self.niveles)
        self.sub.tick_params(axis="x",rotation=45)
        self.sub.set_title ("Ecograma", fontsize=16)
        self.sub.set_ylabel("Amplitud", fontsize=14)
        self.sub.set_xlabel("Tiempo [s]", fontsize=14)
        self.canvas.draw()
        
    def ayuda(self):
        """Funcion encargada de mostrar la ayuda """
        
        self.top1 = tk.Toplevel()
        self.top1.grid()       
        self.top1.resizable(False, False)
        self.top1.title('Ayuda')
        
        self.top1.text= tk.Text(self.top1, font=("Times New Roman", 10))
        self.top1.text.grid()
 
        self.top1.ok = tk.Button(self.top1,
                                 text="Entendido",
                                 pady=10, command=self.cerrar_ayuda)
        self.top1.ok.grid()
        
        texto = """        La idea del trabajo es tener una 
        herramienta intuitiva y sencilla para el calculo de reflexiones de 
        primer orden de una sala rectangular
        
        En la pantalla principal el usuario debe ingresar el tamaño de la sala,
        la posicion de la fuente, y la posicion del microfono.
        El programa ejecutara un error en caso de 
        que falten uno de los datos, se ingrese uno de forma erronea, o si
        la posicion de la fuente o el microfono supera las dimensiones de la 
        sala.
        
        Desde la ventana de mas opciones uno puede variar la velocidad
        del sonido con la que se trabaja (343 m/s predeterminado), variar el
        nivel de potencia de la fuente sonora (94 dB predeterminado) y elegir 
        los coeficientes de absorcion de cada pared, comenzando desde la 
        inferior en el sentido contrario a las agujas del reloj.
        
        Dado que este programa se limita a hacer el analisis en 2 dimensiones
        el programa puede utilizarse tanto como si cada uno de los lados del 
        rectangulo fuera una pared o como si las dos laterales fueran paredes,
        la superior el techo y la inferior el piso
        
        El presente trabajo ha sido desarrollado por Santiago Martinez
        con ayuda de Federico Bosio.
        
        """
        self.top1.text.insert(tk.END, texto) 
        self.top1.text.configure(state = tk.DISABLED)
    def cerrar_ayuda(self):
        self.top1.destroy()
        
    def opciones(self):
        try:
            cel = self.cel.get()
            nPot = self.nPot.get()
            a1 = self.a1.get()
            a2 = self.a2.get()
            a3 = self.a3.get()
            a4 = self.a4.get()
        except:
            cel = self.cel
            nPot = self.nPot
            a1 = self.a1
            a2 = self.a2
            a3 = self.a3
            a4 = self.a4
                
        self.top2 = tk.Toplevel()
        self.top2.title('Mas opciones')
        self.top2.resizable(False, False)
        self.top2.grid()  
        
        #Creacion del texto y del entry referido a la velocidad del sonido
        tk.Label(self.top2, text="Velocidad del sonido (m/s)").grid(row=0,
                column=0, columnspan=2)       
        self.celaux = tk.Entry(self.top2)
        self.celaux.grid(row=0, column=2, columnspan=2, sticky=tk.E)
        predem(self.celaux, str(cel))
        
        #Creacion del texto y del entry referido al nivel de potencia
        tk.Label(self.top2, text="Nivel de potencia (dB)").grid(row=1,
                column=0, columnspan=2)
        self.nPotaux = tk.Entry(self.top2)
        self.nPotaux.grid(row=1, column=2, columnspan=2, sticky=tk.E)
        predem(self.nPotaux,str(nPot)) #Seteo de los valores predeterminados
        
        #Creacion de los labels de los coeficientes de absorcion
        tk.Label(self.top2, text="Coeficientes de absorcion:").grid(row=2,
                column=0, sticky=tk.E)
        tk.Label(self.top2, text="a1 (inferior)").grid(row=3,
                column=0, columnspan=2)
        tk.Label(self.top2, text="a2 (derecha)").grid(row=4,
                column=0, columnspan=2)
        tk.Label(self.top2, text="a3 (superior)").grid(row=5,
                column=0, columnspan=2)
        tk.Label(self.top2, text="a4 (izquierdo)").grid(row=6,
                column=0, columnspan=2)
        
        #Creacion de los entrys de los coeficientes de absorcion
        self.a1aux = tk.Entry(self.top2)
        self.a1aux.grid(row=3, column=2, columnspan=2)
        predem(self.a1aux,str(a1))
        
        self.a2aux = tk.Entry(self.top2)
        self.a2aux.grid(row=4, column=2, columnspan=2)
        predem(self.a2aux,str(a2))
        
        self.a3aux = tk.Entry(self.top2)
        self.a3aux.grid(row=5, column=2, columnspan=2)
        predem(self.a3aux,str(a3))
        
        self.a4aux = tk.Entry(self.top2)
        self.a4aux.grid(row=6, column=2, columnspan=2)
        predem(self.a4aux,str(a4))
        
        self.top2.guardar = tk.Button(self.top2,
                                      text="Guardar cambios",
                                      command=self.guardarCambiosOpciones)
        self.top2.guardar.grid(row=7, column=2)
        
    def guardarCambiosOpciones(self):
        """Guarda los cambios realizados en la ventana "mas opciones" """
       
        try:
            if float(self.a1aux.get())>1 or float(self.a2aux.get())>1 or float(self.a3aux.get())>1 or float(self.a4aux.get())>1:
                #Error por si el coeficiente de absorcion es mayor a 1
                error(4) 
                return
            self.cel = float(self.celaux.get())
            self.nPot = float(self.nPotaux.get())
            self.a1 = float(self.a1aux.get())
            self.a2 = float(self.a2aux.get())
            self.a3 = float(self.a3aux.get())
            self.a4 = float(self.a4aux.get())
            self.top2.destroy()
            
        except:
            error(1)
            
    def exportar(self):
        """Funcion encargada de exportar los datos del grafico en
        una tabla de calculo"""
        try:
            Lw = self.nPot.get()
            c = self.cel.get()
            a = [0, float(self.a1.get()),
                 float(self.a2.get()),
                 float(self.a3.get()), float(self.a4.get())]
        except:
            Lw = self.nPot
            c = self.cel
            a = [0, float(self.a1),
                 float(self.a2), float(self.a3), float(self.a4)]

        #Definicion de los estylos de celda
        paramSal = ("align: wrap yes,vert centre, horiz center;"
                    "pattern: pattern solid, fore-colour light_orange;"
                    "border: left thin,right thin,top thin,bottom thin")
                               
        styleTitulo = ("align: wrap yes,vert centre, horiz center;"
                       "pattern: pattern solid, fore-colour red;"
                       "border: left thin,right thin,top thin,bottom thin")
        
        styleFecha = "align: wrap yes"
        
        paramEnt = ("align: wrap yes,vert centre, horiz center;"
                    "pattern: pattern solid, fore-colour light_blue;"
                    "border: left thin,right thin,top thin,bottom thin")
                               
        styleS = xlwt.easyxf(paramSal)
        styleT = xlwt.easyxf(styleTitulo)
        styleF= xlwt.easyxf(styleFecha)
        styleE= xlwt.easyxf(paramEnt)

        t = self.times
        L = self.niveles
        
        #de
        book = xlwt.Workbook(encoding="utf-8")
        hoja1 = book.add_sheet("Hoja 1")
        hoja1.write_merge(0, 1, 4, 5, "Ecograma", style=styleT)
        
        hoja1.write(2, 1, "Fecha: ", style=styleF)
        hoja1.write(2, 2,
                    datetime.datetime.now().strftime("%d-%m-%y"), style=styleF)
        hoja1.write(3, 1, "Hora: ", style=styleF)
        hoja1.write(3, 2, time.strftime("%H:%M:%S"), style=styleF)
        
        hoja1.write(5, 1, "Parametros de entrada", style=styleE)
        hoja1.write(5, 2, "x[m]", style=styleE)
        hoja1.write(5, 3, "y[m]", style=styleE)
        hoja1.write(6, 1, "Dimensiones de la sala", style=styleE)
        hoja1.write(7, 1, "Posicion de la fuente", style=styleE)
        hoja1.write(8, 1, "Posicion del microfono", style=styleE)
        hoja1.write(6, 2, self.SalaX.get(), style=styleE)
        hoja1.write(6, 3,  self.SalaY.get(), style=styleE)
        hoja1.write(7, 2, self.FuenteX.get(), style=styleE)
        hoja1.write(7, 3,  self.FuenteY.get(), style=styleE)
        hoja1.write(8, 2, self.MicX.get(), style=styleE)
        hoja1.write(8, 3,  self.MicY.get(), style=styleE)
        
        hoja1.write(10, 1, "Velocidad del sonido", style=styleE)
        hoja1.write(10, 2, c, style=styleE)
        
        hoja1.write(11, 1, "Potencia[dB]", style=styleE)
        hoja1.write(11, 2, Lw, style=styleE)
        
        hoja1.write(12, 1, "Coef de absorcion (pared inferior)", style=styleE)
        hoja1.write(12, 2, a[1], style=styleE)
        
        hoja1.write(13, 1, "Coef de absorcion (pared derecha)", style=styleE)
        hoja1.write(13, 2, a[2], style=styleE)
        
        hoja1.write(14, 1, "Coef de absorcion (pared superior)", style=styleE)
        hoja1.write(14, 2, a[3], style=styleE)
        
        hoja1.write(15, 1, "Coef de absorcion (pared izquierda)", style=styleE)
        hoja1.write(15, 2, a[4], style=styleE)
        
                   
        hoja1.write(5, 6, "Parametros de salida", style=styleS)
        hoja1.write(5, 7, "Tiempo [s]", style=styleS)
        hoja1.write(5, 8, "Nivel de presion sonora [dB]", style=styleS)
        hoja1.write(5, 9, "Proveniencia de la reflexion", style=styleS)
        
        hoja1.write(6, 6,"Sonido directo",style=styleS)
        hoja1.write(6, 7, str(round(t[0],5)), style=styleS)
        hoja1.write(6, 8, str(round(L[0],2)), style=styleS)
        hoja1.write(6, 9, self.prov[0], style=styleS)
        
        
        for i in range(1,len(t)):
            hoja1.write(6+i, 6, "Reflexion " + str(i), style=styleS)
            hoja1.write(6+i, 7, str(round(t[i],5)), style=styleS)
            hoja1.write(6+i, 8, str(round(L[i],2)), style=styleS)
            hoja1.write(6+i, 9, str(self.prov[i]), style=styleS)
        
        try:
            path = filedialog.asksaveasfilename()
            book.save(path)
        except:
            return
            
            
        
     
def error(a):
        """Muestra si se comete un error en el programa """
            
        if a==0:
            """Si deja un valor sin completar"""
                
            messagebox.showinfo("ERROR", ("Ha dejado al menos una dimension"
                                        "sin completar, por favor, "
                                        "completela"))
 
        elif a==1:
            """Si se ingresa un valor que no es un numero positivo"""
                
            messagebox.showinfo("ERROR", ("Ha ingresado un valor que no es" 
                                          "un numero positivo, por favor, "
                                          "cambielo"))
        elif a==2:
            """Si la ubicacion de la fuente excede las
                dimensiones de la sala"""
                
            messagebox.showinfo('ERROR', ("Revise la posicion de la fuente, "
                                          "segun lo que has ingresado la "
                                          "fuente está afuera de la sala"))
        elif a==3:
            """Si la ubicacion del microfono excede las
                dimensiones de la sala"""
                
            messagebox.showinfo('ERROR', ("Revise la posicion del microfono"
                                          ", segun lo que has ingresado "
                                          "el microfono está "
                                          "afuera de la sala"))
        elif a==4:
            """Si el coef de absorcion es mayor a 1"""
                
            messagebox.showinfo('ERROR', ( "El coeficiente"
                                         "de absorcion debe ser menor a 1"))
  
    
def predem(a,text):
    """ Muestra los valores predeterminados de los Entry correspondientes 
    en la ventana "Mas opciones"."""
    a.insert(0, text)
    a.config(fg = 'grey')
    def click(event):
            """funcion que e activa cada vez que se clickea en el entry"""
            if a.get() == text:
               a.delete(0, "end") # delete all the text in the entry
               a.insert(0, '') #Insert blank for user input
               a.config(fg = 'black')
    def foco(event):
            if a.get() == '':
               a.insert(0, text)
               a.config(fg = 'grey')
    a.bind('<FocusIn>', click)
    a.bind('<FocusOut>', foco)
    
a = Ecograma(var)
var.mainloop()