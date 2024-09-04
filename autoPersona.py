
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

### seguir en este, actualizar luego subiendo a github

import os
import ctypes

from tkinter import *
from tkinter import ttk

from docx import Document 
from docx.shared import Inches

#from customtkinter import *


path = os.getcwd()
# print(path)

# OJO: todo se está guardando en "c:\Users\jfz\" desde el equino, no en drive

# mejora calidad de imagen de la interfaz a crear
ctypes.windll.shcore.SetProcessDpiAwareness(True)

# creación ventana
# ventana = CTk() custom, para el final, modificar alfinal
ventana = Tk()

ventana.geometry("1200x700")

#set_appearance_mode("dark") # modo oscyro 

ventana.title("autoPersonas") # titulo de la ventana

# ajustar medida de la ventana, no se lee el titulo

#tkinter._test()

# creación de 3 frames (grupos/secciones)
frm_a = ttk.Frame(ventana) # frame a, , backgroud='cyan'
frm_b = ttk.Frame() # frame b
frm_c = ttk.Frame() # frame c

#fred = Button(frm, fg="red", bg="blue")

############ casilla 1

list_1 = ["Decreto", "En fecha", "Regularización", "Modificación"]

label_1 = Label(master = frm_a, text="Tipo de decreto", padx=10)
label_1.grid(row=0, column=0, sticky=N+W)

# Create Dropdown menu 
entradaElegida = StringVar()
entradaElegida.set("Seleccionar")

menu_1 = OptionMenu(frm_a, entradaElegida, *list_1)
menu_1.grid(row=0, column=1, sticky=N+W)

frm_a.grid_rowconfigure(1, minsize=10)
frm_a.place(x = 50, y = 50)

tipo_decreto = entradaElegida.get() #Variable creada: entradaElegida

# imprime en terminal, luego tengo que pasarlo a un variable
def print_answers(): 
    print("Selected Option: {}".format(tipo_decreto)) 
    #luego debo agregar todas las respuestas aqui
    return None

############ casilla 2

label_2 = Label(master = frm_a, text="Solicitud solicitante vía", padx=10)
label_2.grid(row=1, column=0, sticky=N+W)

entrada_SSV = StringVar() # podria ser int binario tambien com IntVarblbl() 
Radiobutton(frm_a, text = "Memo", padx = 1, fg="red", textvariable=entrada_SSV, command=entrada_SSV.get(), value='Memo').grid(row=1, column=1, sticky=N+W)
Radiobutton(frm_a, text = "Correo electrónico", padx = 1, fg="black", textvariable=entrada_SSV, command=entrada_SSV.get(), value='Correo electrónico').grid(row=1, column=2, sticky=N+W)


############ casilla 3

# Dropdown menu options 
list_3 = ["Beneficio", "Plazo", "Renta", "Cometido"]

label_3 = Label(master = frm_a, text = "Modificación", padx=10)
label_3.grid(row=2, column=0, sticky=N+W)

# Create Dropdown menu 
entryModificacion = StringVar()
entryModificacion.set("Seleccionar") 

menu_3 = OptionMenu(frm_a, entryModificacion, *list_3)
menu_3.grid(row=2, column=1, sticky=N+W)

tipo_modificacion = entryModificacion.get() #Variable creada: entryModificacion

# imprime en terminal, luego tengo que pasarlo a un variable
def print_answers(): 
    print("Selected Option: {}".format(tipo_modificacion)) 
    #luego debo agregar todas las respuestas aqui
    return None


# casilla 4

label_4 = Label(master = frm_a, text="Aprobación vía", padx=10)
label_4.grid(row=3, column=0, sticky=N+W)

entrada_AV = StringVar() # podria ser int binario tambien com IntVarblbl() 
Radiobutton(frm_a, text = "Memo", padx = 1, fg="red", textvariable=entrada_AV, command=entrada_AV.get(), value='Memo').grid(row=3, column=1, sticky=N+W)
Radiobutton(frm_a, text = "Correo electrónico", padx = 1, fg="black", textvariable=entrada_AV, command=entrada_AV.get(), value='Correo electrónico').grid(row=3, column=2, sticky=N+W)

### FIX: se marcan SSV y AV ###

'''
# casilla 5

label_5 = Label(master = frm_b, text="Fecha aprobación modificación")  
#labl_1.place(x = 100, y = 130)  
  
entry_5 = Entry(master = frm_b)  
#entry_1.place(x=500,y=130)  

label_5.pack()
entry_5.pack()

# casilla 6

label_6 = Label(master = frm_b, text="Fechainstumento contrat.")  
#labl_1.place(x = 100, y = 130)  
  
entry_6 = Entry(master = frm_b)  
#entry_1.place(x=500,y=130)  

label_6.pack()
entry_6.pack()

# casilla 7

label_7 = Label(master = frm_b, text="Decreto SIAPER aprobación modificación")  
#labl_1.place(x = 100, y = 130)  
  
entry_7 = Entry(master = frm_b)  
#entry_1.place(x=500,y=130)  

label_7.pack()
entry_7.pack()

#################################### en adelante intergrar respuestas de acasillas anteriores ########################

# imprime en terminal, luego tengo que pasarlo a un variable
def print_answers(): 
    print("Selected Option: {}".format(texto_casilla.get())) 
    #luego debo agregar todas las respuestas aqui
    return None

# Submit button 
# Whenever we click the submit button, our submitted 
# option is printed ---Testing purpose 
submit_button = Button(frm_a, text='Submit', command=print_answers) 
submit_button.pack() 

############################################################################################################

# casilla cerrar: quit

# escribinos frame c con boton "quit"
label_c = ttk.Label(master = frm_c, text = "Vistos y considerando")
boton_c = ttk.Button(master = frm_c, text = "Quit", command = ventana.destroy)
label_c.pack()
boton_c.pack()

# escribirmos frames en la ventana de interfaz


frm_a.pack()
frm_b.pack()
frm_c.pack()
'''
# se abre ventana
ventana.mainloop()

# creacion word
document = Document() 
document.add_heading('Vitacura', 0)

# escribir el doc

# revisar la siguiente linea miercoles 4s sept
document.add_paragraph("Lore ipsum", style='Intense Quote')

# guardar el doc

document.save("test_autoPersona.docx") 

# al final debo hace un packing .exe