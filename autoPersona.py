
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

### seguir en este, actualizar luego subiendo a github

import os
import ctypes

from tkinter import *
from tkinter import ttk
from tkcalendar import Calendar, DateEntry

from docx import Document 
from docx.shared import Inches

#from customtkinter import *

# funcion desplegable condicional
def choice(direccion_sol):
    # Checkbutton is checked.
    direccion_sol = direccionSolicitante.get() # Variable creada

    if direccion_sol == "Dirección de Salud y Educación, y Demás Incorporados en la Gestión Municipal": 
    
        label_ems = Label(master = frm_a, text="Departamento", padx=10)
        label_ems.grid(row=6, column=4, sticky=N+W)

        # Create Dropdown menu 
        depto = StringVar()
        depto.set("Seleccionar")

        menu_ems = OptionMenu(frm_a, depto, *["Salud", "Educación"])
        menu_ems.grid(row=6, column=5, sticky=N+W)    

        direccion_sol = depto.get() # Variable creada
    return None

    #else: # ARREGLAR FIX ESTOOOO 

        # ALGO DEBE PASAR PARA QUE AL MOMENTO DE CAMBIAR LA DIRECCION, SE DEJE DE NOSTRAR EL WIDGET DE DEPERTAMENTO SALUD/EDUCACION

# imprime en terminal, luego tengo que pasarlo a un variable
def print_answers(): 
    print("Tipo de decreto: {}".format(tipo_decreto)) 
    print("Solicitud solicitante vía: {}".format("PENDIENTE")) 
    print("Modificación: {}".format("tipo_modificacion")) 
    print("Aprobación vía: {}".format("PENDIENTE")) 
    print("Fecha aprobación modificación: {}".format(selected_date_label)) 
    return None


# path
# path = os.getcwd()
# print(path)

# OJO: todo se está guardando en "c:\Users\jfz\" desde el equino, no en drive

# mejora calidad de imagen de la interfaz a crear
ctypes.windll.shcore.SetProcessDpiAwareness(True)

# creación ventana
# ventana = CTk() custom, para el final, modificar alfinal
ventana = Tk()

ventana.geometry("1600x700")

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

############ casilla 2

label_2 = Label(master = frm_a, text="Solicitud solicitante vía", padx=10)
label_2.grid(row=1, column=0, sticky=N+W)

entrada_SSV = StringVar() # podria ser int binario tambien com IntVarblbl() 
memo_2 = Radiobutton(frm_a, text = "Memo", padx = 1, fg="red", textvariable=entrada_SSV, command=entrada_SSV.get(), value='Memo')
memo_2.grid(row=1, column=1, sticky=N+W)
mail_2 = Radiobutton(frm_a, text = "Correo electrónico", padx = 1, fg="red", textvariable=entrada_SSV, command=entrada_SSV.get(), value='Correo electrónico')
mail_2.grid(row=1, column=2, sticky=N+W)

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

# casilla 4

label_4 = Label(master = frm_a, text="Aprobación vía", padx=10)
label_4.grid(row=3, column=0, sticky=N+W)

entrada_AV = StringVar() # podria ser int binario tambien com IntVarblbl() 
memo_4 = Radiobutton(frm_a, text = "Memo", padx = 1, fg="black", textvariable=entrada_AV, command=entrada_AV.get(), value='Memo')
memo_4.grid(row=3, column=1, sticky=N+W)
mail_4 = Radiobutton(frm_a, text = "Correo electrónico", padx = 1, fg="black", textvariable=entrada_AV, command=entrada_AV.get(), value='Correo electrónico')
mail_4.grid(row=3, column=2, sticky=N+W)

### FIX: se marcan SSV y AV ###


# casilla 5 :  Fecha aprobación modificación

# Add Calendar
 
# Add Button and Label

def get_selected_date():
    selected_date = cal_5.get_date()
    selected_date_label.config(text=f"Selected Date: {selected_date}")

date_5 = Label(frm_a, text='Fecha aprobación modificación')
date_5.grid(row=4, column=0, sticky=N+W)  

cal_5 = DateEntry(frm_a, width=12, background='darkblue',
                    foreground='white', borderwidth=2)
cal_5.grid(row=4, column=1, sticky=N+W) 

get_date_button = Button(frm_a, text="Guardar fecha", command = get_selected_date)
get_date_button.grid(row=4, column=2, sticky=N+W) 

selected_date_label = Label(frm_a, text="") #variabkle?
selected_date_label.grid(row=4, column=3, sticky=N+W) 

# casilla 6

def get_selected_date():
    selected_date = cal_6.get_date()
    selected_date_label_6.config(text=f"Selected Date: {selected_date}")

date_6 = Label(frm_a, text='Fecha instumento contrat.')
date_6.grid(row=5, column=0, sticky=N+W)  

cal_6 = DateEntry(frm_a, width=12, background='darkblue',
                    foreground='white', borderwidth=2)
cal_6.grid(row=5, column=1, sticky=N+W) 

get_date_button = Button(frm_a, text="Guardar fecha", command = get_selected_date)
get_date_button.grid(row=5, column=2, sticky=N+W) 

selected_date_label_6 = Label(frm_a, text="")
selected_date_label_6.grid(row=5, column=3, sticky=N+W) 

print(selected_date_label_6)


# casilla 7: Dirección solicitante EDITING

# PENDIENTE: hacer csv de direccionres y subdirecciones, a esto se le podria asociar los programas
list_7 = ["Dirección de Personas",
        "Secretaría Comunal de Planificación",
        "Dirección de Salud y Educación, y Demás Incorporados en la Gestión Municipal",
        "Dirección de Asesoría Jurídica",
        "Dirección de Sustentabilidad e Innovación",
        "Dirección de Administración y Finanzas",
        "Dirección de Asesoría Urbana",
        "Dirección de Desarrollo Comunitario",
        "Dirección de Informática",
        "Dirección de Obras Municipales",
        "Dirección de Comunicaciones, Asuntos Corporativos y Prensa",
        "Dirección de Medio Ambiente, Aseo y Ornato",
        "Dirección de Infraestructura Comunal",
        "Dirección de Tránsito y Transporte Público",
        "Dirección de Seguridad Pública",
        "Dirección de Control",
        "Secretaría Municipal",
        "Administración Municipal",
        ]

label_7 = Label(master = frm_a, text="Dirección solicitante", padx=10)
label_7.grid(row=6, column=0, sticky=N+W)

# Create Dropdown menu 
direccionSolicitante = StringVar()
direccionSolicitante.set("Seleccionar")

menu_7 = OptionMenu(frm_a, direccionSolicitante, *list_7, command=choice)
menu_7.grid(row=6, column=1, sticky=N+W)

#frm_a.grid_rowconfigure(1, minsize=10)
#frm_a.place(x = 50, y = 50)

# casilla 8

label_8 = Label(master = frm_a, text="Nro Decreto")  
label_8.grid(row=7, column=0, sticky=N+W)
  
entry_8 = Entry(master = frm_a)  
entry_8.grid(row=7, column=1, sticky=N+W)

nro_decreto = entry_8.get()

print(nro_decreto)

#################################### en adelante intergrar respuestas de acasillas anteriores ########################

# REVISAR VALORES DE VARIABLES, NO ESTAN BIEN, ESTO PARA VIERNES 6

# Submit button 
# Whenever we click the submit button, our submitted 
# option is printed ---Testing purpose 
submit_button = Button(frm_a, text='Submit', command=print_answers) 
submit_button.grid(row=9, column=0, sticky=N+W)

############################################################################################################
'''
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

# leemos parafos y escribimos vistos y condiciones

# vistos y considerando, parrafo 1 y 2
vyc_1_2 = open("vyc_1+2_siempre.dat", "r", encoding='utf-8')
content_1_2 = vyc_1_2.readlines() # lista con lineas del texto, finaliza con \n
vyc_1_2.close()

# escribir el docx
for nro_parrafo in content_1_2:
    document.add_paragraph(nro_parrafo)


# vistos y considerando, parrafo 3
vyc_3 = open("vyc_3.dat", "r", encoding='utf-8')
content_3 = vyc_3.readlines() # lista con lineas del texto, finaliza con \n
vyc_3.close()

'''
# escribir el docx, falta elegir el parrafo y traer una variables con el departamento: salud, educacion o municipal
for nro_parrafo in content_3:
    m_e_s, parrafo = nro_parrafo.split("|") # m_e_s es "Municipal", "Educación" o "Salud"
    if m_e_s == "Municipal" or m_e_s == "Educación":
        document.add_paragraph(parrafo)
    elif m_e_s == "Salud":
        document.add_paragraph(parrafo)
'''

# guardar el doc

document.save("test_autoPersona.docx") 

# al final debo hace un packing .exe