
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

### seguir en este, actualizar luego subiendo a github

import os
import ctypes

from tkinter import *
from tkinter import ttk

from tkcalendar import DateEntry

from docx import Document 
from docx.shared import Inches

#from customtkinter import *

def get_entradassv_var():
    global SSV_var
    SSV_var = entrada_SSV.get() # Variable creada
'''
def get_mod_var():
    global tipo_mod_var
    tipo_mod_var = entryModificacion.get() # Variable creada
'''
def get_av_var():
    global entrada_AV_var
    entrada_AV_var = entrada_AV.get() # Variable creada

def modulo_modificacion(): # PENDIENTE: HACER QUE SE ABRA PESTAÑA O PONER COLUMNA A LA DERECHA
    
    ############ casilla 2

    label_2 = Label(master = frm_b, text="Solicitud solicitante vía", padx=10)
    label_2.grid(row=1, column=0, sticky=N+W)

    global entrada_SSV
    entrada_SSV = StringVar() # podria ser int binario tambien com IntVarblbl() 
    
    memo_2 = Radiobutton(frm_b, text = "Memo", padx = 1, fg="black", variable=entrada_SSV, command=get_entradassv_var, value='Memo')
    memo_2.grid(row=1, column=1, sticky=N+W)
    mail_2 = Radiobutton(frm_b, text = "Correo electrónico", padx = 1, fg="black", variable=entrada_SSV, command=get_entradassv_var, value='Correo electrónico')
    mail_2.grid(row=1, column=2, sticky=N+W)

    '''
    global SSV_var
    entrada_SSV_var = str(entrada_SSV.get()) #Variable creada: entradaElegida
    '''
    ############ casilla 3

    # Dropdown menu options 
    list_3 = ["Beneficio", "Plazo", "Renta", "Cometido"]

    label_3 = Label(master = frm_b, text = "Tipo de Modificación", padx=10)
    label_3.grid(row=2, column=0, sticky=N+W)

    # Create Dropdown menu 
    global entryModificacion
    entryModificacion = StringVar()
    entryModificacion.set("Seleccionar") 

    menu_3 = ttk.Combobox(frm_b, width = 27, textvariable = entryModificacion, values = list_3) 
    menu_3.grid(row=2, column=1, sticky=N+W)

    global tipo_mod_var
    tipo_mod_var = "{}".format(entryModificacion.get()) # Variable creada
    menu_3.set(tipo_mod_var) #Variable creada


    ### NO ESTA FUNCIONANDO ### AVANZAR EN ESTO LUEGO DE HACER MODULO DE SALUD/EDUCACION

    # menucombobox.delete("0", tk.END) # this will clear the field after button click

    # casilla 4

    label_4 = Label(master = frm_b, text="Aprobación vía", padx=10)
    label_4.grid(row=3, column=0, sticky=N+W)

    global entrada_AV
    entrada_AV = StringVar() # podria ser int binario tambien com IntVarblbl() 

    memo_4 = Radiobutton(frm_b, text = "Memo", padx = 1, fg="black", variable=entrada_AV, command=get_av_var, value='Memo')
    memo_4.grid(row=3, column=1, sticky=N+W)
    mail_4 = Radiobutton(frm_b, text = "Correo electrónico", padx = 1, fg="black", variable=entrada_AV, command=get_av_var, value='Correo electrónico')
    mail_4.grid(row=3, column=2, sticky=N+W)

    ### FIX: se marca ERROR SSV y MOD ###


    # casilla 5 :  Fecha aprobación modificación

    # Add Calendar
    
    # Add Button and Label

    sig2a_button = Button(frm_b, text='Siguiente 2a', command=det_direccion) 
    sig2a_button.grid(row=12, column=2)

def modulo_modificacion_2():
    selected_date = cal_5.get_date()
    selected_date_label.config(text=f"Selected Date: {selected_date}")

    date_5 = Label(frm_a, text='Fecha aprobación modificación')
    date_5.grid(row=4, column=0, sticky=N+W)  

    cal_5 = DateEntry(frm_a, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern="dd-mm-yyyy")
    cal_5.grid(row=4, column=1, sticky=N+W) 

    get_date_button = Button(frm_a, text="Guardar fecha", command = get_selected_date)
    get_date_button.grid(row=4, column=2, sticky=N+W) 

    selected_date_label = Label(frm_a, text="") #variabkle?
    selected_date_label.grid(row=4, column=3, sticky=N+W) 

    # casilla 6

    selected_date = cal_6.get_date()
    selected_date_label_6.config(text=f"Selected Date: {selected_date}")

    date_6 = Label(frm_a, text='Fecha instumento contrat.')
    date_6.grid(row=5, column=0, sticky=N+W)  

    cal_6 = DateEntry(frm_a, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern="dd-mm-yyyy")
    cal_6.grid(row=5, column=1, sticky=N+W) 

    get_date_button = Button(frm_a, text="Guardar fecha", command = get_selected_date)
    get_date_button.grid(row=5, column=2, sticky=N+W) 

    selected_date_label_6 = Label(frm_a, text="")
    selected_date_label_6.grid(row=5, column=3, sticky=N+W) 

    print(selected_date_label_6)

    # la casilla 7 está por fuera

    # casilla 8

    label_8 = Label(master = frm_a, text="Nro Decreto")  
    label_8.grid(row=7, column=0, sticky=N+W)
    
    entry_8 = Entry(master = frm_a)  
    entry_8.grid(row=7, column=1, sticky=N+W)

    nro_decreto = entry_8.get()

    print(nro_decreto)   

    # falta para este modulo la fecha de decreto siaper aprovacion modificacion 
    '''
    sig_button = Button(frm_b, text='Siguiente 2', command=det_direccion) 
    sig_button.grid(row=12, column=2)
    '''

def det_direccion():

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
    label_7.grid(row=2, column=0, sticky=N+W)

    # Create Dropdown menu 
    global direccionSolicitante
    direccionSolicitante = StringVar()
    direccionSolicitante.set("Seleccionar")

    menu_7 = ttk.Combobox(frm_a, width = 50, textvariable = direccionSolicitante, values = list_7)
    menu_7.grid(row=2, column=1, sticky=N+W)

    global direccion_sol_var
    direccion_sol_var = "{}".format(direccionSolicitante.get()) #Variable creada
    menu_7.set(direccion_sol_var) # Cambiamos el nombre de la selda por la seleccion

    sig2b_button = Button(frm_b, text='Siguiente 2b', command=choice()) 
    sig2b_button.grid(row=3, column=2)

    # Esto está tirando "Seleccionar"


    # funcion desplegable condicional
def choice():
    # Checkbutton is checked

    if direccionSolicitante.get() == "Dirección de Salud y Educación, y Demás Incorporados en la Gestión Municipal": 
    
        label_ems = Label(master = frm_a, text="Departamento", padx=10)
        label_ems.grid(row=4, column=0, sticky=N+W)

        # Create Dropdown menu 
        global depto
        depto = StringVar()
        depto.set("Seleccionar")

        menu_ems = OptionMenu(frm_a, depto, *["Salud", "Educación"])
        menu_ems.grid(row=4, column=1, sticky=N+W)

        direccion_sol_var = depto.get() # Variable creada

        sig3_button = Button(frm_b, text='Siguiente 3', command=print("seleccionaste :{}".format(depto.get()))) # ingressar comando
        sig3_button.grid(row=3, column=2)

        # if salud, entonces: COSAM, CESFAM, o Depto Salud
        
    return None

    #else: # ARREGLAR FIX ESTOOOO 

        # ALGO DEBE PASAR PARA QUE AL MOMENTO DE CAMBIAR LA DIRECCION, SE DEJE DE NOSTRAR EL WIDGET DE DEPERTAMENTO SALUD/EDUCACION

# imprime en terminal, luego tengo que pasarlo a un variable
def print_answers(): 

    global tipo_decreto_var
    tipo_decreto_var = "{}".format(entradaElegida.get())

    sig1_button.grid_forget()

    if tipo_decreto_var == "Modificación":
        print("entró")
        modulo_modificacion() # LUEGO DEBO ENCADENWAR ESTA FUNCION 
    else:
        det_direccion()

    '''
    # REVISAR FIX ESTO ERROR EN QYE NO RECONOCE VARIABLE, PERO LA ESCRIBE IGUAL?
    print("Tipo de decreto: ", tipo_decreto_var)
    print("Solicitud solicitante vía: ", "SSV_var") 
    print("Tipo de modificación: ", tipo_mod_var)
    print("Aprobación vía: ", entrada_AV_var)
    #print("Fecha aprobación modificación: {}".format(selected_date_label_6)) 
    '''



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

#tkinter._test()

# creación de 3 frames (grupos/secciones)

frm_a = ttk.Frame(ventana) # frame a, , backgroud='cyan'
frm_b = ttk.Frame(ventana) # frame b
#frm_c = ttk.Frame() # frame c

#fred = Button(frm, fg="red", bg="blue")

#frm_a.grid_columnconfigure(0, weight=1)
#frm_a.grid_columnconfigure(0, weight=2)


frm_a.grid_rowconfigure(1, minsize=10)
frm_a.place(x = 50, y = 50)

frm_b.grid_rowconfigure(2, minsize=10)
frm_b.place(x = 50, y = 100)

### frm_a.grid_rowconfigure(1, minsize=10)


############ casilla 1

list_1 = ["En fecha", "Regularización", "Modificación"]

label_1 = Label(master = frm_a, text="Tipo de decreto", padx=10)
label_1.grid(row=0, column=0, sticky=N+W)

# Create Dropdown menu 
entradaElegida = StringVar()
entradaElegida.set("Seleccionar")

#menu_1 = OptionMenu(frm_a, entradaElegida, *list_1)

tipo_elegido = ttk.Combobox(frm_a, width = 20, textvariable = entradaElegida, values = list_1) 

tipo_elegido.grid(row=0, column=1, sticky=N+W)

#menu_1.grid(row=0, column=1, sticky=N+W)


#global tipo_decreto
entradaElegida_var = "{}".format(entradaElegida.get()) #Variable creada
tipo_elegido.set(entradaElegida_var) # Cambiamos el nombre de la selda por la seleccion

# ARREGLAR todo con GET: get funciona solo una vez



#################################### en adelante intergrar respuestas de acasillas anteriores ########################

# REVISAR VALORES DE VARIABLES, NO ESTAN BIEN, ESTO PARA VIERNES 6

# Submit button 
# Whenever we click the submit button, our submitted 
# option is printed ---Testing purpose 


sig1_button = Button(frm_b, text='Siguiente 1', command=print_answers) 
sig1_button.grid(row=10, column=1)

'''
if entradaElegida.get() == "Modificación":
    print("entró")
    modulo_modificacion()
    
else:
    None
    '''

# se abre ventana
# debe posicoinarse despues de llamar a todos los widgets
ventana.mainloop()


'''
# casilla cerrar: quit

# escribinos frame c con boton "quit"

label_c = ttk.Label(master = frm_c, text = "Vistos y considerando")
boton_c = ttk.Button(master = frm_c, text = "Quit", command = ventana.destroy)
label_c.pack()
boton_c.pack()

'''

###################################
########## CREACIÓN WORD ##########
###################################

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