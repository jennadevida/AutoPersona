
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

### seguir en este, actualizar luego subiendo a github

import os
import ctypes

from tkinter import *
from tkinter import ttk

from docx import Document 
from docx.shared import Inches

path = os.getcwd()
print(path)

# OJO: todo se está guardando en "c:\Users\jfz\" desde el equino, no en drive

# mejora calidad de imagen de la interfaz a crear
ctypes.windll.shcore.SetProcessDpiAwareness(True)

# creación ventana
ventana = Tk()
ventana.title("autoPersonas") # titulo de la ventana

# ajustar medida de la ventana, no se lee el titulo

#tkinter._test()

# creación de 2 categorias
frm_a = ttk.Frame() # frame a
frm_b = ttk.Frame() # frame b

#fred = Button(frm, fg="red", bg="blue")

# escribinos frame 1 con boton "quit"
label_a = ttk.Label(master = frm_a, text = "Vistos y considerando")
boton_a = ttk.Button(master = frm_a, text = "Quit", command = ventana.destroy)
label_a.pack()
boton_a.pack()

# escribinos frame 2 con celda para texto
label_b = ttk.Label(master = frm_b, text = "Name")
entry_b = ttk.Entry(master = frm_b)
label_b.pack()
entry_b.pack()

# escribirmos frames en la ventana de interfaz
frm_a.pack()
frm_b.pack()

# se abre ventana
ventana.mainloop()

# creacion word
document = Document() 
document.add_heading('Vitacura', 0)

# escribir el doc

# revisar la siguiente linea miercoles 4s sept
#document.add_paragraph(entry_b, style='Intense Quote')

# guardar el doc

document.save("test_autoPersona.docx") 

# al final debo hace un packing .exe