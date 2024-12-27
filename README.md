# README

## Descripción
Este proyecto es una aplicación de escritorio desarrollada en Python utilizando la biblioteca tkinter. La aplicación está diseñada para automatizar la creación de documentos de decretos y contratos a honorarios para la Municipalidad de Vitacura. La interfaz permite ingresar información sobre personas y cargos, y genera documentos en formato Word con los datos proporcionados.

## Requisitos
Python 3.x

## Bibliotecas
tkinter
ttkbootstrap
pandas
docx
dateutil
ctypes
os
sys
re
datetime
locale

Nota: tkinter, ctypes, csv, os, sys, re, datetime, y locale son módulos estándar de Python y no necesitan ser instalados.

### Instalación de Dependencias
Para instalar las dependencias especificadas en el archivo requirements.txt, utiliza el siguiente comando dentro del directorio correspondiente:
    pip install -r requirements.txt

## Clases Principales

*Clase Persona*
Esta clase almacena la información de una persona, incluyendo nombre, RUT, género, domicilio, correo electrónico, nacionalidad, estado civil, profesión, y beneficios seleccionados.

*Clase Cargo*
Esta clase almacena la información de un cargo, incluyendo tipo de contrato, programa, departamento, dirección, renta, fechas, y otros detalles específicos del cargo.

*Clase Aplicacion*
Esta es la clase principal que maneja la interfaz de usuario y la lógica de la aplicación. Configura la ventana principal, inicializa las variables y crea los widgets necesarios.

## Funcionalidades
Inicialización de la Interfaz: Configura la ventana principal y ajusta su tamaño y posición.
Creación de Widgets: Crea y organiza los widgets en la interfaz.
Verificación de RUT: Verifica la validez del RUT ingresado.
Validación de Correo Electrónico: Valida el formato del correo electrónico ingresado.
Generación de Documentos: Genera documentos en formato Word con la información ingresada.

## Contribuciones
Las contribuciones son bienvenidas. Si deseas contribuir: https://github.com/jennadevida/AutoPersona.git

## Licencia
Este proyecto está licenciado bajo la Licencia GNU General Public License v3.0. Consulta el archivo LICENSE para más detalles.

## Contacto
Para cualquier consulta o sugerencia, por favor contacta a JK Fienco: jkfienco@gmail.com
