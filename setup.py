from cx_Freeze import setup, Executable
from setuptools import find_packages

# Opciones de construcción
build_options = {
    "packages": ["os", "sys", "ctypes", "locale", "pandas", "tkinter", 
                 "tkcalendar", "datetime", "docx"],
    "includes": ["tkinter.filedialog", "tkinter.messagebox", "tkinter.ttk", "tkcalendar"],
    "include_files": ["logo-vitacura_icono.ico", "logos-vitacura_sineslogan_vert.png"]
}


# Detalles del ejecutable
executables = [
    Executable(
        script="interfaz_con_clases.py",            # Archivo principal
        target_name="mi_app_autopersonas.exe",    # Nombre del ejecutable
        base="Win32GUI",             # Para GUI sin consola
        icon="logo-vitacura_icono.ico",  # Ícono del ejecutable
    )
]

# Configuración de instalación
setup(
    name="Mi Aplicación",
    version="1.1",
    description="Automatizador para la elaboración de decretos y contratos a honorarios de la Municipalidad de Vitacura",
    options={"build_exe": build_options},
    packages=find_packages(),  # Descubre automáticamente los paquetes
    executables=executables,
)
