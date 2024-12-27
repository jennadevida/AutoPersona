from cx_Freeze import setup, Executable
from setuptools import find_packages

# Opciones de construcción
build_options = {
    "packages": ["os", "sys", "re", "ctypes", "locale", "pandas", "tkinter", 
                 "tkcalendar", "datetime", "docx", "ttkbootstrap", "csv"],
    "includes": ["tkinter.filedialog", "tkinter.messagebox", "tkinter.ttk", "tkcalendar", "ttkbootstrap.constants", "ttkbootstrap.widgets"],
    "include_files": ["logo-vitacura_icono.ico", "logos-vitacura_sineslogan_vert.png", ("clausulas_csv", "clausulas_csv")] # parentesis incluye carpeta
}


# Detalles del ejecutable
executables = [
    Executable(
        script="main.py",            # Archivo principal
        target_name="mi_app_autopersonas.exe",    # Nombre del ejecutable
        base=None,             # Para GUI con consola
        icon="logo-vitacura_icono.ico",  # Ícono del ejecutable
    )
]

# Configuración de instalación
setup(
    name="Mi Aplicación",
    version="1.12",
    description="Automatizador para la elaboración de decretos y contratos a honorarios de la Municipalidad de Vitacura. Versión 27 dic 2024",
    options={"build_exe": build_options},
    packages=find_packages(),  # Descubre automáticamente los paquetes
    executables=executables,
)
