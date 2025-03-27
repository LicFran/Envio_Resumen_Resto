import PyInstaller.__main__
import os

# Obtener la ruta del directorio actual
current_dir = os.path.dirname(os.path.abspath(__file__))

# Configurar los argumentos para PyInstaller
PyInstaller.__main__.run([
    'enviar_resumen.py',
    '--onefile',
    '--name=Enviar Resumen',
    '--add-data=config.py;.',
    '--icon=aula-origen2.ico',
    '--clean',
    '--windowed',
]) 