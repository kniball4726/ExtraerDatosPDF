# ExtraerDatosPDF
Aplicacion para extracción de datos dentro de documentos

1.- Se debe crear un ambiente virtual de python

`` python -m venv .venv ``

2.- Se debe activar el ambiente virtual

`` cd .venv\Scripts ``

`` .\activate ``

una vez activado en la consola debe aparecer al inicio de la ruta (.venv)

se debe volver a la carpeta raiz 

`` cd ../.. ``

3.- Se deben instalar las dependencias que aparecen en requirements.txt.

`` pip install -r requirements.txt``

puede que python solicite actualizar pip, en tal caso se debe actualizar con

`` python.exe -m pip install --upgrade pip ``

luego ejecutar de nuevo 

`` pip install -r requirements.txt``

Para crear el ejecutable bien sea desde Windows, linux o mac, se debe ejecutar

`` pyinstaller --onefile .\descontar.py ``

esto genera una serie de carpetas, de las cuales dentro de la carpeta dist se va a encontrar un ejecutable o archivo descontar.exe
este archivo es el que tiene todo el script necesario y es el que vamos a ejecutar dentro de la carpeta Remitos o donde sea que se alojen los remitos que desea descontar
