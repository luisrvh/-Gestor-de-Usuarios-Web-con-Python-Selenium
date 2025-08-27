Gestor de Usuarios Web con Python + Selenium

Este proyecto es una herramienta de automatización que permite realizar acciones en masa (como cambiar contraseñas) para usuarios listados en un archivo Excel, dentro de una página web personalizada. 
Puede adaptarse fácilmente a cualquier plataforma web modificando los elementos y selectores en el código.


Requisitos

- Python 3.9 o superior
- Navegador Microsoft Edge (o Chrome/Firefox si modificas el código)
- Driver del navegador (`msedgedriver.exe`) en la misma carpeta del script
- Archivos requeridos:
  - `usuarios.xlsx` con una columna llamada `Username`
  - Script `gestor_usuarios.py`


Instalación de dependencias

Ejecuta esto en tu terminal para instalar los paquetes necesarios:

bash pip install selenium pandas numpy openpyxl

Archivos necesarios
Archivo	Descripción
gestor_usuarios.py	Script principal que automatiza el proceso
usuarios.xlsx	Excel con los usuarios a procesar (Username)
msedgedriver.exe	Driver del navegador Microsoft Edge

Para adaptar este proyecto a cualquier otra página web, necesitas modificar los siguientes elementos en el script gestor_usuarios.py:

La URL del sitio web: driver.get('https://tusitio.com/tu-pagina')

Selectores HTML como ID o XPATH:

Ejemplo de selector a modificar:

driver.find_element(By.ID, "selectorValue").send_keys(usuario)


Debes inspeccionar el sitio web que deseas automatizar y reemplazar "selectorValue" y otros valores por los ID, nombres o rutas reales de los campos o botones que uses.

Cómo ejecutar

Desde la terminal (dentro de la carpeta del proyecto):

python gestor_usuarios.py

Se abrirá una ventana donde eliges una versión del proceso.

Ingresa la contraseña base.

Selecciona si quieres usar la misma contraseña para todos los usuarios o una diferente para cada uno (agregando un número incremental).

Pulsa el botón para iniciar el proceso.

Resultado

Al finalizar el proceso, se genera un archivo nuevo:

usuarios_actualizados.xlsx: contiene los usuarios procesados junto con sus nuevas contraseñas.

 Características

✔Interfaz gráfica con Tkinter

✔Compatible con Edge (modificable para otros navegadores)

✔Fácil de personalizar para otras plataformas

✔Soporte para contraseñas incrementales por usuario

Advertencia

Este proyecto es solo para fines educativos o de uso interno autorizado. No utilices esta herramienta para automatizar accesos sin permiso a sitios web de terceros.

