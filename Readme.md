# App Facturas
"""
Este código crea una aplicación GUI utilizando la biblioteca tkinter. Tiene cuatro botones:
1. "Crear una nueva base de datos" - Este botón permite al usuario crear una nueva base de datos seleccionando un archivo de Excel, teniendo como llave primaria el campo "num_fac". Luego, los datos del archivo de Excel se importan a la base de datos.
2. "Actualizar la base de datos" - Este botón permite al usuario actualizar una base de datos existente seleccionando un archivo de Excel. A continuación, los datos del archivo de Excel se actualizan en la base de datos; teniendo en cuenta de que si el registro existe, se actualizan solo los campos: "fec_pag", "num_pag" y "diasconv", pero si el registro no existe se agrega todo el registro a la base de datos.
3. "Exportar los registros a un archivo xls" - Este botón permite al usuario exportar todos los registros de la base de datos a un archivo de Excel.
4. "Salir": este botón cierra la ventana de la aplicación y sale de ella.
"""