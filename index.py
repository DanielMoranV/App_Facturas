import sqlite3
import xlrd
import xlwt
import tkinter as tk
from tkinter import filedialog
import datetime

class FacturacionApp:
    def __init__(self, master):
        self.master = master
        master.title("Facturación App")
        master.config(bg="#00A3E0")
        master.geometry("350x300")  # establecer tamaño fijo de ventana

        # agregando estilo al Label
        self.label = tk.Label(master, text="Seleccione una opción:", bg="#00A3E0", fg="#fff", font=("Arial", 16))
        self.label.pack(pady=10)

        # agregando estilo a los botones
        button_width = 30
        button_height = 2
        self.nuevo_button = tk.Button(master, text="Crear una nueva base de datos", bg="#fff", fg="#00A3E0", font=("Arial", 12), command=self.nuevo, width=button_width, height=button_height)
        self.nuevo_button.pack(pady=5)

        self.actualizar_button = tk.Button(master, text="Actualizar la base de datos", bg="#fff", fg="#00A3E0", font=("Arial", 12), command=self.actualizar, width=button_width, height=button_height)
        self.actualizar_button.pack(pady=5)

        self.exportar_button = tk.Button(master, text="Exportar los registros a un archivo xls", bg="#fff", fg="#00A3E0", font=("Arial", 12), command=self.exportar, width=button_width, height=button_height)
        self.exportar_button.pack(pady=5)

        self.quit_button = tk.Button(master, text="Salir", bg="#fff", fg="#00A3E0", font=("Arial", 12), command=master.quit, width=button_width, height=button_height)
        self.quit_button.pack(pady=5)

    def nuevo(self):
        archivo = self.abrir_archivo()
        if archivo:
            # Abrir el archivo xls
            libro = xlrd.open_workbook(archivo)
            # Obtener la primera hoja del archivo
            hoja = libro.sheet_by_index(0)
            # Obtener los nombres de las columnas
            columnas = hoja.row_values(0)
            # Crear una conexión a la base de datos
            conexion = sqlite3.connect('db_csr.db')
            # Crear un cursor para ejecutar comandos SQL
            cursor = conexion.cursor()
            # Crear la tabla en la base de datos
            cursor.execute('CREATE TABLE IF NOT EXISTS datos (' + ', '.join(columnas) + ', PRIMARY KEY(num_fac))')
            # Insertar los datos en la tabla
            for i in range(1, hoja.nrows):
                fila = hoja.row_values(i)
                cursor.execute('INSERT INTO datos VALUES (' + ', '.join('?' * len(columnas)) + ')', fila)
            # Cerrar la conexión a la base de datos
            conexion.commit()
            conexion.close()
            tk.messagebox.showinfo("Información", "La base de datos se ha creado exitosamente.")
        else:
            tk.messagebox.showerror("Error", "No se ha seleccionado ningún archivo.")
        
    def actualizar(self):
        archivo = self.abrir_archivo()
        if archivo:
            # Abrir el archivo xls
            libro = xlrd.open_workbook(archivo)
            # Obtener la primera hoja del archivo
            hoja = libro.sheet_by_index(0)
            # Obtener los nombres de las columnas
            columnas = hoja.row_values(0)
            # Crear una conexión a la base de datos
            conexion = sqlite3.connect('db_csr.db')
            # Crear un cursor para ejecutar comandos SQL
            cursor = conexion.cursor()
            # Recorrer las filas del archivo
            for i in range(1, hoja.nrows):
                fila = hoja.row_values(i)
                # Verificar si la factura ya existe en la base de datos
                cursor.execute('SELECT * FROM datos WHERE num_fac = ?', (fila[0],))
                existe = cursor.fetchone()
                # Si la factura no existe, insertarla en la base de datos
                if not existe:
                    cursor.execute('INSERT INTO datos (num_fac, fec_fac, importe, fec_pag, num_pag, cod_emp, cia, diasconv, diasfac, observ, nom_pac, nom_emp, fecha_envio, fecha_recepcion, observacion) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', fila + ['', '', ''])
                # Si la factura existe, actualizarla en la base de datos
                else:
                    cursor.execute('UPDATE datos SET fec_pag=?, num_pag=?, diasfac=? WHERE num_fac = ?', (fila[3], fila[4], fila[8], fila[0]))
            tk.messagebox.showinfo("Información", "Registros actualizados exitosamente.")
        # Cerrar la conexión a la base de datos
        conexion.commit()
        conexion.close()
    def exportar(self):
        # Crear una conexión a la base de datos
        conexion = sqlite3.connect('db_csr.db')
        # Crear un cursor para ejecutar comandos SQL
        cursor = conexion.cursor()
        # Obtener los datos de la tabla
        cursor.execute('SELECT * FROM datos')
        datos = cursor.fetchall()
        # Crear un libro de Excel
        libro = xlwt.Workbook()
        # Crear una hoja en el libro
        hoja = libro.add_sheet('datos')
        # Escribir los nombres de las columnas
        cursor.execute("PRAGMA table_info(datos)")
        columnas = [tupla[1] for tupla in cursor.fetchall()]
        # Cerrar la conexión a la base de datos
        conexion.close()
        # Definir los formatos de celda
        estilo_titulo = xlwt.easyxf('font: bold 1, color black;')
        for i in range(len(columnas)):
            hoja.write(0, i, columnas[i],estilo_titulo)
        # Escribir los datos
        for i in range(len(datos)):
            fila = datos[i]
            for j in range(len(fila)):
                hoja.write(i+1, j, (fila[j]))

        # Guardar el archivo xls
        archivo = filedialog.asksaveasfilename(defaultextension='.xls', filetypes=[('Archivo Excel', '*.xls')])
        if archivo:
            libro.save(archivo)
            tk.messagebox.showinfo("Información", "Los datos se han exportado exitosamente.")
        else:
            tk.messagebox.showerror("Error", "No se ha seleccionado ningún archivo.")
        
    def abrir_archivo(self):
        archivo = filedialog.askopenfilename(filetypes=[('Archivo Excel', '*.xls')])
        return archivo


root = tk.Tk()
facturacion_app = FacturacionApp(root)
root.mainloop()
