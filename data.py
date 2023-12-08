import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import fitz  # Importar desde PyMuPDF

def encontrar_paginas_pdf(pdf_path, excel_path):
    try:
        # Cargar datos desde la hoja "Hoja1" del archivo Excel
        df_excel = pd.read_excel(excel_path, sheet_name='Hoja1')

        # Crear un diccionario para almacenar los resultados
        resultados = {'Número de identificación': [], 'Número de Página': []}

        # Abrir el archivo PDF
        with fitz.open(pdf_path) as pdf_document:
            # Iterar sobre cada fila del DataFrame de Excel
            for index, row in df_excel.iterrows():
                numero_id = str(row['Número de identificación'])  # Convertir a cadena

                # Buscar el número de identificación en el PDF
                for pagina in range(pdf_document.page_count):
                    page = pdf_document[pagina]
                    texto = page.get_text()

                    if numero_id in texto:
                        resultados['Número de identificación'].append(numero_id)
                        resultados['Número de Página'].append(pagina + 1)  # Sumar 1 porque las páginas comienzan desde 1

        # Crear un nuevo DataFrame con los resultados
        df_resultados = pd.DataFrame(resultados)

        # Guardar los resultados en un nuevo archivo de Excel
        df_resultados.to_excel('resultados_excel.xlsx', index=False)

        messagebox.showinfo('Proceso Completado', 'La búsqueda se ha completado con éxito. Los resultados se han guardado en "resultados_excel.xlsx".')
    except Exception as e:
        messagebox.showerror('Error', f'Se produjo un error: {str(e)}')

def abrir_archivo():
    ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")])
    entry_archivo.delete(0, tk.END)
    entry_archivo.insert(0, ruta_archivo)

def abrir_pdf():
    ruta_pdf = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")])
    entry_pdf.delete(0, tk.END)
    entry_pdf.insert(0, ruta_pdf)

# Crear la interfaz gráfica
root = tk.Tk()
root.title('Buscar Números de Identificación')

# Etiqueta y entrada para el archivo Excel
label_archivo = tk.Label(root, text='Archivo Excel:')
label_archivo.grid(row=0, column=0, padx=10, pady=10, sticky='e')
entry_archivo = tk.Entry(root, width=50)
entry_archivo.grid(row=0, column=1, padx=10, pady=10)
button_examinar_excel = tk.Button(root, text='Examinar', command=abrir_archivo)
button_examinar_excel.grid(row=0, column=2, padx=10, pady=10)

# Etiqueta y entrada para el archivo PDF
label_pdf = tk.Label(root, text='Archivo PDF:')
label_pdf.grid(row=1, column=0, padx=10, pady=10, sticky='e')
entry_pdf = tk.Entry(root, width=50)
entry_pdf.grid(row=1, column=1, padx=10, pady=10)
button_examinar_pdf = tk.Button(root, text='Examinar', command=abrir_pdf)
button_examinar_pdf.grid(row=1, column=2, padx=10, pady=10)

# Botón para iniciar la búsqueda
button_buscar = tk.Button(root, text='Buscar', command=lambda: encontrar_paginas_pdf(entry_pdf.get(), entry_archivo.get()))
button_buscar.grid(row=2, column=1, pady=20)

# Ejecutar la interfaz gráfica
root.mainloop()
