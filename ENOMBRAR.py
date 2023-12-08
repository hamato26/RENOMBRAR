import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import fitz  # PyMuPDF
from shutil import copyfile

class OrganizadorPDF:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador de PDF")

        self.carpeta_pdf = ""
        self.archivo_excel = ""
        self.carpeta_destino = ""

        self.label1 = tk.Label(root, text="Carpeta PDF:")
        self.label1.pack()

        self.button1 = tk.Button(root, text="Seleccionar Carpeta PDF", command=self.seleccionar_carpeta_pdf)
        self.button1.pack()

        self.label2 = tk.Label(root, text="Archivo Excel:")
        self.label2.pack()

        self.button2 = tk.Button(root, text="Seleccionar Archivo Excel", command=self.seleccionar_archivo_excel)
        self.button2.pack()

        self.label3 = tk.Label(root, text="Carpeta Destino:")
        self.label3.pack()

        self.button3 = tk.Button(root, text="Seleccionar Carpeta Destino", command=self.seleccionar_carpeta_destino)
        self.button3.pack()

        self.organizar_button = tk.Button(root, text="Organizar PDFs", command=self.organizar_pdfs)
        self.organizar_button.pack()

    def seleccionar_carpeta_pdf(self):
        self.carpeta_pdf = filedialog.askdirectory()
        self.label1.config(text=f"Carpeta PDF: {self.carpeta_pdf}")

    def seleccionar_archivo_excel(self):
        self.archivo_excel = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx;*.xls")])
        self.label2.config(text=f"Archivo Excel: {self.archivo_excel}")

    def seleccionar_carpeta_destino(self):
        self.carpeta_destino = filedialog.askdirectory()
        self.label3.config(text=f"Carpeta Destino: {self.carpeta_destino}")

    def organizar_pdfs(self):
        try:
            # Cargar el archivo Excel
            df_excel = pd.read_excel(self.archivo_excel, sheet_name="Hoja1")

            for _, row in df_excel.iterrows():
                numero_identificacion = str(row['Número de identificación'])

                # Filtrar archivos PDF que contienen el número de identificación
                pdfs_filtrados = [pdf for pdf in os.listdir(self.carpeta_pdf) if pdf.endswith('.pdf') and numero_identificacion in self.obtener_texto_pdf(os.path.join(self.carpeta_pdf, pdf))]

                if pdfs_filtrados:
                    # Seleccionar el primer archivo PDF filtrado y obtener el nuevo nombre
                    pdf_seleccionado = pdfs_filtrados[0]
                    nuevo_nombre = f"{row['Nº pers.']}.pdf"
                    
                    # Ruta del nuevo archivo PDF en la carpeta de destino
                    nuevo_path = os.path.join(self.carpeta_destino, nuevo_nombre)

                    try:
                        # Copiar y renombrar el archivo PDF seleccionado
                        copyfile(os.path.join(self.carpeta_pdf, pdf_seleccionado), nuevo_path)
                        print(f"Archivo {pdf_seleccionado} copiado y renombrado como {nuevo_path}")
                    except Exception as copy_error:
                        print(f"Error al copiar y renombrar {pdf_seleccionado}: {copy_error}")

            messagebox.showinfo("Proceso completado", "Los archivos PDF se han renombrado según los Nº pers. de la columna 'D' del Excel.")

        except Exception as e:
            messagebox.showerror("Error", f"Error al organizar los archivos: {str(e)}")

    def obtener_texto_pdf(self, pdf_path):
        try:
            pdf_document = fitz.open(pdf_path)
            texto_pdf = ""
            for pagina in pdf_document.pages():
                texto_pdf += pagina.get_text()
            return texto_pdf
        except Exception as e:
            print(f"Error al obtener texto del PDF {pdf_path}: {str(e)}")
            return ""

if __name__ == "__main__":
    root = tk.Tk()
    app = OrganizadorPDF(root)
    root.mainloop()
