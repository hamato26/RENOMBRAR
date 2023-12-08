import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import fitz  # PyMuPDF

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

            for index, row in df_excel.iterrows():
                numero_identificacion = str(row['Número de identificación'])
                carpeta_numero_identificacion = os.path.join(self.carpeta_destino, numero_identificacion)

                if not os.path.exists(carpeta_numero_identificacion):
                    os.makedirs(carpeta_numero_identificacion)

                for pdf_file in os.listdir(self.carpeta_pdf):
                    if pdf_file.endswith('.pdf'):
                        pdf_path = os.path.join(self.carpeta_pdf, pdf_file)

                        # Usar PyMuPDF para obtener el texto del PDF
                        pdf_document = fitz.open(pdf_path)
                        texto_pdf = ""
                        for pagina in pdf_document.pages():
                            texto_pdf += pagina.get_text()

                        # Verificar si el número de identificación está presente en el texto
                        if numero_identificacion in texto_pdf:
                            shutil.copy(pdf_path, carpeta_numero_identificacion)

            messagebox.showinfo("Proceso completado", "Los archivos se han organizado en carpetas según los números de identificación del Excel.")

        except Exception as e:
            messagebox.showerror("Error", f"Error al organizar los archivos: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OrganizadorPDF(root)
    root.mainloop()
