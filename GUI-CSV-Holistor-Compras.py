#!/usr/bin/python3
import tkinter as tk
import tkinter.ttk as ttk
from BIN.ProcesarCSVHolistorCompras import Procesar_CSV
from tkinter.filedialog import askdirectory

def Cafecito():
    import webbrowser
    webbrowser.open_new("https://cafecito.app/abustos")

def Carpeta():
    Carpeta = askdirectory()
    import os
    Archivos = os.listdir(Carpeta)
    Archivos = [Carpeta + "/" + i for i in Archivos]
    Procesar_CSV(Archivos)

def Abrir_Excel():
    import os
    os.system("start EXCEL.EXE Lista-Archivos-Compras.xlsx")
    
def Procesar_Excel():
    import pandas as pd
    df = pd.read_excel("Lista-Archivos-Compras.xlsx")
    df = df[df['Raiz'].notna()]
    Archivos = df['Lista'].tolist()
    Procesar_CSV(Archivos)


class ModeloPygubuApp:
    def __init__(self, master=None):
        # build ui
        Toplevel_1 = tk.Tk() if master is None else tk.Toplevel(master)
        Toplevel_1.configure(
            background="#2e2e2e",
            cursor="arrow",
            height=250,
            width=335)
        Toplevel_1.iconbitmap("BIN/ABP-blanco-en-fondo-negro.ico")
        Toplevel_1.overrideredirect("False")
        Toplevel_1.resizable(False, False)
        Toplevel_1.title("CSV a Holistor")
        # First object created
        #self.setup_ttk_styles(Toplevel_1)

        Label_3 = ttk.Label(Toplevel_1)
        self.img_ABPblancoenfondonegro111 = tk.PhotoImage(
            file="BIN/ABP-blanco-sin-fondo .png")
        Label_3.configure(
            background="#2e2e2e",
            image=self.img_ABPblancoenfondonegro111)
        Label_3.pack(side="top")
        Label_1 = ttk.Label(Toplevel_1)
        Label_1.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            takefocus=False,
            text='Procesar CSV del Portal IVA al papel de trabajo de importación de Holistor\n',
            wraplength=320)
        Label_1.pack(expand=True, side="top")
        Label_2 = ttk.Label(Toplevel_1)
        Label_2.configure(
            background="#2e2e2e",
            foreground="#ffffff",
            justify="center",
            text='por Agustín Bustos Piasentini\nhttps://www.Agustin-Bustos-Piasentini.com.ar/')
        Label_2.pack(expand=True, side="top")
        self.Carpeta = ttk.Button(Toplevel_1)
        self.Carpeta.configure(text='Seleccionar Carpeta con CSV', command=Carpeta)
        self.Carpeta.pack(expand=True, pady=4, side="top")
        self.Abrir_Excel = ttk.Button(Toplevel_1)
        self.Abrir_Excel.configure(text='Abrir Excel con Ubicaciones', command=Abrir_Excel)
        self.Abrir_Excel.pack(expand=True, pady=4, side="top")
        self.Procesar_Excel = ttk.Button(Toplevel_1)
        self.Procesar_Excel.configure(text='Procesar Lista del Excel' , command=Procesar_Excel)
        self.Procesar_Excel.pack(expand=True, pady=4, side="top")
        self.Donación = ttk.Button(Toplevel_1)
        self.Donación.configure(text='Colaboraciones' , command=Cafecito)
        self.Donación.pack(pady=4, side="top")

        # Main widget
        self.mainwindow = Toplevel_1

    def run(self):
        self.mainwindow.mainloop()


if __name__ == "__main__":
    app = ModeloPygubuApp()
    app.run()
