# -*- coding: utf-8 -*-
#ByEricRoca

#=====Llibreries=====#
import os
from tkinter import *
from tkinter import messagebox
from tkinter import ttk

try:
  import xlrd 
except ImportError:
    print("Error: Panda module not found")
    print("Trying to Install required module: openpyxl")
    os.system('python -m pip install xlrd==1.2.0')
    os.system('python -m pip install openpyxl')
    os.system('python -m pip install pandas')
    import xlrd 

#=====Classes=====#
class Fotografia:
    
    def __init__(self, ruta, nom, id):
        self.ruta = ruta
        self.nom = nom
        self.id = id
        
class Aplicacio(object):
    def __init__(self):
        self.root = Tk()
        self.root.title("Canviador de noms d'arxius v1.1")
        
        self.root.geometry("520x230")
        self.root.resizable(0, 0)

        self.rutaFoto = StringVar()
        self.rutaExcel = StringVar()

        self.labelFotos = Label(
            self.root,
            text="Ruta fotografies:",
            fg='black'
        )

        self.labelExcel = Label(
            self.root,
            text="Ruta excel:",
            fg='black'
        )

        self.labelFotos.place(x=10, y=30)

        self.labelExcel.place(x=10, y=110)

        self.botoAcceptar = ttk.Button(self.root, text="Canvia!", command= lambda: self.mainProgram())

        self.caixaFoto = Entry(self.root, textvariable=self.rutaFoto, width=80).place(x=10, y=50)

        self.caixaExcel = Entry(self.root, textvariable=self.rutaExcel, width=80).place(x=10, y=130)

        self.botoAcceptar.place(x=410, y=190)

        self.root.mainloop()

    def mainProgram(self):
            ruta = self.rutaFoto.get()
            excel = self.rutaExcel.get()

            if ruta == "" or excel == "":
                messagebox.showinfo(title="Alerta!", message="Ruta de fotos o ruta d'excel vuit.")
            elif messagebox.askquestion(message="Les rutes son correctes?", title="Atenció") == "yes":
                
                fotos = cercaFotos(ruta)
                
                if fotos != False:
                    searchFilesName(excel, fotos)
                
                else:
                    messagebox.showerror(title="Error", message="Ruta no trobada!")


def searchFilesName(excData, fotos):
    #i = 0
    try:
        data = xlrd.open_workbook(excData)
        sheet = data.sheet_by_index(0)
        for sheet_i in range(sheet.nrows):
            for foto in fotos:
                if foto.nom[:-4] == "":
                    print(f"{foto.nom[:-4]} == {sheet.row_values(sheet_i)[0]}")
                
                if foto.nom[:-4] == sheet.row_values(sheet_i)[0]:
                    changeFilesName(foto, str(int(sheet.row_values(sheet_i)[1])))
            
            #i += 1

        messagebox.showinfo(title="Ok", message="Fotografies canviades!")
    
    except FileNotFoundError as e:
        print(e)
        messagebox.showerror(title="Error", message="Ruta no trobada!")
    
    #print(i)

def changeFilesName(foto, ralc):
    old_name = foto.ruta + "\\" + foto.nom
    new_name = foto.ruta + "\\" + ralc + ".jpg"
    print(foto.ruta + "\\" + ralc + ".jpg")
    try:
        os.rename(old_name, new_name)
    except FileExistsError:
        if messagebox.askquestion(message=f"Nom arxiu: {foto.nom} i NIE: {ralc} ya existeix. \n Vols substituir-ho?", title="Alerta!") == "yes":
            try:
                os.remove(new_name)
                os.rename(old_name, new_name)
            
            except:
                print("Error, no s'ha pogut substituir el nom.")

def cercaFotos(ruta):
    for base, dirs, files in os.walk(ruta):
        fotos = []

        for foto in files:
            fotos.append(Fotografia(base,foto,""))

    try:
        return fotos
    
    except Exception as e:
        print(e)
        return False



#== Main program ==#
mainCode = lambda: Aplicacio()

mainCode()

#Directori fotos: C:\Users\socle\Desktop\Practicas Castellet\1. Oficines\FOTOS JPG
#Excel: C:\Users\socle\Desktop\Practicas Castellet\1. Oficines\Fotos Eso -Bat.xlsx