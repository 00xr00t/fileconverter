import pdfkit
import os
import jpype
from win32com import client
import tkinter
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as filedialog
from tkinter.messagebox import showinfo
import time
import re

pencere = tk.Tk()
pencere.title("HTML2PDF")
pencere.geometry("500x250")

        
def secimyap_html():
        pencere.filename = filedialog.askopenfilename(title="HTML Dosyası Aç:", defaultextension=".html",filetypes=[("HTML Documents","*.html")])
        if pencere.filename == "":
                showinfo(title="Hata!", message="Lütfen Dosya Seçin.")
        else:
            showinfo(title="Seçiğiniz Dosya: ", message=pencere.filename)
            pencere.filename.endswith(".html")
            label1.config(text=pencere.filename)
        
def convert_html_pdf():
        new_file = os.path.splitext(pencere.filename)[0]
        kayit_yeri = filedialog.asksaveasfilename(initialfile=new_file, defaultextension=".txt", filetypes=[("Pdf Files","*.pdf")])
        pdfkit.from_file(new_file+".html", kayit_yeri)
        for i in range(5):
                pencere.update_idletasks()
                pb1['value'] += 50
        
                time.sleep(0.2)

def secimyap_excel():
        pencere.filename_excel = filedialog.askopenfilename(title="Excel Dosyası Aç:", defaultextension=".xlsx",filetypes=[("Excel Files","*.xlsx"),("Excel Files", "*.xls")])
        if pencere.filename_excel == "":
            showinfo(title="Hata!", message="Lütfen Dosya Seçin.")
        else:
            showinfo(title="Seçiğiniz Dosya: ", message=pencere.filename_excel)
            #pencere.filename_excel.endswith(".xlsx")
            label1.config(text=pencere.filename_excel)

def convert_excel_pdf():
        new_file_excel = os.path.splitext(pencere.filename_excel)[0]
        kayit_yeri_excel = filedialog.asksaveasfilename(initialfile=new_file_excel, defaultextension=".pdf", filetypes=[("PDF Files","*.pdf",)])
        
        excel = client.Dispatch("Excel.Application")
        sheets = excel.Workbooks.Open(pencere.filename_excel)
        work_sheets = sheets.Worksheets[0]
        kayit_yeri_excel = re.sub("[\s]", "_", kayit_yeri_excel)                
        work_sheets.ExportAsFixedFormat(0,kayit_yeri_excel)
        excel.Application.Quit()
        for i in range(5):
                pencere.update_idletasks()
                pb1['value'] += 20
                time.sleep(0.2)
def kapat():
        showinfo(message="Çıkış Yapılıyor...")
        pencere.destroy()

def size_1():
   label2.config(font=('Helvatical bold',15))
def size_2():
   label2.config(font=('Helvatical bold',12))


label2 = ttk.Label(pencere)
label2.config(text="\tHTML2PDF Programına Hoşgeldiniz", background="white", command= size_1())
label2.pack()
label2.place(width=1000)

label2 = ttk.Label(pencere)
label2.config(text="signed by R3TREX",command = size_2())
label2.pack()
label2.place(x=358,y=228)


label1 = ttk.Label(pencere)
label1.config()
label1.pack()
label1.place(x=110,y=40)

buton1 = ttk.Button(pencere)
buton1.config(text="HTML Dosyası Seç",  command=secimyap_html)
buton1.pack()
buton1.place(x=100,y=70,width=150,height=30)
        
buton = ttk.Button(pencere)
buton.config(text="HTML --> PDF Kaydet ", command= lambda : convert_html_pdf())
buton.pack()
buton.place(x=100,y=100,width=150,height=30)

buton2 = ttk.Button(pencere)
buton2.config(text="EXCEL --> PDF Kaydet ", command= lambda : convert_excel_pdf())
buton2.pack()
buton2.place(x=250,y=100,width=150,height=30)

buton3 = ttk.Button(pencere)
buton3.config(text="EXCEL Dosyası Seç", command=secimyap_excel)
buton3.pack()
buton3.place(x=250,y=70,width=150,height=30)

pb1 = ttk.Progressbar(pencere, orient="horizontal", length=300, mode='determinate',)
pb1.pack(expand=True)
pb1.place(x=100,y=150)

buton4 = ttk.Button(pencere)
buton4.config(text="Programı Kapat", command= kapat)
buton4.pack()
buton4.place(x=200,y=200)



pencere.mainloop()


