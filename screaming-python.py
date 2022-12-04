import tkinter
from tkinter import *
import glob, os
import subprocess
import pandas as pd
from pandas import DataFrame, ExcelWriter

ventana = tkinter.Tk()
ventana.geometry("720x540") #tamaño de ventana

NameAplicacion = tkinter.Label(ventana, text="Queseo.es", bg="red")
NameAplicacion.pack(fill=tkinter.X)

#Definimos URL que queremos analizar
EtiquetaUrl = tkinter.Label(ventana, text="Inserta la Url a rastrear:")
EtiquetaUrl.pack()

url_entry_text = tkinter.StringVar()
urlbox = tkinter.Entry(ventana, textvariable=url_entry_text)
urlbox.pack(fill=tkinter.X)
url_entry_text.set( "https://www.queseo.es/" )

#Definimos nuesta ventana de salida
EtiquetaDir = tkinter.Label(ventana, text="Inserta la ruta de destino:")
EtiquetaDir.pack()

dir_entry_text = tkinter.StringVar()
Dirbox = tkinter.Entry(ventana, textvariable=dir_entry_text)
Dirbox.pack(fill=tkinter.X)
dir_entry_text.set( "D:/descargas/screaming/screaming" )

#Definimos qué SO utilizamos para luego identificar dónde está nuestro screaming
OPTIONS = ["Windows 64bits","Windows 32bits","Mac"] #etc
variable = StringVar(ventana)
variable.set(OPTIONS[0]) # default value

w = OptionMenu(ventana, variable, *OPTIONS)
w.pack(fill=tkinter.X)

#prueba de checkobox
#Internal_HTML = tkinter.Checkbutton(ventana, text="Internal:HTML", variable="Internal:HTML")
#Internal_HTML.pack()

#función de crawleo
def crawlearweb():
        #Condicional en función del SELECT. Cambiamos la ubicación para luego ejecutar el comando de screaming
        if variable.get()=="Windows 64bits":
            os.chdir('C:\Program Files (x86)\Screaming Frog SEO Spider')
        elif (variable.get()=="Windows 32bits"):
            os.chdir('C:\Program Files\Screaming Frog SEO Spider')
        else:
            os.chdir('cd ~/.ScreamingFrogSEOSpider/')
        
        #Ejecutamos screaming
        stream = subprocess.Popen('ScreamingFrogSEOSpiderCli.exe --crawl '+urlbox.get()+' -headless --save-crawl --output-folder "'+Dirbox.get()+'/" --export-tabs "Internal:HTML,Page Titles:Same as H1,H1:Multiple" --bulk-export "Images:Images Missing Alt Text Inlinks,Images:Images over X KB Inlinks,Response Codes:Blocked by Robots.txt Inlinks,Response Codes:Internal & External:Blocked Resource Inlinks,Response Codes:Internal & External:No response Inlinks,Response Codes:Internal & External:Redirection (3xx) Inlinks,Response Codes:Internal & External:Client Error (4xx) Inlinks,Response Codes:Internal & External:Server Error (5xx) Inlinks,Security:Mixed Content" --export-format csv',shell=True)
        stream.wait()

        print("Crawl Terminado. Creando excel")

        writer = ExcelWriter(Dirbox.get()+"/setup.xlsx") #Aquí va la ruta donde queremos guardar el excel y su nombre

        for filename in glob.glob(Dirbox.get()+"/*.csv"): #Aquí va la ruta donde se guardan todos los csv extraidos con el screaming por comandos
            df_csv = pd.read_csv(filename)

            (_, f_name) = os.path.split(filename)
            (f_shortname, _) = os.path.splitext(f_name)

            df_csv.to_excel(writer, f_shortname, index=False)

        writer.save()

        print("Excel terminado ;)")

boton1 = tkinter.Button(ventana, text="Crawlear", bg="red", command=crawlearweb)
boton1.pack(side=tkinter.BOTTOM)


ventana.mainloop()
