import customtkinter as ctk
from customtkinter import filedialog    
import json
import os
from pathlib import Path
import sys
from sys import platform
import pandas as pd
from dataclasses import dataclass
import numpy as np
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Mm
from docxcompose.composer import Composer
import io
import win32api
import win32print
import tempfile
from PIL import ImageTk



template_path = ""
icon_path = ""
tempDir = None
filesPath = []

labels = {}

@dataclass
class Label():
    productID : str
    previewPath: str
    units : int
    place : str

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.hasOrders = False
        self.hasPreviews = False

        self.previewsPath = "Nenahráno"

        self.ordersPath = ""

        self.title("Labels Creator")
        self.geometry("800x600")
        self.iconpath = ImageTk.PhotoImage(file=icon_path)
        self.wm_iconbitmap()
        self.iconphoto(False, self.iconpath)        

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.dataPath = f'{Path.home()}/AppData/Roaming/Štítky-Generátor/generator_stitku_data.json'
        self.data = self.getData()

        if "path" in self.data:
            self.previewsPath = self.data["path"]
            self.hasPreviews = True

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.createUI()

    def on_close(self):
        global tempDir
        try:
            tempDir.cleanup()
            self.destroy()
        except:
            self.destroy()

    def createUI(self):
        self.grid_rowconfigure((2,4), weight=1)
        self.grid_columnconfigure((0,1), weight=1)
        self.label = ctk.CTkLabel(self, text="Generátor štítků", font=("Roboto", 20))
        self.label.grid(row=0, column=0, padx=20, pady=20, sticky="ew", columnspan=2)

        self.button1 = ctk.CTkButton(self, text="Nahrát zakázky", command=self.upload, corner_radius=20)
        self.button1.grid(row=1, column=0, padx=20, pady=20, sticky="e")

        self.button2 = ctk.CTkButton(self, text="Generovat", command=self.main, corner_radius=20)
        self.button2.grid(row=1, column=1, padx=20, pady=20, sticky="w")
        #Úspěšně nahráno

        self.upload_frame = ctk.CTkFrame(self)
        self.upload_frame.grid(row=2, column=0, padx=150, pady=20, sticky="ew", columnspan=2)
        self.upload_frame.grid_columnconfigure((0,1,2), weight=1)

        self.viewMSG(self.upload_frame, "Nenahráno", 20, "red")


        self.button3 = ctk.CTkButton(self, text="Změnit adresu výkresů", command=self.edit_previews, corner_radius=20)
        self.button3.grid(row=3, column=0, padx=150, pady=50, sticky="ew", columnspan=2)

        self.previews_frame = ctk.CTkFrame(self)
        self.previews_frame.grid(row=4, column=0, padx=200, pady=0, sticky="ew", columnspan=2)
        self.previews_frame.grid_columnconfigure((0,1), weight=1)

        self.viewMSG(self.previews_frame, self.previewsPath, 15, "white")
    

    def viewMSG(self, frame, name, size, color):
        for item in frame.winfo_children():
            item.destroy()

        self.MSG = ctk.CTkLabel(frame, text=name, font=("Roboto", size), text_color=color, wraplength=320)
        self.MSG.grid(row=0, column=0, padx=20, pady=20, sticky="ew", columnspan=3)

    def viewPrint(self):

        printers = [x[2] for x in list(win32print.EnumPrinters(2))] 
        def_printer = win32print.GetDefaultPrinter()

        for item in self.upload_frame.winfo_children():
            item.destroy()

        index = 0

        for type in labels.keys():
            self.MSG = ctk.CTkButton(self.upload_frame, text=f"Tisk {type}", command=lambda idx=index: printLabels(self, idx), corner_radius=20)
            self.MSG.grid(row=0, column=index, padx=15, pady=20, sticky="ew")
            index += 1
        '''
        self.printTitle = ctk.CTkLabel(self.upload_frame, text="Vybrat tiskárnu", font=("Roboto", 15), text_color="white", wraplength=320)
        self.printTitle.grid(row=1, column=0, padx=150, pady=20, sticky="ew", columnspan=3)
        self.combobox = ctk.CTkComboBox(self.upload_frame, values=printers, corner_radius=20)
        self.combobox.set(def_printer)
        self.combobox.grid(row=2, column=0, padx=70, pady=20, sticky="ew", columnspan=3
        '''

    def getData(self):
        if os.path.exists(self.dataPath):
            try:
                with open(self.dataPath, "r") as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def upload(self):
        filePath = filedialog.askopenfilename(initialdir="/", filetypes=[('Tabulky', '.xlsx .xsl .xlsm .odf .ods .odt')])
        try:
            self.ordersPath = filePath
            self.hasOrders = True
            self.viewMSG(self.upload_frame, "Úspěšně nahráno", 20, "green")
        except:
            self.viewMSG(self.upload_frame, "Chyba, soubor nebyl nahrán", 20, "red")

    def edit_previews(self):
        self.previewsPath = filedialog.askdirectory()
        self.hasPreviews = True
        print(self.previewsPath)

        data = {'path': self.previewsPath+"/"}
        with open(self.dataPath, 'w') as json_file:
            json.dump(data, json_file)

        self.viewMSG(self.previews_frame, self.previewsPath, 15, "white")

    def main(self):
        if self.hasOrders and self.hasPreviews:
            sort_data(self, self.previewsPath, self.ordersPath)
        else:
            self.viewMSG(self.upload_frame, "Nelze generovat, výkresy nebo objednávky nejsou k dispozici ", 20, "red")

def sort_data(ui, previewsPath, ordersPath):
    try:
        search_column = ['VÝROBEK', 'POČET', 'MÍSTO']
        column_index = []
        
        df = pd.read_excel(ordersPath)
        num_row = df.shape[0]

        sort_index = df.index[df.iloc[:, 0] == "VÝKON"]
        df_head = df.iloc[sort_index[0], :].tolist()

        for x in search_column:
            column_index.append(df_head.index(x))
        print(column_index)
        for i in range(len(sort_index)-1):
            data = []
            for order in df.iloc[sort_index[i]+1:sort_index[i+1],column_index].values.tolist():
                if type(order[0]) == float: 
                    break
                else:
                    if any(isinstance(y, float) for y in order):
                        ui.viewMSG(ui.upload_frame, f"Chyba: chybí údaje v řádcích", 20, "red")
                        return   
                    imgPath = next(Path(previewsPath).glob(f'{order[0]}.*'), None) 
                    if imgPath == None:
                        ui.viewMSG(ui.upload_frame, f"Chyba: výkres výrobku {order[0]} není evidován", 20, "red")
                        return
                    x = Label(order[0], str(imgPath),order[1], order[2])
                    data.append(x)
            labels[data[0].place] = data

        data = []
        for order in df.iloc[sort_index[len(sort_index)-1]+1:num_row,column_index].values.tolist():
            if type(order[0]) == float: 
                break
            else:    
                if any(isinstance(y, float) for y in order):
                        ui.viewMSG(ui.upload_frame, f"Chyba: chybí údaje v řádcích", 20, "red")
                        return   
                imgPath = next(Path(previewsPath).glob(f'{order[0]}.*'), None) 
                if imgPath == None:
                    ui.viewMSG(ui.upload_frame, f"Chyba: výkres výrobku {order[0]} není evidován", 20, "red")
                    return
                x = Label(order[0], str(imgPath),order[1], order[2])
                data.append(x)

        labels[data[0].place] = data

        renderMain(ui)
    except Exception as e:
        ui.viewMSG(ui.upload_frame, f"Chyba: {e}", 20, "red") #"Chyba při generování

def render_to_document(tpl, context):
    tpl.render(context)
    buffer = io.BytesIO()
    tpl.save(buffer)
    buffer.seek(0)
    return Document(buffer)


def renderMain(ui: App):
    global tempDir
    tempDir = tempfile.TemporaryDirectory()
    for type, data in labels.items():
        first_page = True
        composer = None
        split_labels = [data[x:x+4] for x in range(0, len(data),4)]
        for group_label in split_labels:
            context_dic = {}
            tpl = DocxTemplate(template_path)
            for label in range(len(group_label)): 
                context_dic[f"productID{label+1}"] = group_label[label].productID
                image = InlineImage(tpl, group_label[label].previewPath, width=Mm(57), height=Mm(36))
                
                context_dic[f"previewPath{label+1}"] = image

                context_dic[f"units{label+1}"] = group_label[label].units
                context_dic[f"place{label+1}"] = group_label[label].place
            if first_page:
                main_doc = render_to_document(tpl,context_dic)
                composer = Composer(main_doc)
                first_page = False
            else:
                doc = render_to_document(tpl,context_dic)
                composer.append(doc)
        savePath = os.path.join(tempDir.name, f"{type}.docx")
        filesPath.append(savePath)
        composer.save(savePath)
    print(filesPath)
    ui.viewPrint()
    
def printLabels(ui: App, type):
    try:     
        filename = filesPath[type]
        os.startfile(filename)
        '''
        select_printer = ui.combobox.get()
        win32print.SetDefaultPrinter(select_printer)
        win32api.ShellExecute(
            0,
            "printto",
            filename,
            '"%s"' % win32print.GetDefaultPrinter(),
            ".",
            0
        )
        '''
    except Exception as e:
        ui.viewMSG(ui.upload_frame, f"Chyba: {e}", 20, "red")


if __name__ == "__main__":
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    template_path = os.path.join(base_path, "templateNew2.docx")
    icon_path = os.path.join(base_path, "icon.ico")

    change_dir_to = os.chdir(f'{Path.home()}/AppData/Roaming')
    try:
        os.mkdir("Štítky-Generátor")
    except:
        pass
    changed_dir_to = os.chdir("Štítky-Generátor")
    app = App()
    app.mainloop()