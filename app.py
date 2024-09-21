import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import arrow
#selecionar o arquivo da planilha
def selectFile():
    filePath = filedialog.askopenfilenames(filetypes=[("Excel Files", ".xlsx;.xlsm")])
    if filePath:
        entryPlan.config(state="normal")
        entryPlan.delete(0, tk.END)
        entryPlan.insert(0, filePath)
        entryPlan.config(state="readonly")

#Selecionar a Pasta Para Salvar os Certificados
def selectFolder():
    folderPath = filedialog.askdirectory()
    if folderPath:
        entryFolder.config(state="normal")
        entryFolder.delete(0, tk.END)
        entryFolder.insert(0, folderPath)
        entryFolder.config(state="readonly")

#logica dos certificados
def createCert():
    

    plan = entryPlan.get()

    if plan:
        workbookAlunos = openpyxl.load_workbook(plan)
        sheetAlunos = workbookAlunos["sheet1"]
        messagebox.showinfo("Acessando", "Abrindo sua planilha.")

        for linha in sheetAlunos.iter_rows(min_row=2):
            #pegar as informações da Planilha
            name = linha[0].value
            dataInit = linha[1].value
            dataEnd = linha[2].value
            dataEmis = linha[3].value

            

            dataInit_str = arrow.get(dataInit).format("DD/MM/YYYY")
            dataEnd_str = arrow.get(dataEnd).format("DD/MM/YYYY")
            dataEmis_str = arrow.get(dataEmis).format("DD/MM/YYYY")


            print(name, dataInit_str, dataEnd_str, dataEmis_str)

            #Abrindo e Criando Certificados
            img = Image.open("./certificado.png")
            draw = ImageDraw.Draw(img)

            #fonts
            fontName = ImageFont.truetype("./fonts/BRUSHSCI.TTF", 110)
            fontGeral = ImageFont.truetype("./fonts/ERASBD.TTF", 30)
            #COlORS
            blue = "#002D69"

            #Name
            draw.text((550, 500), name, font=fontName, fill=blue, align="center")

            #Data de inicio
            draw.text((472, 758), dataInit_str, font=fontGeral, fill=blue, align="center")

            #Data de Termino
            draw.text((710, 758), dataEnd_str, font=fontGeral, fill=blue, align="center")

            #Data de Emissão do Certif.
            draw.text((1440, 823), dataEmis_str, font=fontGeral, fill=blue, align="center")

            pathFolder = entryFolder.get()
            img.save(f"./{pathFolder}/{name} Certificado.png")
        
        messagebox.showinfo("Concluido", "Todos os certificados foram criados!")
    else:
        messagebox.showerror("Erro ao Acessar a Planilha", "Não foi possivel carregar a planilha. Tente Novamente!")



            


#tela
root = tk.Tk()
root.geometry("450x100")
root.title("Automatizador de Certificado")

#cofig Layout
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=4)


#File Excel
btnPlan = ttk.Button(root, text="Selecionar Arquivo", command=selectFile)
btnPlan.grid(row=0, column=0, sticky="ew", padx=20)
entryPlan = ttk.Entry(root,)
entryPlan.grid(row=0,column=1, sticky="ew", padx=20)

#Select Folder
btnFolder = ttk.Button(root, text="Selecionar Pasta", command=selectFolder)
btnFolder.grid(row=1, column=0, sticky="ew", padx=20)
entryFolder = ttk.Entry(root, )
entryFolder.grid(row=1, column=1, sticky="ew", padx=20)

#button
frame = ttk.Frame(root,)
frame.grid(row=2, column=0, columnspan=2)
btnEnd = ttk.Button(frame, text="Criar Certificados", command=createCert)
btnEnd.grid(row=0, column=0, sticky="nsew", padx=30)

root.mainloop()