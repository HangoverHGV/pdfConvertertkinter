import tkinter as tk
from tkinter import ttk, filedialog
from tkinter.filedialog import askopenfile
import os
from pdf2docx import parse
from typing import Tuple



root = tk.Tk()


color = "#c2c2c2"
red = (255, 0, 0)
font1 = ('times', 18)
font2 = ('times', 18, 'bold')

root.title('PDF Converter')
root.configure(background=color)
root.geometry('850x500')
root.resizable(width=False, height=False)

l1 = tk.Label(root, text='Upload File', font=font2)
l1.place(x=100, y=10)


btn = tk.Button(root, text='Browse...', width=20,command=lambda: upload_file(), font=font2)
btn.place(x = 20, y = 50)

btnPath = tk.Button(root, text='Select path', width=20,command=lambda: setpath(), font=font2)
btnPath.place(x = 20, y = 200)

docPath = tk.StringVar()

l3 = tk.Label(root, textvariable=docPath, fg='#FF0000', background=color)
l3.place(x = 20, y = 250)

mystr = tk.StringVar()

l2 = tk.Label(root, textvariable=mystr, fg='#FF0000', background=color)
l2.place(x = 20, y = 100)

btnconv = tk.Button(root, text='Convert', command=lambda: convert(), font=font1)
btnconv.place(x = 400, y = 200)


warn_text = tk.StringVar()
label_warn = tk.Label(root, textvariable=warn_text, background=color,font=font2, fg='red')
label_warn.place(x = 250, y = 400)

warn_text.set('')

mystr.set('')
def upload_file():
    global pathName, fileName
    file = filedialog.askopenfilename(filetypes=[('PDF Files', '.pdf')])
    print(file)
    # fileName, fileextension = os.path.splitext(file.filename)
    pathName = os.path.dirname(file)
    mystr.set(file)


def convert():
    convert_pdf2docx(mystr.get(), docPath.get())

def setpath():
    folder = filedialog.askdirectory()
    docPath.set(folder)

def convert_pdf2docx(input_file: str, output_file: str, pages: Tuple = None):
    """Converts pdf to docx"""
    if pages:
        pages = [int(i) for i in list(pages) if i.isnumeric()]
    result = parse(pdf_file=input_file,
                   docx_with_path=output_file, pages=pages)
    summary = {
        "File": input_file, "Pages": str(pages), "Output File": output_file
    }
    return result

root.mainloop()
