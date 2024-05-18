from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox
from tkinter.ttk import Progressbar
from pdf2docx import parse
from docx2pdf import convert as docx2pdf_convert
import pathlib

def callback():
    name = fd.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("DOCX files", "*.docx")])
    ePath.config(state='normal')
    ePath.delete('1', END)
    ePath.insert('1', name)
    ePath.config(state='readonly')

def select_save_folder():
    folder = fd.askdirectory()
    if folder:
        eFolderPath.config(state='normal')
        eFolderPath.delete('1', END)
        eFolderPath.insert('1', folder)
        eFolderPath.config(state='readonly')

def validate_file(file_path, valid_extensions):
    if not pathlib.Path(file_path).suffix.lower() in valid_extensions:
        messagebox.showerror("Ошибка", f"Неверный формат файла. Допустимые форматы: {', '.join(valid_extensions)}")
        return False
    return True

def convert_pdf_to_docx():
    pdf_file = ePath.get()
    save_folder = eFolderPath.get()
    if not pdf_file or not save_folder:
        status_label.config(text='Выберите PDF файл и папку для сохранения', fg='red')
        return
    if not validate_file(pdf_file, ['.pdf']):
        return

    word_file = pathlib.Path(save_folder) / (pathlib.Path(pdf_file).stem + '.docx')
    progress_bar.start()
    parse(pdf_file, str(word_file))
    progress_bar.stop()
    status_label.config(text='Конвертация PDF в DOCX завершена', fg='lime')

def convert_docx_to_pdf():
    docx_file = ePath.get()
    save_folder = eFolderPath.get()
    if not docx_file or not save_folder:
        status_label.config(text='Выберите DOCX файл и папку для сохранения', fg='red')
        return
    if not validate_file(docx_file, ['.docx']):
        return

    pdf_file = pathlib.Path(save_folder) / (pathlib.Path(docx_file).stem + '.pdf')
    progress_bar.start()
    docx2pdf_convert(docx_file, str(pdf_file))
    progress_bar.stop()
    status_label.config(text='Конвертация DOCX в PDF завершена', fg='lime')

root = Tk()
root.title('Конвертер PDF в Word и DOCX в PDF')
root.geometry('600x500+300+300')
root.resizable(width=False, height=False)
root['bg'] = 'black'

Button(root, text='Выбрать файл', font='Arial 15 bold', fg='lime', bg='black', command=callback).pack(pady=10)
lbPath = Label(root, text='Путь к файлу:', fg='lime', bg='black', font='Arial 15 bold')
lbPath.pack()
ePath = Entry(root, width=60, state='readonly')
ePath.pack(pady=10)

Button(root, text='Выбрать папку для сохранения', font='Arial 15 bold', fg='lime', bg='black', command=select_save_folder).pack(pady=10)
lbFolderPath = Label(root, text='Путь к папке:', fg='lime', bg='black', font='Arial 15 bold')
lbFolderPath.pack()
eFolderPath = Entry(root, width=60, state='readonly')
eFolderPath.pack(pady=10)

Button(root, text='Конвертировать PDF в DOCX', fg='lime', bg='black', font='Arial 15 bold', command=convert_pdf_to_docx).pack(pady=10)
Button(root, text='Конвертировать DOCX в PDF', fg='lime', bg='black', font='Arial 15 bold', command=convert_docx_to_pdf).pack(pady=10)

status_label = Label(root, text='', fg='lime', bg='black', font='Arial 15 bold')
status_label.pack(pady=10)
progress_bar = Progressbar(root, orient=HORIZONTAL, length=300, mode='indeterminate')
progress_bar.pack(pady=10)

root.mainloop()
