# -*- coding: utf-8 -*-

import win32print
import win32api
from pypdf import PdfReader, PdfWriter
from docx2pdf import convert



def scale(arg, size):
    pmm = (1/100*25.4)#80
    ks = min(size) / (max((arg.mediabox.width, arg.mediabox.height)) * pmm)
    arg.scale_by(ks)
    return arg

def resize(arg, size, сopies, method):

    reader = PdfReader(arg)
    writer = PdfWriter()
    if len(reader.pages) == 1 and method: #Страниц в документе 1
        for i in range(сopies):
            i = reader.pages[0]
            i  = scale(i, size)
            writer.add_page(i)
    else:
        for i in reader.pages: #Страниц в документе больше 1
            i  = scale(i, size)
            writer.add_page(i)

    if method:
        reader = PdfReader("exit_list.pdf") #Добавляем в конец документа
        exit_list = reader.pages[0]
        exit_list  = scale(exit_list, size)
        writer.add_page(exit_list)

    
    writer.write("temp.pdf")
    return "temp.pdf"


def execute(input_pdf, сopies, size, device_name, method = 1):

    temp_res = input_pdf.split('.')[-1]
    if temp_res == 'pdf':
        input_pdf = resize(input_pdf, size, сopies, method)
        list_ = [input_pdf]
        сopies = 1 if method else сopies
    elif temp_res == 'docx':
        convert(input_pdf, "temp.pdf")
        input_pdf = resize("temp.pdf", size, сopies, method)
        list_ = [input_pdf]
        сopies = 1 if method else сopies
    else:
        
        list_ = [input_pdf, "exit_list.pdf"] if method else [input_pdf]

    for input_pdf in list_:
        
        # Устанавливаем дефолтный принтер
        win32print.SetDefaultPrinterW(device_name)
        win32print.SetDefaultPrinter(device_name)
        
        preferens = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS} # тут нужные права на использование принтеров
        handle = win32print.OpenPrinter(device_name, preferens)

        #параметры принтера
        properties = win32print.GetPrinter(handle, 2)
        
        Width, Length = size
        properties['pDevMode'].PaperSize = 0 #0 если заданны ширина и высота
        properties['pDevMode'].PaperWidth = Width*10 #ширина 1 mm  *10
        properties['pDevMode'].PaperLength = Length*10 #высота 1 mm *10
        #properties['pDevMode'].Scale = 130 #маштаб
        if input_pdf == "exit_list.pdf":
            сopies = 1
        properties['pDevMode'].Copies = сopies #Количество копий

        win32print.SetPrinter(handle, 2, properties, 0) # Передаем нужные значения в принтер
        win32api.ShellExecute(0,'print', input_pdf, None ,'/manualstoprint',0) # 2 в начале для открытия pdf и его сворачивания, для открытия без сворачивания поменяйте на 1
        win32print.ClosePrinter(handle) # "Закрываем" принтер



if __name__ == '__main__':

    input_pdf = r'test.pdf'
    execute(input_pdf, 2, (56,40),device_name='Pantum PT-B680 Series', method = 1)
