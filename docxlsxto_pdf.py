import os
import tqdm
import win32com.client
from fpdf import FPDF
import pathlib
import PyPDF2

class MS():

    @staticmethod
    def word_to_pdf(word, ext, *args):
        """
        converts docx to pdf and uses the filepath from converter
        """
        file_loc_and_name = str(os.path.join(args[1], args[0][:-len(ext)])) + ".pdf" 
        o = win32com.client.Dispatch('Word.Application')
        o.Visible = False
        doc = o.Documents.Open(word)
        o.DisplayAlerts = False
        doc.SaveAs(file_loc_and_name, FileFormat=17)
        o.DisplayAlerts = True
        doc.Close()
        o.Quit()

        return

    @staticmethod
    def excel_to_pdf(excel,ext, *args):
        """
        converts xlsx to pdf and uses the filepath from converter
        """        
        file_loc_and_name = str(os.path.join(args[1], args[0][:-len(ext)])) + ".pdf"
        xlApp = win32com.client.Dispatch("Excel.Application")
        xlApp.Visible = False
        books = xlApp.Workbooks.Open(excel)
        xlApp.DisplayAlerts = False
        books.SaveAs(file_loc_and_name, FileFormat=57)
        xlApp.DisplayAlerts = True
        books.Close()
        xlApp.Quit()

        return



class Pdf():

    @staticmethod
    def to_text(text, *args):
        """
        converts pdf to txt and uses filepath from converter
        """
        content = ''
        with open(text, "rb") as f:
            #create reader variable that will read the pdf_obj
            pdfreader = PyPDF2.PdfFileReader(f)
            
            for num in range(pdfreader.numPages):
                content += pdfreader.getPage(num).extractText()

        file_loc_and_name = str(os.path.join(args[2], args[1][:-4])) + ".txt" 
        with open(file_loc_and_name, "w", encoding='utf-8') as T:
            T.writelines(content)

        return

class Convtextpdf():
    
    @staticmethod
    def txt_to_pdf(text, *args):
        """
        converts txt to pdf and uses the filepath from converter
        """
        pdf = FPDF()
        pdf.add_page()

        # set style and size of font 
        # that you want in the pdf
        pdf.set_font("Arial", size = 12)

        # open the text file in read mode

        with open(text, "r") as f:
        # insert the texts in pdf
            for x in f:
                pdf.cell(200, 10, txt = x, ln = 1, align = 'C')
        # save the pdf with name .pdf
        file_loc_and_name = str(os.path.join(args[1], args[0][:-4])) + ".pdf" 
        pdf.output(file_loc_and_name) 

        return



p = os.getcwd()
dir_len = len(next(os.walk(p))[1])
for root, dirs, files in tqdm(os.walk(p), total=dir_len):
    if len(root) != 0:
        for filename in files:
            ext = pathlib.Path(filename).suffix
            if ext == ".docx" or ext == ".doc":
                file = os.path.join(root, filename)
                MS.word_to_pdf(file, ext, filename, output)
            elif ext == ".xlsx" or ext == ".xls":
                file = os.path.join(root, filename)
                MS.excel_to_pdf(file, ext, filename, output)
            elif ext == ".txt":
                file = os.path.join(root, filename)
                Convtextpdf.txt_to_pdf(file, filename, output)
            else:
                pass
    else:
        raise KeyError("")  