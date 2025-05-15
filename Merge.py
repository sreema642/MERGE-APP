import tkinter as tk
from tkinter import ttk
from fpdf import FPDF
import PyPDF2
import comtypes.client
import win32com.client
from datetime import datetime
import os
import urllib.parse
from urllib.parse import unquote

from tkinter import filedialog
import pandas as pd
from tkinter import messagebox
import sys
import tempfile

column = []
row=0
frame=None
scrollbar = None

def get_bundle_dir():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS  
    else:
        return os.path.dirname(os.path.abspath(__file__))


def table(event=None):
    global row, column, frame
    col = col_entry.get().split(",")
    row = int(row_entry.get())  
    column = ["S.No"] + [val.strip() for val in col]

    if frame:
        frame.destroy()
    frame = tk.Frame(root)
    frame.pack(padx=5,pady=5)
    for col_idx, col_name in enumerate(column):
        label = tk.Label(frame, text=col_name, relief="solid", width=10)
        label.grid(row=0, column=col_idx, padx=5, pady=5)

    for row_idx in range(1, row + 1):
        for col_idx in range(len(column)):
             if col_idx == 0:  
                sno_label = tk.Label(frame, text=str(row_idx), relief="solid", width=5, anchor="center")
                sno_label.grid(row=row_idx, column=col_idx, padx=5, pady=5)
             else:
                text_widget = tk.Text(frame, width=20, height=1)
                text_widget.grid(row=row_idx, column=col_idx, padx=3, pady=5)


        
def save():
    global row, column, frame

    pdf = FPDF(orientation='P', unit='mm', format=(300, 210))
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    total_table_width = 250
    col_widths = [total_table_width / len(column)] * len(column)

    for col_idx in range(len(column)):
        max_width = 0
        for row_idx in range(1, row + 1):
            if col_idx == 0:
                value = str(row_idx)  
            else:
                widget = frame.grid_slaves(row=row_idx, column=col_idx)
                if widget:
                    text_widget = widget[0]
                    value = text_widget.get("1.0", tk.END).strip()
                else:
                    value = ""
            text_width = pdf.get_string_width(value) + 10
            max_width = max(max_width, text_width)
        col_widths[col_idx] = max_width

    scaling_factor = total_table_width / sum(col_widths)
    col_widths = [width * scaling_factor for width in col_widths]

    # Add table headers
    for idx, col_name in enumerate(column):
        pdf.cell(col_widths[idx], 10, col_name, border=1, align="C")
    pdf.ln()

    # Add table rows
    for row_idx in range(1, row + 1):
        for col_idx in range(len(column)):
            if col_idx == 0:
                value = str(row_idx)  
            else:
                widget = frame.grid_slaves(row=row_idx, column=col_idx)
                if widget:
                    text_widget = widget[0]
                    value = text_widget.get("1.0", tk.END).strip()
                else:
                    value = ""
            pdf.cell(col_widths[col_idx], 10, value, border=1, align="L")
        pdf.ln()

    bundle_dir = get_bundle_dir()
    pdf_file = os.path.join(bundle_dir, "tabledata.pdf")
    pdf.output(pdf_file, "F")
    messagebox.showinfo("Created", f"Table data saved to {pdf_file}")

 
def merge(pdf_path1,pdf_path2,output_path):
    
            with open (pdf_path1,"rb") as file1 , open (pdf_path2,"rb") as file2:
                                      read1 = PyPDF2.PdfReader(file1)
                                      read2 = PyPDF2.PdfReader(file2)
                                      num1= len(read1.pages)
                                      num2 = len(read2.pages)

                                      writer = PyPDF2.PdfWriter()
                                      for i in range (num1):
                                          writer.add_page(read1.pages[i])
                                      for j in range (num2):
                                          writer.add_page(read2.pages[j])
                                      with open(output_path,"wb") as file3:
                                          writer.write(file3)
            messagebox.showinfo("merged", "sucessfully merged")
        
    
   
def upload ():
    
    file = filedialog.askopenfilename(title = "upload file",filetypes = (("all files","*.*"),))
    
    if not file:
        return  

    
    file = unquote(file)
    file=os.path.normpath(file)

    bundle_dir = get_bundle_dir()
    
    if file and file.lower().endswith(('.doc', '.docx')):
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(file)

    
        bundle_dir = get_bundle_dir()
        pdf_path = os.path.join(bundle_dir, "wordfile.pdf")
        temp_pdf_path = os.path.join(bundle_dir, "temp_wordfile.pdf")

        doc.SaveAs2(temp_pdf_path, FileFormat=17)  
        doc.Close()
        word.Quit()

    
        merge(os.path.join(bundle_dir, "tabledata.pdf"), temp_pdf_path, pdf_path)

    
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)


    elif   file and file.lower().endswith(('.xls', '.xlsx')):
        
        excel_path = os.path.normpath(file)
        df = pd.read_excel(excel_path)
        df.to_csv(os.path.join(bundle_dir,"excel_text.txt"), index=False, sep=",")
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial",size=24)
        with open(os.path.join(bundle_dir,"excel_text.txt"),"r") as text_file :
            for line in text_file:
                pdf.multi_cell(0,10, txt=line)
        pdf_path = os.path.join(bundle_dir,"excelfile.pdf")
        pdf.output(pdf_path,"F")
        merge(os.path.join(bundle_dir,"tabledata.pdf"),pdf_path,pdf_path)
        


                                           
    
    else:
        
         merge_path = os.path.join(bundle_dir, "pdffile.pdf")
         merge(os.path.join(bundle_dir ,"tabledata.pdf"),file,merge_path)



def combine():
    save()
    upload()
                                          


root = tk.Tk()
root.geometry("600x600")
col_lable = tk.Label(root,text="Enter the column names(comma separated values) : ")
col_lable.pack(padx=5,pady=10)

col_entry = tk.Entry(root)
col_entry.pack(padx=5,pady=5)

row_label =  tk.Label(root,text="Enter the total number of rows :  ")
row_label.pack(padx=5,pady=5)
row_entry = tk.Entry(root)
row_entry.pack(padx=5,pady=5)

col_entry.bind("<Return>", table)  
row_entry.bind("<Return>", table)

upload_button = tk.Button(root,text="upload and merge",bg="blue",fg="white",command=combine)
upload_button.place(x=900,y=50)




root.mainloop()




