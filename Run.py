from datetime import datetime, timedelta
from tkinter import Label, filedialog, messagebox, scrolledtext, simpledialog, ttk
import tkinter as tk
from tkinter import *
import os
import sys 
import PyPDF2
import fitz
from ttkthemes import ThemedTk
def change_theme():
    if window.tk.call("ttk::style", "theme", "use") == "azure-dark":
        # Set light theme
        window.tk.call("set_theme", "light")
    else:
        # Set dark theme
        window.tk.call("set_theme", "dark")

def return_to_options():
    clear_win()
    options()

def bring_to_front(widget):
    widget.lift()
def read_meta():
    pdf_path=pdfpath
    clear_win()
    
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        info = pdf_reader.metadata
        Label(window, text=f"Title: {info.title}").pack()
        Label(window, text=f"Author: {info.author}").pack()
        Label(window, text=f"Subject: {info.subject}").pack()
        Label(window, text=f"Producer: {info.producer}").pack()
        Label(window, text=f"Creation Date: {info.creation_date}").pack()
        Label(window, text=f"Number of Pages: {len(pdf_reader.pages)}").pack()
        choice = simpledialog.askstring("Input", "Do you want to write the metadata? (yes or no)")

        if choice == "yes":
            write_meta(pdf_path)
        elif choice == "no":
            clear_win()
            return_to_options()
            Label(window,text="Refer the metadata here:",font=("helvetica")).pack(padx=10,pady=10)
            text=scrolledtext.ScrolledText(window, wrap=tk.WORD, width=60, height=20)
            text.pack(padx=10,pady=10)          
            text.insert(tk.END, f"Title: {info.title}\n")
            text.insert(tk.END, f"Author: {info.author}\n")
            text.insert(tk.END, f"Subject: {info.subject}\n")
            text.insert(tk.END, f"Producer: {info.producer}\n")
            text.insert(tk.END, f"Creation Date: {info.creation_date}\n")
            text.insert(tk.END, f"Number of Pages: {len(pdf_reader.pages)}\n")
            

def write_meta(pdf_path):
    window.iconify()
    
    pdf_reader = PyPDF2.PdfReader(pdf_path)
    pdf_writer = PyPDF2.PdfWriter()

    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        pdf_writer.add_page(page)
        
    title = simpledialog.askstring("Input", "Enter Title:")
    author = simpledialog.askstring("Input", "Enter Author:")
    subject = simpledialog.askstring("Input", "Enter Subject:")
    producer = simpledialog.askstring("Input", "Enter Producer:")

    metadata = {}

    if title:
        metadata['/Title'] = title.encode('utf-8')
    if author:
        metadata['/Author'] = author.encode('utf-8')
    if subject:
        metadata['/Subject'] = subject.encode('utf-8')
    if producer:
        metadata['/Producer'] = producer.encode('utf-8')

    utc_now = datetime.utcnow()
    offset = timedelta(hours=5, minutes=30)
    offset_str = '{:+03d}{:02d}'.format(offset.seconds // 3600, (offset.seconds // 60) % 60)
    creation_date = (utc_now + offset).strftime('D:%Y%m%d%H%M%S') + offset_str
    metadata['/CreationDate'] = creation_date.encode('utf-8')

    pdf_writer.add_metadata(metadata)
    output_file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])  
    
    with open(output_file_path, 'wb') as output_file:
        pdf_writer.write(output_file)

def extract_text():
    pdf_path = pdfpath
    clear_win()
    
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            result_text = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=60, height=20)
            result_text.pack(padx=10, pady=10)
            result_text.insert(tk.END, f"Page {page_num + 1}:\n{text}\n")
            tk.Button(window, text="Go to options", command=return_to_options).pack()

def extract_image():
    pdf_path = pdfpath
    pdf_document = fitz.open(pdf_path)
    
    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        image_list = page.get_images(full=True)

        for img_index, img_info in enumerate(image_list):
            image_index = img_info[0]
            base_image = pdf_document.extract_image(image_index)
            image_bytes = base_image["image"]
            image_filename = f"image_page{page_num + 1}_img{img_index + 1}.png"
            output_folder = filedialog.askdirectory(title="Select output folder for images")
            image_path = os.path.join(output_folder, image_filename)

            with open(image_path, "wb") as fp:
                fp.write(image_bytes)
                result_text = tk.Text(window, wrap=tk.WORD, width=70, height=3)
                result_text.pack(padx=10, pady=10)
                result_text.insert(tk.END, f"Image saved: {image_path}\n")
                
    pdf_document.close()
    clear_win()
    return_to_options()

def en_dec():
    clear_win()
    
    tk.Button(window, text="Encrypt your Pdf", command=encrypt).pack()
    tk.Button(window, text="Decrypt your Pdf", command=decrypt).pack()

def encrypt():
    clear_win()
    
    pdf_path = pdfpath
    password = simpledialog.askstring("Input", "Enter a password")

    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        pdf_writer = PyPDF2.PdfWriter()

        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            pdf_writer.add_page(page)

        pdf_writer.encrypt(password)
        encrypted_output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

    if encrypted_output_path:
        with open(encrypted_output_path, 'wb') as encrypted_file:
            pdf_writer.write(encrypted_file)

        result_text = tk.Text(window, wrap=tk.WORD, width=70, height=3)
        result_text.pack(padx=10, pady=10)
        result_text.insert(tk.END, f"PDF saved: {encrypted_output_path}\n")

    return_to_options()

def decrypt():
    clear_win()
    
    pdf_path = pdfpath
    window.iconify()
    password = simpledialog.askstring("Input", "Enter your password")

    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        pdf_writer = PyPDF2.PdfWriter()

        if pdf_reader.is_encrypted:
            if pdf_reader.decrypt(password):
                output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
                
                with open(output_path, 'wb') as output_file:
                    for page_num in range(len(pdf_reader.pages)):
                        pdf_writer.add_page(pdf_reader.pages[page_num])
                    pdf_writer.write(output_file)

                clear_win()
                return_to_options()
                result_text = tk.Text(window, wrap=tk.WORD, width=70, height=3)
                result_text.pack(padx=10, pady=10)
                result_text.insert(tk.END, f"PDF saved: {output_path}\n")
            else:
                Label(window, text="Incorrect password. Unable to decrypt PDF.").pack()
                return_to_options()
        else:
            Label(window, text="PDF is not encrypted. No decryption needed.").pack()
            return_to_options()

def merge():
    clear_win()
    message_label = tk.Label(window, text="Please choose a PDF file", font=("helvetica", 16))
    message_label.pack()
    path2 = ""
    path2 = export_path(path2)
    pdf_paths = [pdfpath, path2]
    pdf_writer = PyPDF2.PdfWriter()

    for pdf_path in pdf_paths:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                pdf_writer.add_page(page)

    merged_output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

    with open(merged_output_path, 'wb') as merged_file:
        pdf_writer.write(merged_file)

    clear_win()
    return_to_options()
    result_text = tk.Text(window, wrap=tk.WORD, width=70, height=3)
    result_text.pack(padx=10, pady=10)
    result_text.insert(tk.END, f"PDFs merged: {merged_output_path}\n")

def watermark():
    clear_win()
    
    watermark_pdf = ""
    output_pdf = ""
    input_pdf = pdfpath
    watermark_pdf = export_path(watermark_pdf)

    with open(input_pdf, 'rb') as input_file:
        pdf_reader = PyPDF2.PdfReader(input_file)
        total_pages = len(pdf_reader.pages)

        with open(watermark_pdf, 'rb') as watermark_file:
            watermark_reader = PyPDF2.PdfReader(watermark_file)
            pdf_writer = PyPDF2.PdfWriter()

            for page_num in range(total_pages):
                page = pdf_reader.pages[page_num]
                watermark_page = watermark_reader.pages[0]  
                page.merge_page(watermark_page)
                pdf_writer.add_page(page)

            output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Pdf files", "*.pdf")])

            with open(output_pdf, 'wb') as output_file:
                pdf_writer.write(output_file)

            clear_win()
            return_to_options()
            result_text = tk.Text(window, wrap=tk.WORD, width=70, height=3)
            result_text.pack(padx=10, pady=10)
            result_text.insert(tk.END, f"PDF merged: {output_pdf}\n")

def options():
    
    style = ttk.Style()
    style.configure('TButton', foreground='black', background='#a1dbcd')

    ttk.Separator(window, orient='horizontal').pack(fill='x')

    btn_metadata = ttk.Button(window, text='Read Metadata', command=read_meta)
    btn_metadata.pack(pady=5)

    btn_extract_text = ttk.Button(window, text='Extract Text', command=extract_text)
    btn_extract_text.pack(pady=5)

    btn_extract_image = ttk.Button(window, text='Extract Image', command=extract_image)
    btn_extract_image.pack(pady=5)

    btn_en_dec = ttk.Button(window, text='Encrypt and Decrypt', command=en_dec)
    btn_en_dec.pack(pady=5)

    btn_merge = ttk.Button(window, text='Merge PDFs', command=merge)
    btn_merge.pack(pady=5)

    btn_watermark = ttk.Button(window, text='Sign/Watermark PDF', command=watermark)
    btn_watermark.pack(pady=5)

    btn_exit = ttk.Button(window, text='Exit', command=sys.exit)
    btn_exit.pack(pady=5)

def export_path(modified_path):
    modified_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    
    if modified_path:
        lab = tk.Label(window, text="File fetched successfully!", font=("helvetica", 16))
        lab.pack()

    return modified_path

def clear_win():
    for widget in window.winfo_children():
        widget.destroy()

def browse():
    global pdfpath
    pdfpath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    
    if pdfpath:
        clear_win()
        lab = tk.Label(window, text="File fetched successfully!", font=("helvetica", 16),padx=10,pady=10)
        lab.pack()
        options()
def toggle_dark_mode():
   window.set_theme('plastik')
window =ThemedTk()
big_frame = ttk.Frame(window)
big_frame.pack()
window.style=ttk.Style(window)
toggle_dark_mode()
window.geometry('600x400')
window.resizable(False,False)
window.title("PDF Editor")
photo=PhotoImage(file ='icon.png')
window.iconphoto(False, photo)
window.call("source","C:\\pyprojects\\Azure-ttk-theme\\azure.tcl")
window.call("set_theme","light")
Label(window, text='Welcome to PDF File Editor', font=("Helvetica", 24)).pack(padx=20,pady=20)
button = ttk.Button(window, text='Select PDF', command=browse)
button.pack()
button = ttk.Button(big_frame, text="Change theme", command=change_theme)
button.pack()
window.mainloop()

