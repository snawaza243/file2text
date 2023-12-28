import os
from tkinter import Tk, Label, Button, Listbox, filedialog, messagebox, StringVar, Toplevel
from tkinter.ttk import Progressbar
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import io
import shutil

def select_folder():
    pdf_folder = filedialog.askdirectory()
    folder_path_var.set(f"PDF Folder: {pdf_folder}")
    update_pdf_list(pdf_folder)

def update_pdf_list(pdf_folder):
    pdf_list.delete(0, 'end')  # Clear the listbox
    if pdf_folder:
        pdf_files = [file for file in os.listdir(pdf_folder) if file.endswith('.pdf')]
        for pdf_file in pdf_files:
            pdf_list.insert('end', pdf_file)

def open_selected_folder():
    if os.path.exists(folder_path_var.get().replace("PDF Folder: ", "")):
        os.startfile(folder_path_var.get().replace("PDF Folder: ", ""))

def select_output():
    output_file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    output_path_var.set(f"Output Text File: {output_file}")

def open_converted_folder():
    if os.path.exists(output_path_var.get().replace("Output Text File: ", "")):
        output_folder = os.path.dirname(output_path_var.get().replace("Output Text File: ", ""))
        os.startfile(output_folder)

def convert_pdfs():
    pdf_folder = folder_path_var.get().replace("PDF Folder: ", "")
    output_file = output_path_var.get().replace("Output Text File: ", "")
    
    if pdf_folder and output_file:
        pdf_files = [file for file in os.listdir(pdf_folder) if file.endswith('.pdf')]
        total_pdfs = len(pdf_files)
        progress_bar["maximum"] = total_pdfs

        progress_list.delete(0, 'end')  # Clear the progress listbox

        for i, pdf_filename in enumerate(pdf_files):
            pdf_path = os.path.join(pdf_folder, pdf_filename)
            text_from_pdf = extract_text(pdf_path, output_file)

            # Update progress bar
            progress_bar["value"] = i + 1
            root.update_idletasks()

            # Update progress list
            progress_list.insert('end', f"{pdf_filename} - Done")

        # Display notification upon completion
        messagebox.showinfo("Conversion Complete", "Text extraction from PDFs is complete!")

    else:
        messagebox.showwarning("Warning", "Please select both the PDF folder and output text file.")

def extract_text(pdf_path, output_file_path):
    # Extract text from a single PDF using PyMuPDF
    text = ""
    with fitz.open(pdf_path) as pdf_document:
        for page_number in range(pdf_document.page_count):
            page = pdf_document[page_number]
            text += page.get_text("text")

        # Extract images from the PDF using PyMuPDF and perform OCR
        with open(output_file_path, 'a', encoding='utf-8') as output_file:
            for page_number in range(pdf_document.page_count):
                page = pdf_document[page_number]
                image_list = page.get_images(full=True)

                for img_index, img in enumerate(image_list):
                    image_index = img[0]
                    base_image = pdf_document.extract_image(image_index)
                    image_bytes = base_image["image"]
                    image = Image.open(io.BytesIO(image_bytes))

                    # Perform OCR on the image
                    text_from_image = image_to_text(image)

                    # Write the extracted text from OCR to the output file
                    image_filename = f"{os.path.splitext(pdf_path)[0]}_page{page_number + 1}_img{img_index}.txt"
                    output_file.write(f"Text from OCR on image {image_filename}:\n")
                    output_file.write(text_from_image + '\n\n')

    return text

def image_to_text(image):
    # Perform OCR on the image
    text = pytesseract.image_to_string(image)
    return text

def reset():
    folder_path_var.set("PDF Folder:")
    output_path_var.set("Output Text File:")
    progress_bar["value"] = 0
    pdf_list.delete(0, 'end')
    progress_list.delete(0, 'end')

def show_info():
    info_text = (
        "PDF to Text Converter\n\n"
        "This application allows you to convert text from multiple PDFs using PyMuPDF and OCR.\n\n"
        "How to use:\n"
        "1. Click 'Select Folder' to choose the folder containing your PDFs.\n"
        "2. Click 'Select Output File' to choose the output text file.\n"
        "3. Click 'Convert' to start the conversion process.\n"
        "4. The progress bar will show the status of the conversion.\n"
        "5. Once the conversion is complete, you can view the extracted text file.\n\n"
        "Requirements:\n"
        "1. Tesseract OCR should be installed. You can download it from https://github.com/tesseract-ocr/tesseract.\n"
        "2. Ensure that the Tesseract executable path is correctly set in the code.\n"
    )

    info_window = Toplevel(root)
    info_window.title("Information")
    info_label = Label(info_window, text=info_text, padx=10, pady=10)
    info_label.pack()

def show_about():
    about_text = (
        "About PDF to Text Converter\n\n"
        "Version: 1.0\n"
        "Developed by: IndianTechnoEra\n\n"
        "Contact:\n"
        "Email: indiantechnoera@gmail.com\n"
        "Website: https://www.indiantechnoera.in"
    )

    about_window = Toplevel(root)
    about_window.title("About")
    about_label = Label(about_window, text=about_text, padx=10, pady=10)
    about_label.pack()


root = Tk()
root.title("PDF to Text Converter")

# Set the main window's height and width
window_width = 410
window_height = 550

# Set main window's geometry
root.geometry(f"{window_width}x{window_height}")

# Disable resizing
root.resizable(width=False, height=False)

# Set the minimum and maximum size to the fixed dimensions
root.minsize(width=window_width, height=window_height)
root.maxsize(width=window_width, height=window_height)

# Set background color
root.configure(bg='#e6e6e6')




# Hero Header
hero_label = Label(root, text="PDF to Text Converter", font=('Helvetica', 16, 'bold'), bg='#33cccc', fg='white', pady=10)
hero_label.grid(row=0, column=0, columnspan=4, sticky='ew')

label_first_line_empty = Label(root, text=" ", bg='#e6e6e6')
label_first_line_empty.grid(row=1, column=0, sticky='w', padx=(10, 0), pady=(5, 0))



select_folder_button = Button(root, text="Select Folder", command=select_folder, bg='#009999', fg='white' , width=10)
select_folder_button.grid(row=2, column=0, sticky='w', padx=(10, 0), pady=(5, 0))

open_selected_folder_button = Button(root, text="Input Folder", command=open_selected_folder, bg='#009999', fg='white')
open_selected_folder_button.grid(row=2, column=1, sticky='w', pady=(5, 0))


label_input = Label(root, text=" Input:", bg='#e6e6e6')
label_input.grid(row=3, column=0, sticky='w', padx=(10, 0), pady=(5, 0))


folder_path_var = StringVar()
folder_path_label = Label(root, textvariable=folder_path_var,  bg='#e6e699', width=55)
folder_path_label.grid(row=3, column=0, columnspan=4, sticky='w', padx=(10, 0), pady=(5, 0))


pdf_list = Listbox(root, selectmode='extended', height=5)
pdf_list.grid(row=4, column=0, columnspan=4, sticky='ew', padx=(10, 0), pady=(5, 0))






select_output_button = Button(root, text="Set File", command=select_output, bg='#009999', fg='white', width=10)
select_output_button.grid(row=5, column=0, sticky='w', padx=(10, 0), pady=(5, 0))

open_converted_folder_button = Button(root, text="Output Folder", command=open_converted_folder, bg='#009999', fg='white', width=10)
open_converted_folder_button.grid(row=5, column=1, sticky='w', pady=(5, 0))

label_output = Label(root, text=" Output: ", bg='#e6e6e6')
label_output.grid(row=6, column=0, sticky='w', padx=(10, 0), pady=(5, 0))



output_path_var = StringVar()
output_path_label = Label(root, textvariable=output_path_var, bg='#e6e699', width=55)
output_path_label.grid(row=6, column=0, columnspan=4, sticky='w', padx=(10, 0), pady=(5, 0))






convert_button = Button(root, text="Convert", command=convert_pdfs, bg='#002295', fg='white', width=10)
convert_button.grid(row=1+6, column=0, sticky='w', padx=(10, 0), pady=(10, 0))

reset_button = Button(root, text="Reset", command=reset, bg='#ff3333', fg='white', width=10)
reset_button.grid(row=1+6, column=1, sticky='w', pady=(10, 0))

progress_label = Label(root, text="Progress:", bg='#e6e6e6')
progress_label.grid(row=1+7, column=0, sticky='w', padx=(10, 0), pady=(10, 0))

progress_bar = Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.grid(row=1+7, column=1, columnspan=3, sticky='w', pady=(10, 0))

progress_list = Listbox(root, selectmode='extended', height=5)
progress_list.grid(row=1+8, column=0, columnspan=4, sticky='ew', padx=(10, 0), pady=(5, 0))





pdf_label = Label(root, text="", bg='#e6e6e6')
pdf_label.grid(row=1+9, column=0, sticky='w', padx=(10, 0), pady=(10, 0))


info_button = Button(root, text="Info", command=show_info, bg='#6699cc', fg='white', width=10)
info_button.grid(row=1+10, column=0, sticky='w', padx=(10, 0), pady=(5, 0))

about_button = Button(root, text="About", command=show_about, bg='#6699cc', fg='white', width=10)
about_button.grid(row=12, column=0, sticky='w', padx=(10, 0), pady=(5, 0))

close_button = Button(root, text="Close Application", command=root.destroy, bg='#ff3333', fg='white')
close_button.grid(row=12, column=2, columnspan=2, sticky='e', padx=(0, 10), pady=(10, 0))

# Set the path to the Tesseract executable (update this with your path)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

root.mainloop()
