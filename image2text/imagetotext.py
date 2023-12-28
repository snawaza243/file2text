import os
from tkinter import Tk, Label, Button, Listbox, Scrollbar, filedialog, messagebox, StringVar, Toplevel
from tkinter.ttk import Progressbar

from PIL import Image
import pytesseract
import subprocess

class ImageToTextConverter:
    def __init__(self, master):
        self.master = master
        master.title("Image to Text Converter")

        # Set background color
        master.configure(bg='#e6e6e6')

        # Hero Header
        self.hero_label = Label(master, text="Image to Text Converter", font=('Helvetica', 16, 'bold'), bg='#33cccc', fg='white', pady=10)
        self.hero_label.pack(fill='x')

        self.label_folder = Label(master, text="Image Folder:", bg='#e6e6e6')
        self.label_folder.pack()

        self.folder_path_var = StringVar()
        self.folder_path_label = Label(master, textvariable=self.folder_path_var, bg='#e6e6e6')
        self.folder_path_label.pack()

        self.select_folder_button = Button(master, text="Select Folder", command=self.select_folder, bg='#009999', fg='white')
        self.select_folder_button.pack(pady=5, padx=10, side='left')

        self.label_output = Label(master, text="Output Text File:", bg='#e6e6e6')
        self.label_output.pack()

        self.output_path_var = StringVar()
        self.output_path_label = Label(master, textvariable=self.output_path_var, bg='#e6e6e6')
        self.output_path_label.pack()

        self.select_output_button = Button(master, text="Select Output File", command=self.select_output, bg='#009999', fg='white')
        self.select_output_button.pack(pady=5, padx=10, side='left')

        self.convert_button = Button(master, text="Convert", command=self.convert_images, bg='#009999', fg='white')
        self.convert_button.pack(pady=10, side='left', padx=10)

        self.selected_images_label = Label(master, text="Selected Images:", bg='#e6e6e6')
        self.selected_images_label.pack()

        self.selected_images_listbox = Listbox(master, selectmode="multiple", height=5, width=50)
        self.selected_images_listbox.pack()

        self.converted_images_label = Label(master, text="Converted Images:", bg='#e6e6e6')
        self.converted_images_label.pack()

        self.converted_images_listbox = Listbox(master, height=5, width=50)
        self.converted_images_listbox.pack()

        self.reset_selected_button = Button(master, text="Reset Selected List", command=self.reset_selected_list, bg='#ff3333', fg='white')
        self.reset_selected_button.pack(pady=5, side='left', padx=10)

        self.reset_converted_button = Button(master, text="Reset Converted List", command=self.reset_converted_list, bg='#ff3333', fg='white')
        self.reset_converted_button.pack(pady=5, side='left', padx=10)

        self.open_text_file_button = Button(master, text="Open Extracted Text File", command=self.open_text_file, bg='#009999', fg='white')
        self.open_text_file_button.pack(pady=10, side='left', padx=10)

        self.progress_label = Label(master, text="Progress:", bg='#e6e6e6')
        self.progress_label.pack()

        self.progress_bar = Progressbar(master, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack()

        self.info_button = Button(master, text="Info", command=self.show_info, bg='#6699cc', fg='white')
        self.info_button.pack(pady=5, padx=10, side='left')

        self.about_button = Button(master, text="About", command=self.show_about, bg='#6699cc', fg='white')
        self.about_button.pack(pady=5, padx=10, side='left')

        self.close_button = Button(master, text="Close Application", command=self.master.destroy, bg='#ff3333', fg='white')
        self.close_button.pack(pady=10, side='right', padx=10)

        # Set the path to the Tesseract executable (update this with your path)
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    def select_folder(self):
        self.image_folder = filedialog.askdirectory()
        self.folder_path_var.set(f"Image Folder: {self.image_folder}")

        # Clear previous selections
        self.selected_images_listbox.delete(0, 'end')

        # Display selected images in the listbox
        for filename in os.listdir(self.image_folder):
            if filename.endswith(('.png', '.jpg', '.jpeg', '.gif')):
                self.selected_images_listbox.insert('end', filename)

    def select_output(self):
        self.output_file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        self.output_path_var.set(f"Output Text File: {self.output_file}")

    def convert_images(self):
        if hasattr(self, 'image_folder') and hasattr(self, 'output_file'):
            total_images = self.selected_images_listbox.size()
            self.progress_bar["maximum"] = total_images

            for i in range(total_images):
                filename = self.selected_images_listbox.get(i)
                image_path = os.path.join(self.image_folder, filename)
                text = image_to_text(image_path)

                # Write the extracted text to the output file
                with open(self.output_file, 'a', encoding='utf-8') as output_file:
                    output_file.write(f"Text from {filename}:\n")
                    output_file.write(text + '\n\n')

                # Update converted images listbox
                self.converted_images_listbox.insert('end', filename)

                # Update progress bar
                self.progress_bar["value"] = i + 1
                self.master.update_idletasks()

            # Display notification upon completion
            messagebox.showinfo("Conversion Complete", "Text extraction from images is complete!")

        else:
            messagebox.showwarning("Warning", "Please select both the image folder and output text file.")

    def reset_selected_list(self):
        self.selected_images_listbox.delete(0, 'end')

    def reset_converted_list(self):
        self.converted_images_listbox.delete(0, 'end')

    def open_text_file(self):
        if hasattr(self, 'output_file'):
            try:
                subprocess.Popen(["notepad.exe", self.output_file], shell=True)
            except Exception as e:
                messagebox.showerror("Error", f"Unable to open the text file.\nError: {str(e)}")
        else:
            messagebox.showwarning("Warning", "Please select the output text file first.")

    def show_info(self):
        info_text = (
            "Image to Text Converter\n\n"
            "This application allows you to convert text from multiple images using OCR.\n\n"
            "How to use:\n"
            "1. Click 'Select Folder' to choose the folder containing your images.\n"
            "2. Click 'Select Output File' to choose the output text file.\n"
            "3. Click 'Convert' to start the conversion process.\n"
            "4. The progress bar will show the status of the conversion.\n"
            "5. Once the conversion is complete, you can open the extracted text file using 'Open Extracted Text File'.\n\n"
            "Requirements:\n"
            "1. Tesseract OCR should be installed. You can download it from https://github.com/tesseract-ocr/tesseract.\n"
            "2. Ensure that the Tesseract executable path is correctly set in the code.\n"
            "3. Images should be in common formats such as PNG, JPG, JPEG, or GIF."
        )

        info_window = Toplevel(self.master)
        info_window.title("Information")
        info_label = Label(info_window, text=info_text, padx=10, pady=10)
        info_label.pack()

    def show_about(self):
        about_text = (
            "About Image to Text Converter\n\n"
            "Version: 1.0\n"
            "Developed by: IndianTechnoEra\n\n"
            "Contact:\n"
            "Email: indiantechnoera@gmail.com\n"
            "Website: https://www.indiantechnoera.in"
        )

        about_window = Toplevel(self.master)
        about_window.title("About")
        about_label = Label(about_window, text=about_text, padx=10, pady=10)
        about_label.pack()

def image_to_text(image_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img)
    return text

if __name__ == "__main__":
    root = Tk()
    app = ImageToTextConverter(root)
    root.mainloop()
