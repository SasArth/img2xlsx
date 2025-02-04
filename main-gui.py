import tkinter as tk
from tkinter import filedialog, messagebox
from img2table.ocr import PaddleOCR
from img2table.document import Image
import threading  # To run the OCR process in the background

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR Table Extraction Tool")

        self.lang_var = tk.StringVar(value="en")  # Default language
        self.input_file = None
        self.output_file = None

        # Language Selection
        tk.Label(self.root, text="Select Language:").grid(row=0, column=0, padx=10, pady=10)
        self.lang_menu = tk.OptionMenu(self.root, self.lang_var, "en", "fr", "de", "es", "it")  # Extend with more languages if needed
        self.lang_menu.grid(row=0, column=1, padx=10, pady=10)

        # Input File Selection
        tk.Label(self.root, text="Select Input File:").grid(row=1, column=0, padx=10, pady=10)
        self.input_button = tk.Button(self.root, text="Browse", command=self.select_input_file)
        self.input_button.grid(row=1, column=1, padx=10, pady=10)
        
        self.input_file_label = tk.Label(self.root, text="No file selected")
        self.input_file_label.grid(row=1, column=2, padx=10, pady=10)

        # Output File Selection
        tk.Label(self.root, text="Select Output File:").grid(row=2, column=0, padx=10, pady=10)
        self.output_button = tk.Button(self.root, text="Browse", command=self.select_output_file)
        self.output_button.grid(row=2, column=1, padx=10, pady=10)

        self.output_file_label = tk.Label(self.root, text="No file selected")
        self.output_file_label.grid(row=2, column=2, padx=10, pady=10)

        # Loading Label (hidden initially)
        self.loading_label = tk.Label(self.root, text="Loading, please wait...", fg="blue")
        self.loading_label.grid(row=3, column=0, columnspan=3, pady=20)
        self.loading_label.grid_forget()  # Hide the loading label initially

        # Start Button to Process
        self.start_button = tk.Button(self.root, text="Start OCR", command=self.start_ocr_process)
        self.start_button.grid(row=4, column=0, columnspan=3, pady=20)

    def select_input_file(self):
        self.input_file = filedialog.askopenfilename()
        if self.input_file:
            self.input_file_label.config(text=self.input_file)  # Update the label to show the selected file

    def select_output_file(self):
        self.output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if self.output_file:
            self.output_file_label.config(text=self.output_file)  # Update the label to show the selected output path

    def start_ocr_process(self):
        if not self.input_file or not self.output_file:
            messagebox.showerror("Error", "Please select both input and output files.")
            return

        # Show the loading indicator
        self.loading_label.grid(row=3, column=0, columnspan=3, pady=20)

        # Run the OCR process in a separate thread to avoid blocking the UI
        threading.Thread(target=self.run_ocr, daemon=True).start()

    def run_ocr(self):
        try:
            # Instantiation of OCR with selected language
            ocr = PaddleOCR(lang=self.lang_var.get())

            # Instantiation of document
            doc = Image(self.input_file)

            # Extraction of tables and creation of an xlsx file containing tables
            doc.to_xlsx(dest=self.output_file,
                        ocr=ocr,
                        implicit_rows=False,
                        implicit_columns=False,
                        borderless_tables=False,
                        min_confidence=50)

            # After completion, hide the loading indicator and show a success message
            self.loading_label.grid_forget()
            messagebox.showinfo("Success", f"Table extraction successful! Output saved to {self.output_file}")
        except Exception as e:
            # If an error occurs, hide the loading label and show an error message
            self.loading_label.grid_forget()
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()
