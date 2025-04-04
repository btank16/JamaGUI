import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from ..utils.duplicate_item import highlight_duplicates_in_column

class DuplicateSearchScreen(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        # Create a main frame to hold all widgets
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection")
        file_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # Add file selection button and label
        ttk.Button(file_frame, text="Select File", command=self.select_file).grid(row=0, column=0, padx=5, pady=5)
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.grid(row=0, column=1, padx=5, pady=5)

        # Add output filename input
        ttk.Label(file_frame, text="Output Filename:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.output_filename_var = tk.StringVar()
        self.output_filename_entry = ttk.Entry(file_frame, textvariable=self.output_filename_var, width=40)
        self.output_filename_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Inputs section
        inputs_frame = ttk.LabelFrame(main_frame, text="Inputs")
        inputs_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        # Create and layout input field
        ttk.Label(inputs_frame, text="Search Column:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.column_var = tk.StringVar()
        self.column_entry = ttk.Entry(inputs_frame, textvariable=self.column_var, width=15)
        self.column_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Add Generate button
        ttk.Button(main_frame, text="Generate", command=self.generate).grid(row=2, column=0, pady=20)

    def select_file(self):
        filename = filedialog.askopenfilename(title="Select File")
        if filename:
            self.file_label.config(text=filename)
            base_name = os.path.splitext(os.path.basename(filename))[0]
            self.output_filename_var.set(f"{base_name}_duplicates")

    def generate(self):
        if not hasattr(self, 'file_label') or not self.file_label.cget('text') or self.file_label.cget('text') == "No file selected":
            messagebox.showerror("Error", "Please select a file first")
            return
        
        file_path = self.file_label.cget('text')
        
        try:
            # Get input value
            search_column = self.column_var.get()
            
            # Validate that field is filled
            if not search_column:
                messagebox.showerror("Error", "Please enter a search column")
                return
            
            # Create output path
            output_filename = self.output_filename_var.get()
            output_dir = os.path.dirname(file_path)
            output_path = os.path.join(output_dir, f"{output_filename}.xlsx")
            
            # Copy file to new location
            highlight_duplicates_in_column(file_path, search_column)
            
            messagebox.showinfo("Success", f"File has been processed and saved as:\n{output_path}")
            
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}") 