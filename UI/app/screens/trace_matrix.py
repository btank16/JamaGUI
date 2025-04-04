import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from app.utils.general_format import format_general
from app.utils.border_format import add_merged_borders
import os

class TraceMatrixScreen(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # Configure grid weights to center the content
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self.create_widgets()

    def create_widgets(self):
        # Create a main frame to hold all widgets
        main_frame = ttk.Frame(self)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        # Configure main_frame grid weights to center its content
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_rowconfigure(4, weight=1)  # Add weight to last row
        main_frame.grid_columnconfigure(0, weight=1)

        # Back Button
        back_button = ttk.Button(
            main_frame, 
            text="← Back", 
            command=lambda: self.parent.master.show_frame("StartScreen")
        )
        back_button.grid(row=1, column=0, sticky="nw", pady=(0, 10))

        # Title Label
        title_label = ttk.Label(main_frame, text="Trace Matrix Format", font=("Arial", 16, "bold"))
        title_label.grid(row=2, column=0, pady=(0, 20))

        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection")
        file_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

        # File selection button and label
        self.file_button = ttk.Button(file_frame, text="Select File", command=self.select_file)
        self.file_button.grid(row=0, column=0, pady=10, padx=5)

        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # Add output filename input
        ttk.Label(file_frame, text="Output Filename:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.output_filename_var = tk.StringVar()
        self.output_filename_entry = ttk.Entry(file_frame, textvariable=self.output_filename_var, width=40)
        self.output_filename_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Add header row input
        ttk.Label(file_frame, text="Header Row:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.header_row_var = tk.StringVar(value="4")  # Default value of 4
        self.header_row_entry = ttk.Entry(file_frame, textvariable=self.header_row_var, width=5)
        self.header_row_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # Add borders dropdown
        ttk.Label(file_frame, text="Add Borders:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.borders_var = tk.StringVar(value="Yes")
        borders_dropdown = ttk.Combobox(file_frame, textvariable=self.borders_var, values=["Yes", "No"], width=5, state="readonly")
        borders_dropdown.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # Generate Button
        style = ttk.Style()
        style.configure('Large.TButton', padding=10, font=('Arial', 16, 'bold'))
        self.generate_button = ttk.Button(main_frame, text="Generate", command=self.generate, style='Large.TButton')
        self.generate_button.grid(row=4, column=0, pady=20, sticky="ew")

    def select_file(self):
        filename = filedialog.askopenfilename(title="Select File")
        if filename:
            self.file_label.config(text=filename)
            base_name = os.path.splitext(os.path.basename(filename))[0]
            self.output_filename_var.set(f"{base_name}_formatted")

    def generate(self):
        if not hasattr(self, 'file_label') or not self.file_label.cget('text') or self.file_label.cget('text') == "No file selected":
            messagebox.showerror("Error", "Please select a file first")
            return
        
        file_path = self.file_label.cget('text')
        output_filename = self.output_filename_var.get()
        header_row = self.header_row_var.get()
        
        try:
            # Validate header row is a number
            if not header_row.isdigit():
                raise ValueError("Header row must be a number")
            
            # Apply general formatting (without harm ID column)
            wb = format_general(file_path, header_row=int(header_row), harm_id_col=None)
            
            # Add thick borders if selected
            if self.borders_var.get() == "Yes":
                wb = add_merged_borders(wb)
            
            # Save the final workbook with custom filename
            output_dir = os.path.dirname(file_path)
            output_path = os.path.join(output_dir, f"{output_filename}.xlsx")
            wb.save(output_path)
            
            messagebox.showinfo("Success", f"File has been processed and saved as:\n{output_path}")
            
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")