import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from app.utils.usertask_format import usertask_format
import os

class MarathonUTScreen(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        # Create a main frame to hold all widgets
        main_frame = ttk.Frame(self)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        # Configure main_frame grid weights
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_rowconfigure(4, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # Back Button
        back_button = ttk.Button(
            main_frame, 
            text="← Back", 
            command=lambda: self.parent.master.show_frame("StartScreen")
        )
        back_button.grid(row=1, column=0, sticky="nw", pady=(0, 10))

        # Title Label
        title_label = ttk.Label(main_frame, text="Marathon UT Format", font=("Arial", 16, "bold"))
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

        # Inputs section
        inputs_frame = ttk.LabelFrame(main_frame, text="Inputs")
        inputs_frame.grid(row=4, column=0, padx=10, pady=10, sticky="ew")

        # Create and layout input fields
        input_labels = [
            "Header Row",
            "Item Type Column",
            "US Name Column",
            "UT Name Column"
        ]
        
        self.input_vars = {}
        for idx, label in enumerate(input_labels):
            # Calculate row and column positions
            row = idx // 2
            col = (idx % 2) * 2
            
            ttk.Label(inputs_frame, text=label).grid(
                row=row, column=col, padx=5, pady=5, sticky="e"
            )
            
            self.input_vars[label] = tk.StringVar()
            entry = ttk.Entry(inputs_frame, textvariable=self.input_vars[label], width=15)
            entry.grid(row=row, column=col + 1, padx=(5, 15), pady=5, sticky="w")

        # Generate Button
        style = ttk.Style()
        style.configure('Large.TButton', padding=10, font=('Arial', 16, 'bold'))
        self.generate_button = ttk.Button(main_frame, text="Generate", command=self.generate, style='Large.TButton')
        self.generate_button.grid(row=5, column=0, pady=20, sticky="ew")

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
        
        try:
            # Get input values
            header_row = self.input_vars["Header Row"].get()
            item_type_col = self.input_vars["Item Type Column"].get()
            us_name_col = self.input_vars["US Name Column"].get()
            ut_name_col = self.input_vars["UT Name Column"].get()
            output_filename = self.output_filename_var.get()
            
            # Validate that all fields are filled
            if not all([header_row, item_type_col, us_name_col, ut_name_col, output_filename]):
                messagebox.showerror("Error", "Please fill in all input fields")
                return

            # Process the file
            wb = usertask_format(
                file_path,
                header_row,
                item_type_col,
                us_name_col,
                ut_name_col
            )
            
            # Save the final workbook with custom filename
            output_dir = os.path.dirname(file_path)
            output_path = os.path.join(output_dir, f"{output_filename}.xlsx")
            wb.save(output_path)
            
            messagebox.showinfo("Success", f"File has been processed and saved as:\n{output_path}")
            
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}") 