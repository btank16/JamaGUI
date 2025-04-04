import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from app.utils.risk_matrix import RiskMatrix
from app.utils.general_format import format_general
from app.utils.risk_score import calculate_risk_score
from app.utils.border_format import add_merged_borders
import os

class PremitRiskScreen(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent  # Store the parent reference
        self.create_widgets()

    def create_widgets(self):
        # Create a main frame to hold all widgets
        main_frame = ttk.Frame(self)
        main_frame.grid(row=0, column=0, padx=20, pady=20)

        # Back Button (updated to use correct parent reference)
        back_button = ttk.Button(
            main_frame, 
            text="← Back", 
            command=lambda: self.parent.master.show_frame("StartScreen")
        )
        back_button.grid(row=0, column=0, sticky="nw", pady=(0, 10))

        # Title Label
        title_label = ttk.Label(main_frame, text="Pre-Mitigated dFMEA Formatting", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20))

        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection")
        file_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        # File selection button and label in same row
        self.file_button = ttk.Button(file_frame, text="Select File", command=self.select_file)
        self.file_button.grid(row=0, column=0, pady=10, padx=5)

        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # Add output filename input
        ttk.Label(file_frame, text="Output Filename:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.output_filename_var = tk.StringVar()
        self.output_filename_entry = ttk.Entry(file_frame, textvariable=self.output_filename_var, width=40)
        self.output_filename_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Add borders dropdown (new)
        ttk.Label(file_frame, text="Add Borders:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.borders_var = tk.StringVar(value="Yes")
        borders_dropdown = ttk.Combobox(file_frame, textvariable=self.borders_var, values=["Yes", "No"], width=5, state="readonly")
        borders_dropdown.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # Inputs section
        inputs_frame = ttk.LabelFrame(main_frame, text="Inputs")
        inputs_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

        # Create and layout input fields
        input_labels = [
            "Header Row",
            "Harm ID Column",
            "Occurrence Score Column",
            "Severity Score Column",
            "Risk Analysis Column"
        ]
        
        self.input_vars = {}  # Dictionary to store StringVar for each input
        for idx, label in enumerate(input_labels):
            # Calculate row and column positions
            row = idx // 2  # Integer division to determine row
            col = (idx % 2) * 2  # Even numbers (0, 2) for labels
            
            # Create label
            ttk.Label(inputs_frame, text=label).grid(
                row=row, column=col, padx=5, pady=5, sticky="e"
            )
            
            # Create entry with StringVar
            self.input_vars[label] = tk.StringVar()
            entry = ttk.Entry(inputs_frame, textvariable=self.input_vars[label], width=15)  # Set smaller width
            entry.grid(row=row, column=col + 1, padx=(5, 15), pady=5, sticky="w")  # Added right padding

        # Configure grid weights for inputs_frame columns
        inputs_frame.grid_columnconfigure(1, weight=0)  # No expansion for first entry column
        inputs_frame.grid_columnconfigure(3, weight=0)  # No expansion for second entry column

        # Risk Matrix section
        matrix_frame = ttk.LabelFrame(main_frame, text="Risk Matrix")
        matrix_frame.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")

        # Add the risk matrix
        self.risk_matrix = RiskMatrix(matrix_frame)
        self.risk_matrix.grid(row=0, column=0, padx=10, pady=10)

        # Generate Button
        style = ttk.Style()
        style.configure('Large.TButton', padding=10, font=('Arial', 16, 'bold'))
        self.generate_button = ttk.Button(main_frame, text="Generate", command=self.generate, style='Large.TButton')
        self.generate_button.grid(row=4, column=0, pady=20, sticky="ew")

        # Configure grid weights
        main_frame.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

    def select_file(self):
        filename = filedialog.askopenfilename(title="Select File")
        if filename:
            self.file_label.config(text=filename)
            # Set default output filename
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
            harm_id_col = self.input_vars["Harm ID Column"].get()
            occurrence_col = self.input_vars["Occurrence Score Column"].get()
            severity_col = self.input_vars["Severity Score Column"].get()
            risk_analysis_col = self.input_vars["Risk Analysis Column"].get()
            output_filename = self.output_filename_var.get()
            
            # Validate that all fields are filled
            if not all([header_row, harm_id_col, occurrence_col, severity_col, risk_analysis_col, output_filename]):
                messagebox.showerror("Error", "Please fill in all input fields")
                return
            
            # Get risk matrix values
            risk_matrix_values = self.risk_matrix.get_matrix_values()
            
            # Apply general formatting first
            wb = format_general(file_path, header_row, harm_id_col)
            
            # Then calculate risk scores
            wb = calculate_risk_score(
                wb, 
                occurrence_col, 
                severity_col, 
                risk_analysis_col, 
                header_row,
                risk_matrix_values
            )
            
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