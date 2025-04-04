import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from app.utils.risk_matrix import RiskMatrix
from app.utils.general_format_URRA import format_urra
from app.utils.risk_score import calculate_risk_score
from app.utils.border_format import add_merged_borders
import os

class PostmitURRAScreen(tk.Frame):
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
        main_frame.grid_rowconfigure(6, weight=1)  # Add weight to last row
        main_frame.grid_columnconfigure(0, weight=1)

        # Back Button
        back_button = ttk.Button(
            main_frame, 
            text="← Back", 
            command=lambda: self.parent.master.show_frame("StartScreen")
        )
        back_button.grid(row=1, column=0, sticky="nw", pady=(0, 10))

        # Title Label
        title_label = ttk.Label(main_frame, text="Post-Mitigated URRA Formatting", font=("Arial", 16, "bold"))
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

        # Add borders dropdown
        ttk.Label(file_frame, text="Add Borders:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.borders_var = tk.StringVar(value="Yes")
        borders_dropdown = ttk.Combobox(file_frame, textvariable=self.borders_var, values=["Yes", "No"], width=5, state="readonly")
        borders_dropdown.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # Inputs section
        inputs_frame = ttk.LabelFrame(main_frame, text="Inputs")
        inputs_frame.grid(row=4, column=0, padx=10, pady=10, sticky="ew")

        # Create and layout input fields
        input_labels = [
            "Header Row",
            "First Parent Column",
            "Second Parent Column",
            "Harm ID Column",
            "Severity Score Column",
            "Pre-Mitigation Occurrence Score Column",
            "Pre-Mitigation Risk Analysis Column",
            "Post-Mitigation Occurrence Score Column",
            "Post-Mitigation Risk Analysis Column"
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
            entry = ttk.Entry(inputs_frame, textvariable=self.input_vars[label], width=15)
            entry.grid(row=row, column=col + 1, padx=(5, 15), pady=5, sticky="w")

        # Configure grid weights for inputs_frame columns
        inputs_frame.grid_columnconfigure(1, weight=0)
        inputs_frame.grid_columnconfigure(3, weight=0)

        # Risk Matrix section
        matrix_frame = ttk.LabelFrame(main_frame, text="Risk Matrix")
        matrix_frame.grid(row=5, column=0, padx=10, pady=10, sticky="nsew")

        # Add the risk matrix
        self.risk_matrix = RiskMatrix(matrix_frame)
        self.risk_matrix.grid(row=0, column=0, padx=10, pady=10)

        # Generate Button
        style = ttk.Style()
        style.configure('Large.TButton', padding=10, font=('Arial', 16, 'bold'))
        self.generate_button = ttk.Button(main_frame, text="Generate", command=self.generate, style='Large.TButton')
        self.generate_button.grid(row=6, column=0, pady=20, sticky="ew")

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
            first_parent_col = self.input_vars["First Parent Column"].get()
            second_parent_col = self.input_vars["Second Parent Column"].get()
            harm_id_col = self.input_vars["Harm ID Column"].get()
            severity_col = self.input_vars["Severity Score Column"].get()
            pre_occurrence_col = self.input_vars["Pre-Mitigation Occurrence Score Column"].get()
            pre_risk_analysis_col = self.input_vars["Pre-Mitigation Risk Analysis Column"].get()
            post_occurrence_col = self.input_vars["Post-Mitigation Occurrence Score Column"].get()
            post_risk_analysis_col = self.input_vars["Post-Mitigation Risk Analysis Column"].get()
            output_filename = self.output_filename_var.get()
            
            # Validate that all fields are filled
            if not all([header_row, first_parent_col, second_parent_col, harm_id_col, severity_col, 
                       pre_occurrence_col, pre_risk_analysis_col,
                       post_occurrence_col, post_risk_analysis_col, output_filename]):
                messagebox.showerror("Error", "Please fill in all input fields")
                return
            
            # Get risk matrix values
            risk_matrix_values = self.risk_matrix.get_matrix_values()
            
            # Apply URRA formatting first (replaces format_general)
            wb = format_urra(
                file_path, 
                header_row,
                first_parent_col,
                second_parent_col,
                harm_id_col
            )
            
            # Calculate pre-mitigation risk scores
            wb = calculate_risk_score(
                wb, 
                pre_occurrence_col, 
                severity_col, 
                pre_risk_analysis_col, 
                header_row,
                risk_matrix_values
            )
            
            # Calculate post-mitigation risk scores
            wb = calculate_risk_score(
                wb, 
                post_occurrence_col, 
                severity_col, 
                post_risk_analysis_col, 
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