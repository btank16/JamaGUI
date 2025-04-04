import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
from app.utils.risk_control import merge_risk_control

class RiskControlScreen(tk.Frame):
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
        title_label = ttk.Label(main_frame, text="Risk Control Merge", font=("Arial", 16, "bold"))
        title_label.grid(row=2, column=0, pady=(0, 20))

        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection")
        file_frame.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

        # Risk Document file selection
        ttk.Label(file_frame, text="Risk Document:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.risk_file_button = ttk.Button(file_frame, text="Select File", command=lambda: self.select_file("risk"))
        self.risk_file_button.grid(row=0, column=1, pady=5, padx=5)
        self.risk_file_label = ttk.Label(file_frame, text="No file selected")
        self.risk_file_label.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        # Control Document file selection
        ttk.Label(file_frame, text="Control Document:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.control_file_button = ttk.Button(file_frame, text="Select File", command=lambda: self.select_file("control"))
        self.control_file_button.grid(row=1, column=1, pady=5, padx=5)
        self.control_file_label = ttk.Label(file_frame, text="No file selected")
        self.control_file_label.grid(row=1, column=2, padx=10, pady=5, sticky="w")

        # Output filename
        ttk.Label(file_frame, text="Output Filename:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.output_filename_var = tk.StringVar()
        self.output_filename_entry = ttk.Entry(file_frame, textvariable=self.output_filename_var, width=40)
        self.output_filename_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky="w")

        # Inputs section
        inputs_frame = ttk.LabelFrame(main_frame, text="Inputs")
        inputs_frame.grid(row=4, column=0, padx=10, pady=10, sticky="ew")

        # Create left frame (Risk Document)
        risk_frame = ttk.Frame(inputs_frame)
        risk_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        # Risk Document title
        risk_title = ttk.Label(risk_frame, text="Risk Document", font=("Arial", 12, "bold"))
        risk_title.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # Risk Document inputs
        risk_labels = ["Header Row", "Risk ID Column", "Paste Column"]
        self.risk_vars = {}
        for idx, label in enumerate(risk_labels):
            ttk.Label(risk_frame, text=label).grid(
                row=idx+1, column=0, padx=5, pady=5, sticky="e"
            )
            self.risk_vars[f"risk_{label}"] = tk.StringVar()
            entry = ttk.Entry(risk_frame, textvariable=self.risk_vars[f"risk_{label}"], width=15)
            entry.grid(row=idx+1, column=1, padx=5, pady=5, sticky="w")

        # Separator
        ttk.Separator(inputs_frame, orient="vertical").grid(
            row=0, column=1, sticky="ns", padx=10, pady=10
        )

        # Create right frame (Control Document)
        control_frame = ttk.Frame(inputs_frame)
        control_frame.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")
        
        # Control Document title
        control_title = ttk.Label(control_frame, text="Control Document", font=("Arial", 12, "bold"))
        control_title.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        # Control Document inputs
        control_labels = ["Header Row", "Risk ID Column", "Risk Control Column"]
        self.control_vars = {}
        for idx, label in enumerate(control_labels):
            ttk.Label(control_frame, text=label).grid(
                row=idx+1, column=0, padx=5, pady=5, sticky="e"
            )
            self.control_vars[f"control_{label}"] = tk.StringVar()
            entry = ttk.Entry(control_frame, textvariable=self.control_vars[f"control_{label}"], width=15)
            entry.grid(row=idx+1, column=1, padx=5, pady=5, sticky="w")

        # Configure grid weights for inputs_frame
        inputs_frame.grid_columnconfigure(0, weight=1)  # Risk Document side
        inputs_frame.grid_columnconfigure(1, weight=0)  # Separator
        inputs_frame.grid_columnconfigure(2, weight=1)  # Control Document side

        # Generate Button
        style = ttk.Style()
        style.configure('Large.TButton', padding=10, font=('Arial', 16, 'bold'))
        self.generate_button = ttk.Button(main_frame, text="Generate", command=self.generate, style='Large.TButton')
        self.generate_button.grid(row=5, column=0, pady=20, sticky="ew")

    def select_file(self, file_type):
        filename = filedialog.askopenfilename(title=f"Select {file_type.title()} Document")
        if filename:
            if file_type == "risk":
                self.risk_file_label.config(text=filename)
                # Set default output filename based on risk document name
                base_name = os.path.splitext(os.path.basename(filename))[0]
                self.output_filename_var.set(f"{base_name}_merged")
            else:
                self.control_file_label.config(text=filename)

    def generate(self):
        # Validate file selection
        if not hasattr(self, 'risk_file_label') or self.risk_file_label.cget('text') == "No file selected":
            messagebox.showerror("Error", "Please select a Risk Document")
            return
        if not hasattr(self, 'control_file_label') or self.control_file_label.cget('text') == "No file selected":
            messagebox.showerror("Error", "Please select a Control Document")
            return
            
        try:
            # Get file paths
            risk_file = self.risk_file_label.cget('text')
            control_file = self.control_file_label.cget('text')
            
            # Get input values
            risk_header_row = int(self.risk_vars["risk_Header Row"].get())
            control_header_row = int(self.control_vars["control_Header Row"].get())
            risk_id_col = self.risk_vars["risk_Risk ID Column"].get()
            control_id_col = self.control_vars["control_Risk ID Column"].get()
            paste_col = self.risk_vars["risk_Paste Column"].get()
            control_content_col = self.control_vars["control_Risk Control Column"].get()
            output_filename = self.output_filename_var.get()
            
            # Validate all inputs are provided
            if not all([risk_header_row, control_header_row, risk_id_col, 
                       control_id_col, paste_col, control_content_col, output_filename]):
                messagebox.showerror("Error", "Please fill in all input fields")
                return
            
            # Perform merge operation
            wb = merge_risk_control(
                risk_file, control_file,
                risk_header_row, control_header_row,
                risk_id_col, control_id_col,
                paste_col, control_content_col
            )
            
            # Save the final workbook with custom filename
            output_dir = os.path.dirname(risk_file)
            output_path = os.path.join(output_dir, f"{output_filename}.xlsx")
            wb.save(output_path)
            
            messagebox.showinfo("Success", f"File has been processed and saved as:\n{output_path}")
            
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")