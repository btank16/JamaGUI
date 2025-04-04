import tkinter as tk
from tkinter import ttk

class RiskMatrix(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # Define headers and labels
        self.occurrence_labels = [
            "Frequent 5", "Likely 4", "Occasional 3", "Remote 2", "Incredible 1"
        ]
        self.severity_labels = [
            "Negligible 1", "Minor 2", "Serious 3", "Critical 4", "Catastrophic 5"
        ]
        self.risk_options = ["LOW", "MOD", "INT"]
        
        # Default matrix values - can be overridden
        self.default_values = [
            ["LOW", "MOD", "MOD", "INT", "INT"],      # Frequent 5
            ["LOW", "MOD", "MOD", "MOD", "INT"],      # Likely 4
            ["LOW", "LOW", "LOW", "MOD", "INT"],      # Occasional 3
            ["LOW", "LOW", "LOW", "LOW", "MOD"],      # Remote 2
            ["LOW", "LOW", "LOW", "LOW", "LOW"]       # Incredible 1
        ]
        
        self.create_matrix()

    def create_matrix(self):
        # Create empty corner cell for alignment
        empty_corner = ttk.Label(self, text="")
        empty_corner.grid(row=0, column=0, padx=5, pady=5)
        
        # Create "SEVERITY" label above the first row
        severity_label = ttk.Label(self, text="SEVERITY", font=("Arial", 10, "bold"))
        severity_label.grid(row=0, column=2, columnspan=5, padx=5, pady=5)
        
        # Create "OCCURRENCE" label to the left of the leftmost column
        occurrence_label = ttk.Label(self, text="OCCURRENCE", font=("Arial", 10, "bold"))
        occurrence_label.grid(row=2, column=0, rowspan=5, padx=(15,5), pady=5)
        
        # Create severity headers
        for col, label in enumerate(self.severity_labels, start=1):
            header = ttk.Label(self, text=label, font=("Arial", 10, "bold"))
            header.grid(row=1, column=col+1, padx=5, pady=5)
        
        self.dropdowns = []
        
        for row, label in enumerate(self.occurrence_labels, start=2):
            # Add occurrence label
            row_label = ttk.Label(self, text=label, font=("Arial", 10, "bold"))
            row_label.grid(row=row, column=1, padx=5, pady=5, sticky='w')
            
            row_dropdowns = []
            # Add dropdowns for each cell
            for col in range(5):
                combo = ttk.Combobox(self, values=self.risk_options, width=5, state='readonly')
                combo.set(self.default_values[row-2][col])
                combo.grid(row=row, column=col+2, padx=5, pady=5)
                
                row_dropdowns.append(combo)
            self.dropdowns.append(row_dropdowns)

        # Configure grid weights
        for i in range(8):  # Adjust based on total rows
            self.grid_rowconfigure(i, weight=1)
        for i in range(8):  # Adjust based on total columns
            self.grid_columnconfigure(i, weight=1)

    def get_matrix_values(self):
        """Returns current values of all cells as a 2D list"""
        return [[combo.get() for combo in row] for row in self.dropdowns]

    def set_matrix_values(self, values):
        """Sets values for all cells from a 2D list"""
        for i, row in enumerate(values):
            for j, value in enumerate(row):
                if value in self.risk_options:
                    self.dropdowns[i][j].set(value)

    def set_cell_value(self, row, col, value):
        """Sets value for a specific cell"""
        if 0 <= row < 5 and 0 <= col < 5 and value in self.risk_options:
            self.dropdowns[row][col].set(value)
