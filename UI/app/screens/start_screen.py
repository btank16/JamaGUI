import tkinter as tk
from tkinter import ttk

class StartScreen(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Title
        title_label = ttk.Label(main_frame, text="Jama Format Machine", font=("Arial", 20, "bold"))
        title_label.grid(row=0, column=0, columnspan=5, pady=(0, 30))
        
        # Column Labels
        general_label = ttk.Label(main_frame, text="General", font=("Arial", 14, "bold"))
        general_label.grid(row=1, column=0, pady=(0, 20))
        
        dfmea_label = ttk.Label(main_frame, text="dFMEA", font=("Arial", 14, "bold"))
        dfmea_label.grid(row=1, column=2, pady=(0, 20))
        
        urra_label = ttk.Label(main_frame, text="URRA", font=("Arial", 14, "bold"))
        urra_label.grid(row=1, column=4, pady=(0, 20))
        
        # Vertical Separators
        separator1 = ttk.Separator(main_frame, orient='vertical')
        separator1.grid(row=1, column=1, rowspan=4, padx=30, sticky='ns')
        
        separator2 = ttk.Separator(main_frame, orient='vertical')
        separator2.grid(row=1, column=3, rowspan=4, padx=30, sticky='ns')
        
        # General Buttons
        trace_matrix = ttk.Button(main_frame, text="Trace Matrix Merge", command=lambda: self.parent.master.show_frame("TraceMatrixScreen"))
        trace_matrix.grid(row=2, column=0, pady=5)
        
        duplicate_search = ttk.Button(main_frame, text="Duplicate Search", command=lambda: self.parent.master.show_frame("DuplicateSearchScreen"))
        duplicate_search.grid(row=3, column=0, pady=5)
        
        # dFMEA Buttons
        dfmea_pre = ttk.Button(main_frame, text="dFMEA Pre-Mitigation Format", 
                              command=lambda: self.parent.master.show_frame("PremitRiskScreen"))
        dfmea_pre.grid(row=2, column=2, pady=5)
        
        dfmea_post = ttk.Button(main_frame, text="dFMEA Post-Mitigation Format", command=lambda: self.parent.master.show_frame("PostmitRiskScreen"))
        dfmea_post.grid(row=3, column=2, pady=5)
        
        dfmea_merge = ttk.Button(main_frame, text="dFMEA Risk Control Merge", command=lambda: self.parent.master.show_frame("RiskControlScreen"))
        dfmea_merge.grid(row=4, column=2, pady=5)
        
        # URRA Buttons
        urra_pre = ttk.Button(main_frame, text="URRA Pre-Mitigation Format", command=lambda: self.parent.master.show_frame("PremitURRAScreen"))
        urra_pre.grid(row=2, column=4, pady=5)
        
        urra_post = ttk.Button(main_frame, text="URRA Post-Mitigation Format", command=lambda: self.parent.master.show_frame("PostmitURRAScreen"))
        urra_post.grid(row=3, column=4, pady=5)
        
        urra_merge = ttk.Button(main_frame, text="URRA Risk Control Merge", command=lambda: self.parent.master.show_frame("RiskControlScreen"))
        urra_merge.grid(row=4, column=4, pady=5)
        
        marathon_ut = ttk.Button(main_frame, text="Marathon UT Format", 
                                command=lambda: self.parent.master.show_frame("MarathonUTScreen"))
        marathon_ut.grid(row=5, column=4, pady=5)
