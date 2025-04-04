import tkinter as tk
from tkinter import ttk
from app.screens.start_screen import StartScreen
from app.screens.premit_risk_screen import PremitRiskScreen
from app.screens.postmit_risk_screen import PostmitRiskScreen
from app.screens.trace_matrix import TraceMatrixScreen
from app.screens.risk_control_merge import RiskControlScreen
from app.screens.premit_urra_screen import PremitURRAScreen
from app.screens.postmit_urra_screen import PostmitURRAScreen
from app.screens.duplicate_search_screen import DuplicateSearchScreen
from app.screens.marathon_ut_screen import MarathonUTScreen

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Jama Format Machine")
        self.geometry("1000x850")
        
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        self.frames = {}
        for F in (
            StartScreen,
            TraceMatrixScreen,
            PremitRiskScreen,
            PostmitRiskScreen,
            RiskControlScreen,
            PremitURRAScreen,
            PostmitURRAScreen,
            DuplicateSearchScreen,
            MarathonUTScreen
        ):
            frame = F(container)
            self.frames[F.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        
        self.show_frame("StartScreen")
    
    def show_frame(self, frame_name):
        frame = self.frames[frame_name]
        frame.tkraise()

if __name__ == "__main__":
    app = App()
    app.mainloop()
