import tkinter as tk
from tkinter import ttk, messagebox
from core.formula_comparator import ExcelFormulaComparator
from ui.workspace_view import Workspace
from core.mode_manager import ModeManager, AppMode
from ui.modes.inspect_mode import InspectMode

class ExcelToolsApp:
    """Main application class with mode management support"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.mode_manager = ModeManager(self.root)
        
        # Initialize UI components
        self.notebook = None
        self.comparator_frame = None
        self.workspace_frame = None
        self.mode_controls_frame = None
        
        # Component instances
        self.excel_comparator = None
        self.workspace = None
        
        # Window size management
        self.default_size = "900x1100"  # Default size for single worksheet
        self.expanded_size = "1800x1100"  # Expanded size for dual worksheets
        self.current_size_mode = "default"  # "default" or "expanded"
        
        self.setup_window()
        self.setup_mode_controls()
        self.setup_normal_mode()
        
        # Register for mode switch notifications
        self.mode_manager.register_mode_switch_callback(self.on_mode_switch)
    
    def on_closing(self):
        """Handle window closing event."""
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.root.destroy()

    def setup_window(self):
        """Setup the main window"""
        self.root.title("Excel Tools - Integrated")
        self.root.geometry("900x1100")
        self.root.attributes("-topmost", True)
        
        # Remove topmost after window is displayed so user can switch to other apps
        self.root.after(100, lambda: self.root.attributes("-topmost", False))
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_mode_controls(self):
        """Setup mode control buttons"""
        self.mode_controls_frame = ttk.Frame(self.root)
        self.mode_controls_frame.pack(pady=5, padx=10, fill="x")
        
        # Mode switch button
        self.mode_switch_btn = ttk.Button(
            self.mode_controls_frame, 
            text="Switch to Inspect Mode",
            command=self.toggle_mode
        )
        self.mode_switch_btn.pack(side=tk.LEFT, padx=5)
        
        # Always on top button
        self.always_on_top_btn = ttk.Button(
            self.mode_controls_frame,
            text="Always On Top: OFF",
            command=self.toggle_always_on_top
        )
        self.always_on_top_btn.pack(side=tk.LEFT, padx=5)
        
        # Window size control button
        self.window_size_btn = ttk.Button(
            self.mode_controls_frame,
            text="Reset Window Size",
            command=self.reset_window_size
        )
        self.window_size_btn.pack(side=tk.LEFT, padx=5)
        
        # Mode indicator label
        self.mode_label = ttk.Label(
            self.mode_controls_frame,
            text="Mode: Normal",
            font=("Arial", 10, "bold")
        )
        self.mode_label.pack(side=tk.RIGHT, padx=5)
    
    def setup_normal_mode(self):
        """Setup Normal Mode UI (existing functionality)"""
        if self.notebook:
            self.notebook.destroy()
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=10, padx=10, expand=True, fill="both")

        # Excel Formula Comparator tab
        self.comparator_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.comparator_frame, text='Excel Formula Comparator')
        self.excel_comparator = ExcelFormulaComparator(self.comparator_frame, self.root)

        # Workspace tab
        self.workspace_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.workspace_frame, text='Workspace')
        self.workspace = Workspace(self.workspace_frame)
    
    def setup_inspect_mode(self):
        """Setup Inspect Mode UI with simplified worksheet functionality"""
        if self.notebook:
            self.notebook.destroy()
        
        # Create Inspect Mode container
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=10, padx=10, expand=True, fill="both")
        
        # Create Inspect Mode frame
        inspect_frame = ttk.Frame(self.notebook)
        self.notebook.add(inspect_frame, text='Inspect Mode - Cell Analysis')
        
        # Import and create Inspect Mode with simplified worksheet functionality
        from ui.modes.inspect_mode import InspectMode
        self.inspect_mode = InspectMode(inspect_frame, self.root)
    
    def toggle_mode(self):
        """Toggle between Normal and Inspect modes"""
        self.mode_manager.toggle_mode()
    
    def toggle_always_on_top(self):
        """Toggle always on top functionality"""
        self.mode_manager.toggle_always_on_top()
        self.update_always_on_top_button()
    
    def update_always_on_top_button(self):
        """Update the always on top button text"""
        status = "ON" if self.mode_manager.always_on_top else "OFF"
        self.always_on_top_btn.config(text=f"Always On Top: {status}")
    
    def reset_window_size(self):
        """Reset window to default size and hide worksheet2 interface"""
        self.root.geometry(self.default_size)
        self.current_size_mode = "default"
        self.window_size_btn.config(text="Reset Window Size")
        
        # Hide worksheet2 interface in Normal Mode
        if hasattr(self, 'excel_comparator') and self.excel_comparator:
            try:
                # Use the new hide_worksheet2_interface method
                self.excel_comparator.hide_worksheet2_interface()
            except Exception as e:
                print(f"Could not hide worksheet2 interface: {e}")
        
        print(f"Window size reset to default: {self.default_size}")
    
    def expand_window_size(self):
        """Expand window for dual worksheet view"""
        self.root.geometry(self.expanded_size)
        self.current_size_mode = "expanded"
        self.window_size_btn.config(text="Reset Window Size")
        print(f"Window size expanded to: {self.expanded_size}")
    
    def toggle_window_size(self):
        """Toggle between default and expanded window sizes"""
        if self.current_size_mode == "default":
            self.expand_window_size()
        else:
            self.reset_window_size()
    
    def on_mode_switch(self, old_mode, new_mode):
        """Handle mode switch events"""
        if new_mode == AppMode.NORMAL:
            self.setup_normal_mode()
            self.mode_switch_btn.config(text="Switch to Inspect Mode")
            self.mode_label.config(text="Mode: Normal")
        elif new_mode == AppMode.INSPECT:
            self.setup_inspect_mode()
            self.mode_switch_btn.config(text="Switch to Normal Mode")
            self.mode_label.config(text="Mode: Inspect")
        
        # Update always on top button
        self.update_always_on_top_button()
    
    def run(self):
        """Start the application"""
        self.root.mainloop()

def main():
    """Main entry point"""
    app = ExcelToolsApp()
    app.run()

if __name__ == "__main__":
    main()