import os
import re
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

# Find and parse SolidWorks parameter file
def get_app_dir():
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(__file__)

def find_file():
    for f in os.listdir(get_app_dir()):
        if f.endswith('.txt'):
            return os.path.join(get_app_dir(), f)
    raise FileNotFoundError("No .txt file found")

def parse_line(line):
    m = re.match(r'^\s*"([^"]+)"\s*=\s*([\d.]+)mm', line.strip())
    return (m.group(1), float(m.group(2))) if m else None

def load_data(file_path=None):
    if not file_path:
        file_path = find_file()
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        lines = f.readlines()
    params = [parse_line(line) for line in lines]
    return [p for p in params if p], lines, file_path

def save_data(params, original_lines, file_path):
    param_dict = dict(params)
    output = []
    for line in original_lines:
        parsed = parse_line(line)
        if parsed and parsed[0] in param_dict:
            name = parsed[0]
            value = int(param_dict[name]) if param_dict[name].is_integer() else param_dict[name]
            output.append(f'"{name}"= {value}mm\n')
        else:
            output.append(line if line.endswith('\n') else line + '\n')
    
    with open(file_path, 'w', encoding='utf-8-sig') as f:
        f.writelines(output)

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Room Dimensions")
        
        # Initialize data
        self.params = []
        self.original_lines = []
        self.file_path = ""
        self.entries = {}
        self.preset_entries = {}
        
        # Try to load default file
        try:
            self.params, self.original_lines, self.file_path = load_data()
        except FileNotFoundError:
            pass  # Will show empty UI with browse option
        
        # Calculate window size based on parameters and ensure adequate space for preset tab
        num_params = len(self.params)
        # Minimum height to accommodate preset generator tab with all sections
        min_height_for_presets = 650  # Header + File Management + Instructions + Parameters + Button
        editor_height = 200 + (num_params * 30)
        window_height = max(min_height_for_presets, editor_height)
        window_width = 700  # Wider for better preset functionality
        
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.resizable(True, True)
        self.root.minsize(600, 500)  # Set minimum window size
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.editor_frame = ttk.Frame(self.notebook)
        self.preset_frame = ttk.Frame(self.notebook)
        
        self.notebook.add(self.editor_frame, text="Parameter Editor")
        self.notebook.add(self.preset_frame, text="Preset Generator")
        
        # Setup editor tab
        self.setup_editor_tab()
        
        # Setup preset tab
        self.setup_preset_tab()
    
    def setup_editor_tab(self):
        # File selection frame with better styling
        file_frame = ttk.LabelFrame(self.editor_frame, text="📁 Current File", padding=10)
        file_frame.pack(fill='x', padx=20, pady=10)
        
        # File info section
        file_info_frame = ttk.Frame(file_frame)
        file_info_frame.pack(fill='x', pady=(0, 8))
        
        # File status icon and name
        file_status_frame = ttk.Frame(file_info_frame)
        file_status_frame.pack(side='left', fill='x', expand=True)
        
        if self.file_path:
            ttk.Label(file_status_frame, text="✅", font=('Arial', 10)).pack(side='left')
            self.file_label = ttk.Label(file_status_frame, text=os.path.basename(self.file_path), 
                                       foreground='#27ae60', font=('Arial', 10, 'bold'))
            self.file_label.pack(side='left', padx=(5, 0))
            
            # File path
            ttk.Label(file_info_frame, text=f"📂 {os.path.dirname(self.file_path)}", 
                     font=('Arial', 8), foreground='#7f8c8d').pack(anchor='w', pady=(2, 0))
        else:
            ttk.Label(file_status_frame, text="❌", font=('Arial', 10)).pack(side='left')
            self.file_label = ttk.Label(file_status_frame, text="No file selected", 
                                       foreground='#e74c3c', font=('Arial', 10))
            self.file_label.pack(side='left', padx=(5, 0))
        
        # Action buttons
        button_frame = ttk.Frame(file_frame)
        button_frame.pack(fill='x')
        
        ttk.Button(button_frame, text="📂 Browse Files", command=self.browse_file).pack(side='left')
        
        if self.params:
            ttk.Button(button_frame, text="🔄 Reload", command=self.reload_file).pack(side='left', padx=(8, 0))
        
        # Help button
        ttk.Button(button_frame, text="❓ Help", command=self.show_help).pack(side='right')
        
        # Mode toggle frame with better styling
        mode_frame = ttk.LabelFrame(self.editor_frame, text="📝 Editor Mode", padding=10)
        mode_frame.pack(fill='x', padx=20, pady=(0, 15))
        
        # Mode description
        mode_desc_frame = ttk.Frame(mode_frame)
        mode_desc_frame.pack(fill='x')
        
        self.editor_mode = tk.StringVar(value="Advanced")
        
        # Current mode indicator
        self.mode_indicator = ttk.Label(mode_desc_frame, text="🔧 Advanced Mode - Direct value editing", 
                                       font=('Arial', 9), foreground='#2c3e50')
        self.mode_indicator.pack(side='left')
        
        # Toggle button with better styling
        self.mode_button = ttk.Button(mode_desc_frame, text="🎛️ Switch to Simple Mode", 
                                     command=self.toggle_editor_mode)
        self.mode_button.pack(side='right')
        
        # Create scrollable parameter area
        self.create_parameter_area()
        
        # Save button frame with better styling
        button_frame = ttk.Frame(self.editor_frame)
        button_frame.pack(fill='x', padx=20, pady=15, side='bottom')
        
        # Save button with icon and better styling
        save_button = ttk.Button(button_frame, text="💾 Save Changes", command=self.save)
        save_button.pack(side='right', ipadx=10, ipady=5)
        
        # Status text
        ttk.Label(button_frame, text="💡 Remember to save your changes!", 
                 font=('Arial', 8), foreground='#7f8c8d').pack(side='left')
    
    def create_parameter_area(self):
        # Clear existing parameter area if it exists
        for widget in self.editor_frame.winfo_children():
            if isinstance(widget, tk.Canvas):
                widget.destroy()
        
        if not self.params:
            ttk.Label(self.editor_frame, text="No parameters found. Please select a valid SolidWorks file.", 
                     foreground='gray').pack(pady=20)
            return
        
        mode = getattr(self, 'editor_mode', None)
        if mode and mode.get() == "Simple":
            self.create_simple_parameter_area()
        else:
            self.create_advanced_parameter_area()
    
    def create_advanced_parameter_area(self):
        # Scrollable frame for parameters
        canvas = tk.Canvas(self.editor_frame)
        scrollbar = ttk.Scrollbar(self.editor_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # Configure scrollable frame to update scroll region when content changes
        def configure_scroll_region(event=None):
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        scrollable_frame.bind("<Configure>", configure_scroll_region)
        
        # Create window in canvas and configure scrolling
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Make sure the scrollable frame fills the canvas width
        def configure_canvas_width(event=None):
            canvas_width = canvas.winfo_width()
            canvas.itemconfig(canvas_window, width=canvas_width)
        
        canvas.bind("<Configure>", configure_canvas_width)
        
        # Only show scrollbar if needed (more than 15 parameters)
        num_params = len(self.params)
        if num_params > 15:
            canvas.pack(side="left", fill="both", expand=True, padx=(20, 0))
            scrollbar.pack(side="right", fill="y")
        else:
            canvas.pack(fill="both", expand=True, padx=20)
        
        # Create parameter entries
        self.entries = {}
        for i, (name, value) in enumerate(self.params):
            ttk.Label(scrollable_frame, text=name).grid(row=i, column=0, sticky='w', pady=2, padx=(0, 10))
            entry = ttk.Entry(scrollable_frame, width=12)
            entry.insert(0, str(int(value) if value.is_integer() else value))
            entry.grid(row=i, column=1, pady=2, padx=(0, 5))
            ttk.Label(scrollable_frame, text="mm").grid(row=i, column=2, sticky='w', pady=2)
            self.entries[name] = entry
    
    def create_simple_parameter_area(self):
        # Create scrollable frame
        canvas = tk.Canvas(self.editor_frame)
        scrollbar = ttk.Scrollbar(self.editor_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Filter parameters that have presets
        preset_params = []
        for name, value in self.params:
            if hasattr(self, 'loaded_presets') and name in self.loaded_presets:
                preset_params.append((name, value))
        
        if not preset_params:
            # No presets message with better styling
            no_presets_frame = ttk.Frame(scrollable_frame)
            no_presets_frame.pack(expand=True, fill='both', padx=30, pady=50)
            
            # Icon and main message
            ttk.Label(no_presets_frame, text="🎯", font=('Arial', 24)).pack(pady=(0, 10))
            ttk.Label(no_presets_frame, text="No Presets Available", 
                     font=('Arial', 14, 'bold'), foreground='#34495e').pack()
            ttk.Label(no_presets_frame, text="Simple Mode requires preset values to create sliders.", 
                     font=('Arial', 10), foreground='#7f8c8d').pack(pady=(5, 15))
            
            # Instructions
            instructions_frame = ttk.Frame(no_presets_frame)
            instructions_frame.pack()
            
            ttk.Label(instructions_frame, text="📋 To get started:", 
                     font=('Arial', 10, 'bold'), foreground='#2c3e50').pack(anchor='w')
            ttk.Label(instructions_frame, text="1. Switch to the 'Preset Generator' tab", 
                     font=('Arial', 9), foreground='#34495e').pack(anchor='w', padx=(15, 0))
            ttk.Label(instructions_frame, text="2. Create presets for your parameters", 
                     font=('Arial', 9), foreground='#34495e').pack(anchor='w', padx=(15, 0))
            ttk.Label(instructions_frame, text="3. Return here to use the sliders", 
                     font=('Arial', 9), foreground='#34495e').pack(anchor='w', padx=(15, 0))
            
            canvas.pack(side="left", fill="both", expand=True, padx=(20, 0))
            return
        
        # Add header for Simple Mode with better styling
        header_frame = ttk.Frame(scrollable_frame)
        header_frame.pack(fill='x', padx=15, pady=(10, 20))
        
        # Main title with icon
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(fill='x')
        
        ttk.Label(title_frame, text="🎛️ Simple Mode", 
                 font=('Arial', 14, 'bold'), foreground='#2c3e50').pack(side='left')
        
        # Status indicator
        status_frame = ttk.Frame(header_frame)
        status_frame.pack(fill='x', pady=(5, 0))
        
        ttk.Label(status_frame, text="📁", font=('Arial', 10)).pack(side='left')
        file_name = os.path.basename(self.file_path) if self.file_path else "No file"
        ttk.Label(status_frame, text=f"Editing: {file_name}", 
                 font=('Arial', 9), foreground='#34495e').pack(side='left', padx=(5, 0))
        
        # Instructions
        ttk.Label(header_frame, text="💡 Drag sliders to adjust dimensions, then click Save to apply changes", 
                 font=('Arial', 9), foreground='#7f8c8d').pack(anchor='w', pady=(8, 0))
        
        # Parameter sliders
        self.entries = {}
        self.sliders = {}
        
        for i, (name, value) in enumerate(preset_params):
            preset_values = sorted(self.loaded_presets[name])  # Sort preset values
            min_val = min(preset_values)
            max_val = max(preset_values)
            
            # Main parameter container with subtle border
            param_frame = ttk.LabelFrame(scrollable_frame, text="", padding=15)
            param_frame.pack(fill='x', padx=15, pady=8)
            
            # Parameter name and current value display
            header_frame = ttk.Frame(param_frame)
            header_frame.pack(fill='x', pady=(0, 10))
            
            # Parameter name with icon
            name_frame = ttk.Frame(header_frame)
            name_frame.pack(side='left')
            
            ttk.Label(name_frame, text="📏", font=('Arial', 10)).pack(side='left')
            ttk.Label(name_frame, text=f"{name}", font=('Arial', 11, 'bold'), 
                     foreground='#2c3e50').pack(side='left', padx=(5, 0))
            
            # Current value display (larger and more prominent)
            value_frame = ttk.Frame(header_frame)
            value_frame.pack(side='right')
            
            value_label = ttk.Label(value_frame, text=f"{int(value) if value.is_integer() else value}", 
                                   font=('Arial', 16, 'bold'), foreground='#e74c3c')
            value_label.pack(side='left')
            ttk.Label(value_frame, text="mm", font=('Arial', 10), 
                     foreground='#7f8c8d').pack(side='left', padx=(2, 0))
            
            # Range info with min/max indicators
            range_frame = ttk.Frame(param_frame)
            range_frame.pack(fill='x', pady=(0, 8))
            
            ttk.Label(range_frame, text=f"Min: {int(min_val) if min_val.is_integer() else min_val}mm", 
                     font=('Arial', 8), foreground='#95a5a6').pack(side='left')
            ttk.Label(range_frame, text=f"Max: {int(max_val) if max_val.is_integer() else max_val}mm", 
                     font=('Arial', 8), foreground='#95a5a6').pack(side='right')
            
            # Find current value index in preset values (or closest)
            current_index = 0
            if value in preset_values:
                current_index = preset_values.index(value)
            else:
                # Find closest preset value
                closest_val = min(preset_values, key=lambda x: abs(x - value))
                current_index = preset_values.index(closest_val)
            
            # Improved slider with better styling
            slider_frame = ttk.Frame(param_frame)
            slider_frame.pack(fill='x', pady=(0, 5))
            
            slider = tk.Scale(slider_frame, from_=0, to=len(preset_values)-1, orient='horizontal', 
                             resolution=1, length=350, showvalue=0, 
                             bg='#ecf0f1', fg='#2c3e50', troughcolor='#bdc3c7',
                             highlightthickness=0, bd=1)
            slider.set(current_index)
            slider.pack(fill='x')
            
            # Update value label when slider changes with improved feedback
            def update_value_label(index_str, label=value_label, param_name=name, values=preset_values):
                index = int(index_str)
                actual_value = values[index]
                display_val = int(actual_value) if actual_value.is_integer() else actual_value
                
                # Update the large value display
                label.config(text=f"{display_val}")
                
                # Brief visual feedback (color change)
                label.config(foreground='#27ae60')  # Green when changing
                label.after(200, lambda: label.config(foreground='#e74c3c'))  # Back to red
                
                # Update the parameter value in self.params
                for j, (pname, pval) in enumerate(self.params):
                    if pname == param_name:
                        self.params[j] = (pname, actual_value)
                        break
            
            slider.config(command=update_value_label)
            
            # Store references
            self.sliders[name] = slider
            
            # Create a simple entry object for compatibility with save function
            class SimpleEntry:
                def __init__(self, slider_ref, preset_vals):
                    self.slider = slider_ref
                    self.preset_values = preset_vals
                
                def get(self):
                    try:
                        index = int(self.slider.get())
                        return str(self.preset_values[index])
                    except (ValueError, IndexError):
                        return "0"  # Fallback value
            
            self.entries[name] = SimpleEntry(slider, preset_values) 
        
        # Pack canvas and scrollbar - always enable scrolling for better UX
        canvas.pack(side="left", fill="both", expand=True, padx=(20, 0))
        scrollbar.pack(side="right", fill="y")
        
        # Enable mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        # Bind mouse wheel to canvas and scrollable frame
        canvas.bind("<MouseWheel>", _on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        
        # Update scroll region when content changes
        def update_scroll_region():
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        # Call update after a short delay to ensure all widgets are rendered
        canvas.after(100, update_scroll_region)
    
    def show_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("❓ Room Dimensions Editor - Help")
        help_window.geometry("700x600")
        help_window.resizable(True, True)
        
        # Create scrollable text area
        text_frame = ttk.Frame(help_window)
        text_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Scrollable text widget
        text_widget = tk.Text(text_frame, wrap='word', font=('Arial', 10), 
                             bg='#f8f9fa', fg='#2c3e50', padx=15, pady=15)
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # Help content
        help_content = """🏠 ROOM DIMENSIONS EDITOR - USER GUIDE

📋 WHAT IS THIS PROGRAM?
This program helps you easily modify room dimensions in SolidWorks parameter files. It's designed for outdoor room businesses to quickly adjust dimensions like wall heights, lengths, stud spacing, and other structural parameters.

🎯 USE CASES:
• Customize room dimensions for different customer requirements
• Quickly generate multiple size variations of outdoor rooms
• Test different dimension combinations using preset values
• Streamline the design process for outdoor structures

📁 FILE REQUIREMENTS:

The program works with SolidWorks dimension text files that contain parameters in this format:
"Parameter_Name"= ValueUnit

Example file content (dimenstion.txt):
"External"= 1000mm
"External_length"= 500mm
"Stud_center_spacing"= 600mm
"Internal_height"= 2400mm

📂 FILE PLACEMENT:
• Place your dimension file (e.g., "dimenstion.txt") in the same folder as this program
• Or use the "Browse Files" button to select any .txt file from your computer
• The program will automatically load "dimenstion.txt" if found in the program folder

🔧 HOW TO USE:

1️⃣ ADVANCED MODE (Direct Editing):
   • Edit parameter values directly in text boxes
   • Type exact values you want
   • Click "Save Changes" to update the file

2️⃣ SIMPLE MODE (Slider Controls):
   • First, create presets in the "Preset Generator" tab
   • Switch to Simple Mode to use sliders
   • Drag sliders to select from preset values
   • Click "Save Changes" to apply

3️⃣ PRESET GENERATOR:
   • Check parameters you want to create presets for
   • Enter comma-separated values (e.g., 200,300,400,500)
   • Click "Generate/Update Preset File" to save
   • These presets become available as slider options in Simple Mode

💡 WORKFLOW TIPS:

For New Users:
1. Load your SolidWorks dimension file
2. Try Advanced Mode first to understand your parameters
3. Create presets for commonly changed dimensions
4. Use Simple Mode for quick adjustments

For Regular Use:
1. Set up presets once for your standard sizes
2. Use Simple Mode for daily dimension changes
3. Save changes to update your SolidWorks file
4. Import the updated file back into SolidWorks

🔄 FILE UPDATES:
When you save changes, the program updates your original text file. You can then:
• Import this file back into SolidWorks
• Use it to update your 3D models
• Share it with team members

⚠️ IMPORTANT NOTES:
• Always backup your original files before editing
• The program preserves the exact format SolidWorks expects
• Parameter names must match exactly with your SolidWorks model
• Values are automatically formatted with proper units (mm)

🆘 TROUBLESHOOTING:
• If no parameters appear: Check your file format matches the example above
• If Simple Mode is empty: Create presets first in the Preset Generator tab
• If save fails: Ensure the file isn't open in another program
• If values don't update: Check parameter names match your SolidWorks model exactly"""

        # Insert content and configure tags for styling
        text_widget.insert('1.0', help_content)
        
        # Configure text styling
        text_widget.tag_configure('heading', font=('Arial', 12, 'bold'), foreground='#2c3e50')
        text_widget.tag_configure('subheading', font=('Arial', 11, 'bold'), foreground='#34495e')
        text_widget.tag_configure('code', font=('Courier', 9), background='#ecf0f1', foreground='#e74c3c')
        
        # Make text read-only
        text_widget.config(state='disabled')
        
        # Pack widgets
        text_widget.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Close button
        close_frame = ttk.Frame(help_window)
        close_frame.pack(fill='x', padx=20, pady=(0, 20))
        ttk.Button(close_frame, text="✅ Got it!", command=help_window.destroy).pack(side='right')
        
        # Center the window
        help_window.transient(self.root)
        help_window.grab_set()

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select SolidWorks Parameter File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialdir=get_app_dir()
        )
        
        if file_path:
            try:
                self.params, self.original_lines, self.file_path = load_data(file_path)
                self.file_label.config(text=os.path.basename(self.file_path))
                self.create_parameter_area()
                self.update_preset_tab()
                
                # Update window size for new parameters
                num_params = len(self.params)
                window_height = max(400, 200 + (num_params * 30))
                self.root.geometry(f"600x{window_height}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{e}")
    
    def reload_file(self):
        if self.file_path:
            try:
                self.params, self.original_lines, _ = load_data(self.file_path)
                self.create_parameter_area()
                messagebox.showinfo("Success", "File reloaded!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to reload file:\n{e}")
    
    def toggle_editor_mode(self):
        current_mode = self.editor_mode.get()
        if current_mode == "Advanced":
            # Check if preset file is available for simple mode
            if not hasattr(self, 'loaded_presets') or not self.loaded_presets:
                # Try to load default preset file
                default_preset_path = os.path.join(get_app_dir(), "presets.txt")
                if os.path.exists(default_preset_path):
                    self.load_preset_file(default_preset_path)
                
                if not hasattr(self, 'loaded_presets') or not self.loaded_presets:
                    messagebox.showwarning("No Presets", "Simple mode requires a preset file. Please create presets first or switch to the Preset Generator tab.")
                    return
            
            self.editor_mode.set("Simple")
            self.mode_button.config(text="🔧 Switch to Advanced Mode")
            self.mode_indicator.config(text="🎛️ Simple Mode - Preset sliders", foreground='#27ae60')
        else:
            self.editor_mode.set("Advanced")
            self.mode_button.config(text="🎛️ Switch to Simple Mode")
            self.mode_indicator.config(text="🔧 Advanced Mode - Direct value editing", foreground='#2c3e50')
        
        # Refresh the parameter area with the new mode
        self.create_parameter_area()
    
    def setup_preset_tab(self):
        # Header section with title and description
        header_frame = ttk.Frame(self.preset_frame)
        header_frame.pack(fill='x', padx=20, pady=10)
        
        title_label = ttk.Label(header_frame, text="🎯 Preset Generator", font=('Arial', 14, 'bold'))
        title_label.pack(anchor='w')
        
        desc_label = ttk.Label(header_frame, 
                              text="Create quick-select presets for your parameters. Perfect for testing different dimensions!",
                              font=('Arial', 9), foreground='#666666')
        desc_label.pack(anchor='w', pady=(2, 0))
        
        # Preset file section with colored background
        file_section = ttk.LabelFrame(self.preset_frame, text="📁 Preset File Management", padding=10)
        file_section.pack(fill='x', padx=20, pady=(0, 8))
        
        # File status frame
        file_status_frame = ttk.Frame(file_section)
        file_status_frame.pack(fill='x', pady=(0, 8))
        
        ttk.Label(file_status_frame, text="Current file:", font=('Arial', 9, 'bold')).pack(side='left')
        self.preset_file_label = ttk.Label(file_status_frame, text="No preset file loaded", 
                                          foreground='#ff6b6b', font=('Arial', 9))
        self.preset_file_label.pack(side='left', padx=(8, 0))
        
        # File action buttons
        file_buttons_frame = ttk.Frame(file_section)
        file_buttons_frame.pack(fill='x')
        
        load_btn = ttk.Button(file_buttons_frame, text="🔄 Load Default", command=self.load_default_preset)
        load_btn.pack(side='left', padx=(0, 5))
        
        browse_btn = ttk.Button(file_buttons_frame, text="📂 Browse File", command=self.browse_preset_file)
        browse_btn.pack(side='left')
        
        # Instructions section
        instructions_frame = ttk.LabelFrame(self.preset_frame, text="📝 How to Use", padding=8)
        instructions_frame.pack(fill='x', padx=20, pady=(0, 8))
        
        instructions = [
            "1. ✅ Check the parameters you want to create presets for",
            "2. ✏️ Enter comma-separated values (e.g., 200,300,400,500)",
            "3. 💾 Click 'Generate/Update' to save your preset file",
            "4. 🎛️ Use 'Simple Mode' in Parameter Editor to access sliders"
        ]
        
        for instruction in instructions:
            ttk.Label(instructions_frame, text=instruction, font=('Arial', 8), 
                     foreground='#4a4a4a').pack(anchor='w', pady=1)
        
        # Action buttons section - pack this FIRST to ensure it's always visible
        action_frame = ttk.Frame(self.preset_frame)
        action_frame.pack(fill='x', padx=20, pady=(5, 10), side='bottom')
        
        # Help text
        help_label = ttk.Label(action_frame, 
                              text="💡 Tip: Generated presets will be used for slider ranges in Simple Mode",
                              font=('Arial', 8), foreground='#666666')
        help_label.pack(side='left')
        
        # Generate button with icon
        generate_btn = ttk.Button(action_frame, text="💾 Generate/Update Preset File", 
                                 command=self.generate_presets)
        generate_btn.pack(side='right')
        
        # Parameters section - now pack this to fill remaining space
        params_section = ttk.LabelFrame(self.preset_frame, text="⚙️ Parameter Presets", padding=10)
        params_section.pack(fill='both', expand=True, padx=20, pady=(0, 5))
        
        # Scrollable frame for preset entries
        self.preset_canvas = tk.Canvas(params_section, highlightthickness=0)
        preset_scrollbar = ttk.Scrollbar(params_section, orient="vertical", command=self.preset_canvas.yview)
        self.preset_scrollable = ttk.Frame(self.preset_canvas)
        
        self.preset_scrollable.bind(
            "<Configure>",
            lambda e: self.preset_canvas.configure(scrollregion=self.preset_canvas.bbox("all"))
        )
        
        self.preset_canvas.create_window((0, 0), window=self.preset_scrollable, anchor="nw")
        self.preset_canvas.configure(yscrollcommand=preset_scrollbar.set)
        
        self.preset_canvas.pack(side="left", fill="both", expand=True)
        preset_scrollbar.pack(side="right", fill="y")
        
        # Initialize preset file path and load default if exists
        self.preset_file_path = None
        self.loaded_presets = {}
        self.load_default_preset()
        
        # Update preset tab with current parameters
        self.update_preset_tab()
    
    def browse_preset_file(self):
        file_path = filedialog.askopenfilename(
            title="🔍 Select Preset File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialdir=get_app_dir()
        )
        if file_path:
            self.load_preset_file(file_path)
    
    def load_default_preset(self):
        # Try to find preset file based on main file name
        preset_file_found = None
        search_files = []
        
        if self.file_path:
            # Get the base name without extension
            base_name = os.path.splitext(os.path.basename(self.file_path))[0]
            app_dir = get_app_dir()
            
            # Try different preset file naming patterns
            preset_suffixes = ['_presets', '_preset', '_p']
            
            for suffix in preset_suffixes:
                preset_filename = f"{base_name}{suffix}.txt"
                preset_path = os.path.join(app_dir, preset_filename)
                search_files.append(preset_filename)
                if os.path.exists(preset_path):
                    preset_file_found = preset_path
                    break
        
        # Fallback to generic presets.txt
        if not preset_file_found:
            default_preset_path = os.path.join(get_app_dir(), "presets.txt")
            search_files.append("presets.txt")
            if os.path.exists(default_preset_path):
                preset_file_found = default_preset_path
        
        if preset_file_found:
            self.load_preset_file(preset_file_found)
        else: 
            self.preset_file_path = None
            self.loaded_presets = {}
            self.preset_file_label.config(text="📄 No preset file found")
            
            # Show helpful message about what files were searched
            if search_files:
                searched_info = "Searched for:\n• " + "\n• ".join(search_files)
                messagebox.showinfo("No Preset File", 
                                  f"🔍 No preset file found.\n\n{searched_info}\n\n💡 Create presets using the form below!")
            
            self.update_preset_tab()
    
    def load_preset_file(self, file_path):
        try:
            self.loaded_presets = {}
            param_count = 0
            
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                for line in f:
                    line = line.strip()
                    # Skip comments and empty lines
                    if line and not line.startswith('#'):
                        # Parse format: parameter_name = value1,value2,value3
                        if '=' in line:
                            param_name, values_str = line.split('=', 1)
                            param_name = param_name.strip()
                            values_str = values_str.strip()
                            
                            # Parse comma-separated values
                            if values_str:
                                try:
                                    values = [float(v.strip()) for v in values_str.split(',') if v.strip()]
                                    if values:
                                        self.loaded_presets[param_name] = sorted(values)  # Sort for consistency
                                        param_count += 1
                                except ValueError:
                                    continue  # Skip invalid lines
            
            self.preset_file_path = file_path
            self.preset_file_label.config(text=f"📁 {os.path.basename(file_path)}", foreground='#28a745')
            self.update_preset_tab()
            
            if self.loaded_presets:
                # Show success message with details
                total_values = sum(len(values) for values in self.loaded_presets.values())
                success_message = (
                    f"✅ Preset file loaded successfully!\n\n"
                    f"📁 File: {os.path.basename(file_path)}\n"
                    f"📊 Parameters: {param_count}\n"
                    f"🎯 Total preset values: {total_values}"
                )
                messagebox.showinfo("Success", success_message)
            else:
                messagebox.showwarning("No Presets", "⚠️ No valid preset data found in the file.")
                
        except Exception as e:
            messagebox.showerror("Error", f"❌ Failed to load preset file:\n{e}")
            self.preset_file_path = None
            self.loaded_presets = {}
            self.preset_file_label.config(text="❌ Failed to load preset file", foreground='#ff6b6b')
    
    def update_preset_tab(self):
        # Clear existing preset entries
        for widget in self.preset_scrollable.winfo_children():
            widget.destroy()
        self.preset_entries = {}
        
        if not self.params:
            # No parameters message with icon
            no_params_frame = ttk.Frame(self.preset_scrollable)
            no_params_frame.pack(expand=True, fill='both')
            
            ttk.Label(no_params_frame, text="📄", font=('Arial', 24)).pack(pady=(40, 10))
            ttk.Label(no_params_frame, text="No parameters loaded", 
                     font=('Arial', 12, 'bold'), foreground='#666666').pack()
            ttk.Label(no_params_frame, text="Please load a parameter file first in the Parameter Editor tab", 
                     font=('Arial', 9), foreground='#999999').pack(pady=(5, 0))
            return
        
        # Create preset entries for each parameter
        for i, (name, value) in enumerate(self.params):
            # Determine if this parameter has existing presets
            has_presets = name in self.loaded_presets
            
            # Parameter frame with conditional styling
            if has_presets:
                frame_text = f"✅ {name} (current: {value}mm)"
                frame_color = '#e8f5e8'  # Light green background
            else:
                frame_text = f"⚪ {name} (current: {value}mm)"
                frame_color = '#f8f9fa'  # Light gray background
            
            param_frame = ttk.LabelFrame(self.preset_scrollable, text=frame_text, padding=12)
            param_frame.pack(fill='x', padx=5, pady=8)
            
            # Checkbox with better styling
            checkbox_frame = ttk.Frame(param_frame)
            checkbox_frame.pack(fill='x', pady=(0, 8))
            
            var = tk.BooleanVar()
            checkbox = ttk.Checkbutton(checkbox_frame, text="✓ Include in presets", variable=var)
            checkbox.pack(side='left')
            
            if has_presets:
                status_label = ttk.Label(checkbox_frame, text="(has existing presets)", 
                                       foreground='#28a745', font=('Arial', 8))
                status_label.pack(side='right')
            
            # Instructions with example
            instruction_frame = ttk.Frame(param_frame)
            instruction_frame.pack(fill='x', pady=(0, 5))
            
            ttk.Label(instruction_frame, text="Preset values:", 
                     font=('Arial', 9, 'bold')).pack(side='left')
            ttk.Label(instruction_frame, text="Enter comma-separated values (e.g., 200,300,400,500)", 
                     font=('Arial', 8), foreground='#666666').pack(side='right')
            
            # Text entry with better styling
            preset_entry = tk.Text(param_frame, height=2, width=50, font=('Consolas', 9),
                                  relief='solid', borderwidth=1)
            preset_entry.pack(fill='x', pady=(0, 5))
            
            # Populate with existing preset values if available
            if has_presets:
                var.set(True)  # Enable checkbox if presets exist
                values = self.loaded_presets[name]
                # Convert values to clean format and join with commas
                clean_values = []
                for val in values:
                    if val.is_integer():
                        clean_values.append(str(int(val)))
                    else:
                        clean_values.append(str(val))
                preset_entry.insert("1.0", ",".join(clean_values))
                preset_entry.configure(bg='#f0fff0')  # Light green background for existing presets
            else:
                preset_entry.configure(bg='#ffffff')  # White background for new entries
            
            # Add placeholder text for empty entries
            if not has_presets:
                preset_entry.insert("1.0", "e.g., 200,300,400,500,600")
                preset_entry.configure(foreground='#cccccc')
                
                def on_focus_in(event, entry=preset_entry):
                    if entry.get("1.0", tk.END).strip() == "e.g., 200,300,400,500,600":
                        entry.delete("1.0", tk.END)
                        entry.configure(foreground='#000000')
                
                def on_focus_out(event, entry=preset_entry):
                    if not entry.get("1.0", tk.END).strip():
                        entry.insert("1.0", "e.g., 200,300,400,500,600")
                        entry.configure(foreground='#cccccc')
                
                preset_entry.bind("<FocusIn>", on_focus_in)
                preset_entry.bind("<FocusOut>", on_focus_out)
            
            self.preset_entries[name] = {
                'enabled': var,
                'values': preset_entry,
                'current': value
            }
    
    def generate_presets(self):
        if not self.params:
            messagebox.showwarning("No Parameters", "Please load a parameter file first.")
            return
        
        # Collect enabled parameters and their preset values
        preset_data = {}
        validation_errors = []
        
        for name, data in self.preset_entries.items():
            if data['enabled'].get():
                values_text = data['values'].get("1.0", tk.END).strip()
                
                # Skip placeholder text
                if values_text == "e.g., 200,300,400,500,600":
                    validation_errors.append(f"❌ {name}: Please replace the example text with actual values")
                    continue
                    
                if values_text:
                    try:
                        # Parse comma-separated values
                        values = [float(v.strip()) for v in values_text.split(',') if v.strip()]
                        if values:
                            # Sort values for better organization
                            values.sort()
                            preset_data[name] = values
                        else:
                            validation_errors.append(f"❌ {name}: No valid values found")
                    except ValueError:
                        validation_errors.append(f"❌ {name}: Invalid number format. Use comma-separated numbers only")
        
        # Show validation errors if any
        if validation_errors:
            error_message = "Please fix the following issues:\n\n" + "\n".join(validation_errors)
            messagebox.showerror("Validation Error", error_message)
            return
        
        if not preset_data:
            messagebox.showwarning("No Data", "⚠️ No parameters selected or no valid values provided.\n\nPlease:\n1. Check at least one parameter\n2. Enter valid preset values")
            return
        
        # Generate preset file
        try:
            # Use the currently loaded preset file path, or create based on main file name
            if self.preset_file_path:
                preset_file_path = self.preset_file_path
                action = "updated"
                action_icon = "🔄"
            else:
                # Create preset file name based on main file
                if self.file_path:
                    base_name = os.path.splitext(os.path.basename(self.file_path))[0]
                    preset_filename = f"{base_name}_presets.txt"
                    preset_file_path = os.path.join(get_app_dir(), preset_filename)
                else:
                    preset_file_path = os.path.join(get_app_dir(), "presets.txt")
                action = "generated"
                action_icon = "✨"
            
            with open(preset_file_path, 'w', encoding='utf-8-sig') as f:
                f.write("# ═══════════════════════════════════════════════════════════\n")
                f.write("# 📋 PARAMETER PRESET VALUES\n")
                f.write("# ═══════════════════════════════════════════════════════════\n")
                f.write("# Generated from: " + (os.path.basename(self.file_path) if self.file_path else "Unknown") + "\n")
                f.write("# Format: parameter_name = value1,value2,value3\n")
                f.write("# Use these presets in Simple Mode for quick parameter selection\n")
                f.write("# ═══════════════════════════════════════════════════════════\n\n")
                
                # Write each parameter with its preset values
                for param_name, preset_values in preset_data.items():
                    # Convert values to clean format (remove .0 from integers)
                    clean_values = []
                    for value in preset_values:
                        if value.is_integer():
                            clean_values.append(str(int(value)))
                        else:
                            clean_values.append(str(value))
                    
                    # Write parameter with comma-separated values
                    values_string = ",".join(clean_values)
                    f.write(f"{param_name} = {values_string}\n")
                
                f.write(f"\n# ═══════════════════════════════════════════════════════════\n")
                f.write(f"# Total parameters with presets: {len(preset_data)}\n")
                f.write(f"# Total preset values: {sum(len(values) for values in preset_data.values())}\n")
                f.write("# ═══════════════════════════════════════════════════════════")
            
            # Create detailed success message with visual feedback
            param_count = len(preset_data)
            total_values = sum(len(values) for values in preset_data.values())
            
            success_message = (
                f"{action_icon} Preset file {action} successfully!\n\n"
                f"📁 File: {os.path.basename(preset_file_path)}\n"
                f"📊 Parameters: {param_count}\n"
                f"🎯 Total preset values: {total_values}\n\n"
                f"💡 Tip: Use 'Simple Mode' in Parameter Editor to access these presets with sliders!"
            )
            
            messagebox.showinfo("Success", success_message)
            
            # Update the loaded presets and file path
            self.preset_file_path = preset_file_path
            self.load_preset_file(preset_file_path)
            
            # Update the file path display
            if hasattr(self, 'preset_file_label'):
                self.preset_file_label.config(text=os.path.basename(preset_file_path), foreground='#28a745')
            
        except Exception as e:
            messagebox.showerror("Error", f"❌ Failed to generate preset file:\n{e}")
    
    def save(self):
        try:
            updated = []
            for name, entry in self.entries.items():
                value = float(entry.get())
                if value <= 0:
                    raise ValueError(f"{name} must be > 0")
                updated.append((name, value))
            
            save_data(updated, self.original_lines, self.file_path)
            messagebox.showinfo("Success", "Parameters saved!")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    App().run()
    """
    Writes updated parameters back to the file. Keeps original order where lines were valid.
    Unknown lines (if any) are preserved at the end.
    """
    # Build a map for quick lookup
 