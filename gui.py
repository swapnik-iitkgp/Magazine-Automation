#!/usr/bin/env python
# gui.py

import os
import tkinter as tk
from tkinter import ttk
import tkinter.filedialog as fd
import tkinter.messagebox as mb

# Import the sv_ttk module for a modern dark theme.
import sv_ttk

from config_module import load_config, save_config
from automation import run_automation

def find_file_recursive(directory, filename):
    """
    Recursively search for a file within the given directory.
    Returns the full path to the file if found; otherwise, returns None.
    """
    for root_dir, _, files in os.walk(directory):
        if filename in files:
            return os.path.join(root_dir, filename)
    return None

def run_gui():
    root = tk.Tk()
    root.title("InDesign Automation")
    
    # Set a fixed default window size and make it non-resizable.
    default_width = 1000
    default_height = 700
    root.geometry(f"{default_width}x{default_height}")
    root.resizable(False, False)
    root.configure(background="#e6e6e6")
    # Use curly braces around the font name because it contains spaces.
    root.option_add("*Font", "{Segoe UI} 10")
    
    # Apply the sv_ttk dark theme for a modern look.
    sv_ttk.set_theme("dark")
    
    # Additional style modifications to complement the dark theme.
    style = ttk.Style()
    style.configure('TFrame', background="#2e2e2e")
    style.configure('TLabel', background="#2e2e2e", foreground="#ffffff")
    style.configure('TLabelFrame', background="#3e3e3e", foreground="#ffffff",
                    font=('Segoe UI', 10, 'bold'), relief="groove", borderwidth=2)
    style.configure('TButton', font=('Segoe UI', 10, 'bold'))
    style.configure('Action.TButton', padding=10, font=('Segoe UI', 10, 'bold'))
    style.configure('Countdown.TLabel', font=('Segoe UI', 10),
                    foreground='red', background="#2e2e2e", padding=5)
    
    # Create a container frame with a canvas and scrollbar for scrolling content.
    container = ttk.Frame(root)
    container.pack(fill="both", expand=True)
    
    canvas = tk.Canvas(container, background="#2e2e2e", highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)
    
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Create a frame inside the canvas which will hold all the widgets.
    # Note: We remove extra padding to help eliminate empty space.
    scrollable_frame = ttk.Frame(canvas)
    # Capture the window id so we can update its width later.
    window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    
    # Update the scroll region and force the scrollable frame to fill the canvas width.
    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    scrollable_frame.bind("<Configure>", on_configure)
    
    def on_canvas_configure(event):
        # Set the width of the inner frame to match the canvas width.
        canvas.itemconfig(window_id, width=event.width)
    canvas.bind("<Configure>", on_canvas_configure)
    
    # Bind mouse wheel scrolling to the canvas.
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    # ----- Variables -----
    indesign_path_var = tk.StringVar()
    project_dir_var = tk.StringVar()
    new_project_var = tk.BooleanVar(value=True)
    use_existing_coords_var = tk.BooleanVar(value=False)

    # New project-specific basic settings.
    target_page_var = tk.IntVar(value=8)
    split_page_var = tk.IntVar(value=8)

    # New project advanced options variables.
    credits_font_size_var = tk.IntVar(value=24)
    single_prob_var = tk.DoubleVar(value=0.3)
    double_prob_var = tk.DoubleVar(value=0.4)
    four_prob_var = tk.DoubleVar(value=0.3)

    # Toggle options stored in dictionaries mapping option -> BooleanVar.
    credits_font_vars = {
        "Blackadder ITC\tRegular": tk.BooleanVar(value=True),
        "Arial\tRegular": tk.BooleanVar(value=True)
    }
    credits_colors_vars = {
        "Black": tk.BooleanVar(value=True),
        "Red": tk.BooleanVar(value=True),
        "Blue": tk.BooleanVar(value=True),
        "Green": tk.BooleanVar(value=True),
        "Yellow": tk.BooleanVar(value=True)
    }
    text_box_position_vars = {
        "bottom_center": tk.BooleanVar(value=True),
        "top_center": tk.BooleanVar(value=False),
        "center": tk.BooleanVar(value=False)
    }
    layouts_vars = {
        "single": tk.BooleanVar(value=True),
        "double": tk.BooleanVar(value=True),
        "four": tk.BooleanVar(value=True)
    }

    # Load global config (for the InDesign executable).
    global_config_path = "global_config.json"
    global_config = load_config(global_config_path)
    if "indesign_exe" in global_config:
        indesign_path_var.set(global_config["indesign_exe"])
    if "last_project_dir" in global_config and os.path.isdir(global_config["last_project_dir"]):
        project_dir_var.set(global_config["last_project_dir"])

    def browse_indesign():
        path = fd.askopenfilename(
            title="Select Adobe InDesign.exe",
            filetypes=[("Executable", "*.exe")]
        )
        if path:
            indesign_path_var.set(path)

    def browse_project():
        directory = fd.askdirectory(title="Select Project Folder")
        if directory:
            project_dir_var.set(directory)

    def countdown(remaining, local_config):
        if remaining > 0:
            countdown_label.config(text=f"Starting in {remaining} seconds...")
            root.after(1000, countdown, remaining - 1, local_config)
        else:
            countdown_label.config(text="Starting now!")
            root.after(500, begin_automation, local_config)

    def begin_automation(local_config):
        root.withdraw()
        try:
            run_automation(local_config)
        except Exception as ex:
            mb.showerror("Automation Error", str(ex))
        finally:
            root.deiconify()
            btn_run.state(['!disabled'])
            countdown_label.config(text="")

    def on_run():
        indesign_path = indesign_path_var.get().strip()
        project_dir = project_dir_var.get().strip()
        is_new_project = new_project_var.get()
        use_coords = use_existing_coords_var.get()

        if not indesign_path or not os.path.isfile(indesign_path):
            mb.showerror("Error", "Please specify a valid InDesign.exe path.")
            return
        if not project_dir or not os.path.isdir(project_dir):
            mb.showerror("Error", "Please select a valid project folder.")
            return

        # Save the global config.
        global_config["indesign_exe"] = indesign_path
        global_config["last_project_dir"] = project_dir
        save_config(global_config, global_config_path)

        # Build or load the project (local) config from project_dir/config.json.
        local_config_path = os.path.join(project_dir, "config.json")
        local_config = load_config(local_config_path)

        if is_new_project:
            # Basic check: at least one .jpg image should exist (recursively).
            jpg_found = False
            for root_dir, _, files in os.walk(project_dir):
                if any(f.lower().endswith(".jpg") for f in files):
                    jpg_found = True
                    break
            if not jpg_found:
                mb.showerror("Error", "No .jpg images found in project folder.")
                return

            # Recursively look for the Credits.txt file.
            credits_file = local_config.get("credits_file", "Credits.txt")
            credits_path = find_file_recursive(project_dir, credits_file)
            if not credits_path:
                mb.showerror("Error", f"{credits_file} not found in project folder.")
                return

            # Recursively look for the template.indd file.
            template_file = local_config.get("template_file", "template.indd")
            template_path = find_file_recursive(project_dir, template_file)
            if not template_path:
                mb.showerror("Error", f"Template file not found: {template_file}")
                return

            # --- Gather Advanced Options from the toggle checkbuttons ---
            selected_credits_font = [option for option, var in credits_font_vars.items() if var.get()]
            selected_credits_colors = [option for option, var in credits_colors_vars.items() if var.get()]
            selected_text_box_positions = [option for option, var in text_box_position_vars.items() if var.get()]
            selected_layouts = [option for option, var in layouts_vars.items() if var.get()]

            layout_probabilities = {}
            if "single" in selected_layouts:
                layout_probabilities["single"] = single_prob_var.get()
            if "double" in selected_layouts:
                layout_probabilities["double"] = double_prob_var.get()
            if "four" in selected_layouts:
                layout_probabilities["four"] = four_prob_var.get()

            local_config.update({
                "project_dir": project_dir,
                "template_file": template_file,
                "credits_file": credits_file,
                "target_page": target_page_var.get(),
                "split_page": split_page_var.get(),
                "credits_font": selected_credits_font,
                "credits_colors": selected_credits_colors,
                "credits_font_size": credits_font_size_var.get(),
                "text_box_position": selected_text_box_positions,
                "layouts": selected_layouts,
                "layout_probabilities": layout_probabilities,
            })

            if not use_coords:
                for key in ["text_frame_top_left_ratio", "text_frame_bottom_right_ratio",
                            "text_frame_top_left", "text_frame_bottom_right"]:
                    local_config.pop(key, None)
        else:
            local_config["project_dir"] = project_dir
            local_config.setdefault("template_file", "template.indd")
            local_config.setdefault("credits_file", "Credits.txt")
            local_config.setdefault("credits_font", ["Blackadder ITC\tRegular", "Arial\tRegular"])
            local_config.setdefault("credits_font_size", 24)
            local_config.setdefault("text_box_position", ["bottom_center"])
            local_config.setdefault("layouts", ["single", "double", "four"])
            if not use_coords:
                for key in ["text_frame_top_left_ratio", "text_frame_bottom_right_ratio",
                            "text_frame_top_left", "text_frame_bottom_right"]:
                    local_config.pop(key, None)

        save_config(local_config, local_config_path)
        btn_run.state(['disabled'])
        countdown_label.config(text="Starting in 5 seconds...")
        root.after(1000, countdown, 4, local_config)

    # ------------- Build Main GUI Content -------------
    # All widgets are added to scrollable_frame.
    # Basic Configuration
    path_frame = ttk.LabelFrame(scrollable_frame, text="Configuration", padding="10 10 10 10")
    path_frame.pack(fill=tk.X, pady=(0, 10))
    ttk.Label(path_frame, text="Adobe InDesign.exe Path:").pack(anchor=tk.W)
    path_entry_frame = ttk.Frame(path_frame)
    path_entry_frame.pack(fill=tk.X, pady=(5, 10))
    ttk.Entry(path_entry_frame, textvariable=indesign_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    ttk.Button(path_entry_frame, text="Browse...", command=browse_indesign).pack(side=tk.RIGHT)
    ttk.Label(path_frame, text="Project Folder:").pack(anchor=tk.W)
    dir_entry_frame = ttk.Frame(path_frame)
    dir_entry_frame.pack(fill=tk.X, pady=5)
    ttk.Entry(dir_entry_frame, textvariable=project_dir_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    ttk.Button(dir_entry_frame, text="Browse...", command=browse_project).pack(side=tk.RIGHT)
    
    # Project Options
    options_frame = ttk.LabelFrame(scrollable_frame, text="Project Options", padding="10 10 10 10")
    options_frame.pack(fill=tk.X, pady=(0, 10))
    ttk.Radiobutton(options_frame, text="Create New Project", variable=new_project_var, value=True).pack(anchor=tk.W, pady=2)
    ttk.Radiobutton(options_frame, text="Use Existing Project", variable=new_project_var, value=False).pack(anchor=tk.W, pady=2)
    ttk.Checkbutton(options_frame, text="Use Existing Coordinates (if found)", variable=use_existing_coords_var).pack(anchor=tk.W, pady=(5, 0))
    
    # New Project Basic Settings
    new_project_frame = ttk.LabelFrame(scrollable_frame, text="New Project Settings", padding="10 10 10 10")
    new_project_frame.pack(fill=tk.X, pady=(0, 10))
    ttk.Label(new_project_frame, text="Target Page:").grid(row=0, column=0, sticky=tk.W, pady=2)
    target_page_entry = ttk.Entry(new_project_frame, textvariable=target_page_var, width=10)
    target_page_entry.grid(row=0, column=1, sticky=tk.W, pady=2, padx=(5,0))
    ttk.Label(new_project_frame, text="Split Page:").grid(row=1, column=0, sticky=tk.W, pady=2)
    split_page_entry = ttk.Entry(new_project_frame, textvariable=split_page_var, width=10)
    split_page_entry.grid(row=1, column=1, sticky=tk.W, pady=2, padx=(5,0))
    
    # Advanced New Project Options
    advanced_frame = ttk.LabelFrame(scrollable_frame, text="Advanced New Project Options", padding="10 10 10 10")
    advanced_frame.pack(fill=tk.X, pady=(0, 10))
    
    # Credits Font toggles
    ttk.Label(advanced_frame, text="Credits Font:").grid(row=0, column=0, sticky=tk.W, pady=2)
    credits_font_frame = ttk.Frame(advanced_frame)
    credits_font_frame.grid(row=0, column=1, sticky=tk.W, pady=2, padx=(5,0))
    for option, var in credits_font_vars.items():
        cb = ttk.Checkbutton(credits_font_frame, text=option, variable=var)
        cb.pack(side=tk.LEFT, padx=5)
    
    # Credits Colors toggles
    ttk.Label(advanced_frame, text="Credits Colors:").grid(row=1, column=0, sticky=tk.W, pady=2)
    credits_colors_frame = ttk.Frame(advanced_frame)
    credits_colors_frame.grid(row=1, column=1, sticky=tk.W, pady=2, padx=(5,0))
    for option, var in credits_colors_vars.items():
        cb = ttk.Checkbutton(credits_colors_frame, text=option, variable=var)
        cb.pack(side=tk.LEFT, padx=5)
    
    # Credits Font Size
    ttk.Label(advanced_frame, text="Credits Font Size:").grid(row=2, column=0, sticky=tk.W, pady=2)
    credits_font_size_entry = ttk.Entry(advanced_frame, textvariable=credits_font_size_var, width=10)
    credits_font_size_entry.grid(row=2, column=1, sticky=tk.W, pady=2, padx=(5,0))
    
    # Text Box Position toggles
    ttk.Label(advanced_frame, text="Text Box Position:").grid(row=3, column=0, sticky=tk.W, pady=2)
    text_box_position_frame = ttk.Frame(advanced_frame)
    text_box_position_frame.grid(row=3, column=1, sticky=tk.W, pady=2, padx=(5,0))
    for option, var in text_box_position_vars.items():
        cb = ttk.Checkbutton(text_box_position_frame, text=option, variable=var)
        cb.pack(side=tk.LEFT, padx=5)
    
    # Layouts toggles
    ttk.Label(advanced_frame, text="Layouts:").grid(row=4, column=0, sticky=tk.W, pady=2)
    layouts_frame = ttk.Frame(advanced_frame)
    layouts_frame.grid(row=4, column=1, sticky=tk.W, pady=2, padx=(5,0))
    for option, var in layouts_vars.items():
        cb = ttk.Checkbutton(layouts_frame, text=option, variable=var)
        cb.pack(side=tk.LEFT, padx=5)
    
    # Layout Probabilities
    prob_frame = ttk.LabelFrame(advanced_frame, text="Layout Probabilities", padding="5 5 5 5")
    prob_frame.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=5)
    ttk.Label(prob_frame, text="Single:").grid(row=0, column=0, sticky=tk.W, pady=2)
    single_prob_entry = ttk.Entry(prob_frame, textvariable=single_prob_var, width=10)
    single_prob_entry.grid(row=0, column=1, sticky=tk.W, pady=2, padx=(5,0))
    ttk.Label(prob_frame, text="Double:").grid(row=1, column=0, sticky=tk.W, pady=2)
    double_prob_entry = ttk.Entry(prob_frame, textvariable=double_prob_var, width=10)
    double_prob_entry.grid(row=1, column=1, sticky=tk.W, pady=2, padx=(5,0))
    ttk.Label(prob_frame, text="Four:").grid(row=2, column=0, sticky=tk.W, pady=2)
    four_prob_entry = ttk.Entry(prob_frame, textvariable=four_prob_var, width=10)
    four_prob_entry.grid(row=2, column=1, sticky=tk.W, pady=2, padx=(5,0))
    
    def update_new_project_settings_state(*args):
        state = 'normal' if new_project_var.get() else 'disabled'
        for widget in [target_page_entry, split_page_entry,
                       credits_font_size_entry, single_prob_entry, double_prob_entry, four_prob_entry]:
            try:
                widget.config(state=state)
            except tk.TclError:
                pass
        # Update toggle groups.
        for child in credits_font_frame.winfo_children():
            child.config(state=state)
        for child in credits_colors_frame.winfo_children():
            child.config(state=state)
        for child in text_box_position_frame.winfo_children():
            child.config(state=state)
        for child in layouts_frame.winfo_children():
            child.config(state=state)
    
    new_project_var.trace_add('write', update_new_project_settings_state)
    update_new_project_settings_state()
    
    # Action Buttons
    action_frame = ttk.Frame(scrollable_frame)
    action_frame.pack(fill=tk.X, pady=(10, 0))
    btn_run = ttk.Button(action_frame, text="Run Automation", command=on_run, style='Action.TButton')
    btn_run.pack(pady=10)
    countdown_label = ttk.Label(action_frame, text="", style='Countdown.TLabel')
    countdown_label.pack(pady=5)
    
    # Center the window on the screen.
    root.update_idletasks()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (default_width // 2)
    y = (screen_height // 2) - (default_height // 2)
    root.geometry(f"{default_width}x{default_height}+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    run_gui()
