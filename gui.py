#!/usr/bin/env python
# gui.py

import os
import tkinter as tk
from tkinter import ttk
import tkinter.filedialog as fd
import tkinter.messagebox as mb

from config_module import load_config, save_config
from automation import run_automation

def run_gui():
    root = tk.Tk()
    root.title("InDesign Automation")
    
    # Apply theme
    style = ttk.Style()
    style.theme_use('clam')  # You can try other themes like 'alt', 'default', 'classic'
    
    # Configure custom styles
    style.configure('Action.TButton',
                   padding=10,
                   font=('Helvetica', 10, 'bold'))
    
    style.configure('Header.TLabel',
                   font=('Helvetica', 11, 'bold'),
                   padding=5)
    
    style.configure('Countdown.TLabel',
                   font=('Helvetica', 10),
                   foreground='red',
                   padding=5)

    # ----- Variables -----
    indesign_path_var = tk.StringVar()
    project_dir_var = tk.StringVar()
    new_project_var = tk.BooleanVar(value=True)
    use_existing_coords_var = tk.BooleanVar(value=False)

    # Load global config
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
            root.after(1000, countdown, remaining-1, local_config)
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

        # Save config
        global_config["indesign_exe"] = indesign_path
        global_config["last_project_dir"] = project_dir
        save_config(global_config, global_config_path)

        # Build or load local config
        local_config_path = os.path.join(project_dir, "config.json")
        local_config = load_config(local_config_path)

        if is_new_project:
            files_in_dir = os.listdir(project_dir)
            if not any(f.lower().endswith(".jpg") for f in files_in_dir):
                mb.showerror("Error", "No .jpg image found in project folder.")
                return
            
            credits_file = local_config.get("credits_file", "Credits.txt")
            print(credits_file)
            if not os.path.isfile(os.path.join(project_dir, credits_file)):
                mb.showerror("Error", f"{credits_file} not found.")
                return
            
            template_file = local_config.get("template_file", "template.indd")
            if not os.path.isfile(os.path.join(project_dir, template_file)):
                mb.showerror("Error", f"Template file not found: {template_file}")
                return

            local_config.update({
                "project_dir": project_dir,
                "template_file": template_file,
                "credits_file": credits_file,
                "credits_font": "Blackadder ITC\tRegular",
                "credits_font_size": 24,
                "text_box_position": "bottom_center"
            })

            if not use_coords:
                for key in ["text_frame_top_left_ratio", "text_frame_bottom_right_ratio",
                          "text_frame_top_left", "text_frame_bottom_right"]:
                    local_config.pop(key, None)
        else:
            local_config["project_dir"] = project_dir
            local_config.setdefault("template_file", "template.indd")
            local_config.setdefault("credits_file", "Credits.txt")

            if not use_coords:
                for key in ["text_frame_top_left_ratio", "text_frame_bottom_right_ratio",
                          "text_frame_top_left", "text_frame_bottom_right"]:
                    local_config.pop(key, None)

        save_config(local_config, local_config_path)

        # Begin countdown
        btn_run.state(['disabled'])
        countdown_label.config(text="Starting in 5 seconds...")
        root.after(1000, countdown, 4, local_config)

    # Create main container with padding
    main_frame = ttk.Frame(root, padding="20 20 20 20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Path Selection Section
    path_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10 10 10 10")
    path_frame.pack(fill=tk.X, pady=(0, 10))

    # InDesign Path
    ttk.Label(path_frame, text="Adobe InDesign.exe Path:").pack(anchor=tk.W)
    path_entry_frame = ttk.Frame(path_frame)
    path_entry_frame.pack(fill=tk.X, pady=(5, 10))
    ttk.Entry(path_entry_frame, textvariable=indesign_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    ttk.Button(path_entry_frame, text="Browse...", command=browse_indesign).pack(side=tk.RIGHT)

    # Project Directory
    ttk.Label(path_frame, text="Project Folder:").pack(anchor=tk.W)
    dir_entry_frame = ttk.Frame(path_frame)
    dir_entry_frame.pack(fill=tk.X, pady=5)
    ttk.Entry(dir_entry_frame, textvariable=project_dir_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    ttk.Button(dir_entry_frame, text="Browse...", command=browse_project).pack(side=tk.RIGHT)

    # Project Options Section
    options_frame = ttk.LabelFrame(main_frame, text="Project Options", padding="10 10 10 10")
    options_frame.pack(fill=tk.X, pady=(0, 10))

    ttk.Radiobutton(options_frame, text="Create New Project", variable=new_project_var, value=True).pack(anchor=tk.W, pady=2)
    ttk.Radiobutton(options_frame, text="Use Existing Project", variable=new_project_var, value=False).pack(anchor=tk.W, pady=2)
    ttk.Checkbutton(options_frame, text="Use Existing Coordinates (if found)", variable=use_existing_coords_var).pack(anchor=tk.W, pady=(5, 0))

    # Action Section
    action_frame = ttk.Frame(main_frame)
    action_frame.pack(fill=tk.X, pady=(10, 0))

    btn_run = ttk.Button(action_frame, text="Run Automation", command=on_run, style='Action.TButton')
    btn_run.pack(pady=10)

    countdown_label = ttk.Label(action_frame, text="", style='Countdown.TLabel')
    countdown_label.pack(pady=5)

    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()

if __name__ == "__main__":
    run_gui()