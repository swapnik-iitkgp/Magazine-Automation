#!/usr/bin/env python
# automation.py

import os
import time
import random
import pyautogui
import cv2
import numpy as np
import win32com.client
import pywintypes


pyautogui.FAILSAFE = False


from get_split import split_template
from merge_indd import merge_indd_files
from config_module import load_config, save_config

# Global list to store created text frame COM objects (if you wish to adjust their line spacing later).
created_text_frames = []

########################################
# OPENCV REGION SELECTION (FOR TEXT BOX)
########################################

_selected_points = []

def _click_event(event, x, y, flags, param):
    global _selected_points
    image = param
    if event == cv2.EVENT_LBUTTONDOWN:
        _selected_points.append((x, y))
        cv2.circle(image, (x, y), 5, (0, 255, 0), -1)
        cv2.imshow("Screenshot", image)

def get_region_from_opencv():
    """
    Take a screenshot and let the user click the TOP-LEFT and BOTTOM-RIGHT
    corners for the text region. Return the two corner ratios.
    """
    global _selected_points
    _selected_points = []

    screenshot = pyautogui.screenshot()
    image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    clone = image.copy()
    screen_h, screen_w = clone.shape[:2]

    cv2.namedWindow("Screenshot", cv2.WINDOW_NORMAL)
    cv2.imshow("Screenshot", clone)
    cv2.setMouseCallback("Screenshot", _click_event, clone)

    print("[INFO] Click the TOP-LEFT corner, then the BOTTOM-RIGHT corner for the text region.")
    while True:
        cv2.imshow("Screenshot", clone)
        if len(_selected_points) >= 2:
            break
        cv2.waitKey(1)
    cv2.destroyAllWindows()

    x1, y1 = _selected_points[0]
    x2, y2 = _selected_points[1]
    left   = min(x1, x2)
    right  = max(x1, x2)
    top    = min(y1, y2)
    bottom = max(y1, y2)

    ratio_left   = left / screen_w
    ratio_top    = top / screen_h
    ratio_right  = right / screen_w
    ratio_bottom = bottom / screen_h

    print(f"[INFO] Selected region ratios: L={ratio_left:.3f}, T={ratio_top:.3f}, R={ratio_right:.3f}, B={ratio_bottom:.3f}")
    return (ratio_left, ratio_top), (ratio_right, ratio_bottom)

def retrieve_ratio_region(config):
    """
    If ratio-based corners exist in the config, convert them to pixel coordinates.
    """
    if "text_frame_top_left_ratio" in config and "text_frame_bottom_right_ratio" in config:
        ratio_left, ratio_top = config["text_frame_top_left_ratio"]
        ratio_right, ratio_bottom = config["text_frame_bottom_right_ratio"]

        screenshot = pyautogui.screenshot()
        screen_w, screen_h = screenshot.size  # (width, height)

        left   = int(ratio_left * screen_w)
        top    = int(ratio_top * screen_h)
        right  = int(ratio_right * screen_w)
        bottom = int(ratio_bottom * screen_h)
        return (left, top), (right, bottom)
    else:
        return None

########################################
# PAGE SELECTION: GET OR INSERT EMPTY PAGE
########################################

def get_empty_page(doc, target_page=None):
    """
    If a target_page is provided and that page exists and is empty (has no PageItems),
    return that page. Otherwise, scan the document for the first empty page.
    """
    if target_page is not None and target_page <= doc.Pages.Count:
        page = doc.Pages.Item(target_page)
        if page.PageItems.Count == 0:
            return page
    for i in range(1, doc.Pages.Count + 1):
        page = doc.Pages.Item(i)
        if page.PageItems.Count == 0:
            return page
    return None

########################################
# LAYOUT SELECTION FUNCTION
########################################

def choose_layout(num_remaining, config):
    """
    Choose a layout mode based on the number of images remaining and the probabilities
    specified in the config.
    """
    layout_probs = config.get("layout_probabilities", {"single": 0.33, "double": 0.33, "four": 0.34})
    if num_remaining == 1:
        return "single"
    elif num_remaining == 2:
        return "double"
    elif num_remaining == 3:
        population = ["single", "double"]
        weights = [layout_probs.get("single", 0.5), layout_probs.get("double", 0.5)]
        return random.choices(population, weights=weights)[0]
    else:
        population = ["single", "double", "four"]
        weights = [layout_probs.get("single", 0.33),
                   layout_probs.get("double", 0.33),
                   layout_probs.get("four", 0.34)]
        return random.choices(population, weights=weights)[0]

########################################
# TEXT BOX AND CREDITS INSERTION
########################################

def compute_text_box_coordinates(region_top_left, region_bottom_right, config, text_content=None):
    """
    Compute the coordinates for the text box that will contain the credits.
    The height (box_h) is calculated based on the number of lines in text_content.
    If text_content is not provided or is empty, a default of 3 lines is assumed.
    
    After the initial calculation, this function shifts the box vertically if needed
    so that the entire box remains within the defined region.
    """
    base_font_size = config.get("credits_font_size", 24)
    position  = config.get("text_box_position", ["bottom_center"])[0]

    region_width  = region_bottom_right[0] - region_top_left[0]
    region_height = region_bottom_right[1] - region_top_left[1]

    line_multiplier = 1.2
    if text_content:
        line_count = len(text_content.splitlines())
        if line_count == 0:
            line_count = 3
    else:
        line_count = 3

    # Calculate the desired box height based on the number of lines.
    box_h = int(line_count * base_font_size * line_multiplier)
    box_w = region_width

    # Determine the initial top-left y-coordinate based on the desired position.
    if position == "center":
        tl_y = region_top_left[1] + (region_height - box_h) // 2
    elif position == "top_center":
        tl_y = region_top_left[1]
    elif position == "bottom_center":
        tl_y = region_bottom_right[1] - box_h
    else:
        tl_y = region_top_left[1] + (region_height - box_h) // 2

    tl_x = region_top_left[0]

    # --- SHIFT THE BOX IF IT FALLS OUTSIDE THE REGION ---
    if box_h > region_height:
        box_h = region_height
        tl_y = region_top_left[1]
    else:
        if tl_y < region_top_left[1]:
            tl_y = region_top_left[1]
        if tl_y + box_h > region_bottom_right[1]:
            tl_y = region_bottom_right[1] - box_h

    box_tl = (tl_x, tl_y)
    box_br = (tl_x + box_w, tl_y + box_h)
    click_pt = ((box_tl[0] + box_br[0]) // 2, (box_tl[1] + box_br[1]) // 2)
    return box_tl, box_br, click_pt

def insert_text_frame_and_type(text_content, drag_start, drag_end, click_point, config, is_first_page=False):
    """
    Create a text frame via PyAutoGUI. Immediately after creating the frame and setting
    the insertion point, use COM to pre-apply the default font, size, text color and alignment,
    and then force the text frame’s geometric bounds (to match the computed coordinates).
    Finally, type the text.
    
    (For first-page credits, an optional bold formatting for the first paragraph is applied after typing.)
    """
    import random, time, pyautogui, win32com.client

    # Choose a random font from the list provided.
    fonts = config["credits_font"]
    base_font = random.choice(fonts)
    base_size = config["credits_font_size"]
    combined_text = text_content

    print("[INFO] Creating text frame via PyAutoGUI...")
    time.sleep(1)
    pyautogui.press('t')  # Select Type Tool
    time.sleep(0.5)

    # Drag to create the text frame.
    pyautogui.moveTo(drag_start[0], drag_start[1], duration=0.5)
    pyautogui.mouseDown()
    pyautogui.moveTo(drag_end[0], drag_end[1], duration=1)
    pyautogui.mouseUp()
    time.sleep(0.5)

    # Click inside the frame to set the insertion point.
    pyautogui.moveTo(click_point[0], click_point[1], duration=0.5)
    pyautogui.click()
    time.sleep(0.5)
    print("[INFO] Text frame created.")

    # --- SET TEXT COLOR AND ALIGNMENT AFTER CREATING THE TEXT FRAME ---
    try:
        indesign = win32com.client.Dispatch("InDesign.Application")
        doc = indesign.ActiveDocument
        if doc.Selection.Count > 0:
            textFrame = doc.Selection.Item(1)
            
            # Choose a random color from the config list.
            colors = config["credits_colors"]
            chosen_color = random.choice(colors)
            try:
                swatch = doc.Colors.Item(chosen_color)
            except Exception as e:
                print(f"[INFO] Swatch '{chosen_color}' not found. Creating it.")
                swatch = doc.Colors.Add()
                swatch.Name = chosen_color
                # Use CMYK values (adjust these values as needed)
                color_values = {
                    "red":    [0, 100, 100, 0],
                    "black":  [0, 0, 0, 100],
                    "blue":   [100, 75, 0, 0],
                    "green":  [75, 0, 100, 0],
                    "yellow": [0, 0, 100, 0]
                }
                chosen_lower = chosen_color.lower()
                if chosen_lower in color_values:
                    swatch.ColorValue = color_values[chosen_lower]
                else:
                    swatch.ColorValue = [0, 0, 0, 100]
            
            # Apply the chosen color to the text.
            textFrame.ParentStory.Texts.Item(1).FillColor = swatch

        else:
            print("[WARN] No text frame selected for color formatting.")
    except Exception as e:
        print("[ERROR] Setting text color and alignment failed:", e)

    # --- PRE-APPLY FONT & SIZE VIA COM BEFORE TYPING ---
    try:
        indesign = win32com.client.Dispatch("InDesign.Application")
        doc = indesign.ActiveDocument
        if doc.Selection.Count > 0:
            textFrame = doc.Selection.Item(1)
            textFrame.ParentStory.AppliedFont = base_font
            textFrame.ParentStory.PointSize = base_size
            print(f"[INFO] Pre-applied default formatting: {base_font} at size {base_size}")
            # --- FORCE THE TEXT FRAME SIZE VIA COM ---
            top = drag_start[1]
            left = drag_start[0]
            bottom = drag_end[1]
            right = drag_end[0]
            try:
                textFrame.GeometricBounds = [top, left, bottom, right]
                print(f"[INFO] Set text frame geometric bounds to: {[top, left, bottom, right]}")
            except Exception as e:
                print("[ERROR] Setting text frame geometric bounds via COM failed:", e)
        else:
            print("[WARN] No text frame selection found for pre-formatting.")
    except Exception as e:
        print("[ERROR] Pre-formatting via COM failed:", e)

    print("[INFO] Typing credits text...")
    pyautogui.typewrite(combined_text, interval=0.05)
    time.sleep(0.5)

    # Optionally, if this is the first page, make the first paragraph bold.
    try:
        if doc.Selection.Count > 0:
            textFrame = doc.Selection.Item(1)
            paras = textFrame.ParentStory.Paragraphs
            if paras.Count >= 1:
                main_para = paras.Item(1)
                main_font = base_font + " Bold"
                main_para.AppliedFont = main_font
                main_para.PointSize = base_size * 1.5
                print(f"[INFO] Applied bold formatting to first paragraph: {main_font} at size {base_size * 1.5}")
    except Exception as e:
        print("[ERROR] Bold formatting via COM failed:", e)

    print("[INFO] Finished typing credits.")

    # Save the text frame reference for later line spacing adjustment.
    try:
        indesign = win32com.client.Dispatch("InDesign.Application")
        doc = indesign.ActiveDocument
        if doc.Selection.Count > 0:
            tf = doc.Selection.Item(1)
        else:
            tf = doc.TextFrames.Item(doc.TextFrames.Count)
        created_text_frames.append(tf)
        print("[INFO] Saved text frame for later line spacing adjustment.")
    except Exception as e:
        print("[ERROR] Could not save text frame reference:", e)


########################################
# PROCESSING IMAGES PER MODEL FOLDER
########################################


import pyautogui
import time

def place_model_images(doc, model_folder, config, target_page=None):
    """
    Process images from a model folder and place them on pages.
    On the first page for a model folder, overlay the credits (if available).
    """
    image_files = sorted([f for f in os.listdir(model_folder) if f.lower().endswith(".jpg")])
    if not image_files:
        print(f"[WARN] No images found in {model_folder}.")
        return

    # Read credits from Credits.txt (if available).
    credits_path = os.path.join(model_folder, "Credits.txt")
    credits_text = ""
    if os.path.isfile(credits_path):
        with open(credits_path, "r", encoding="utf-8") as f:
            credits_text = f.read().strip()
    else:
        print(f"[WARN] No Credits.txt found in {model_folder}.")

    image_index = 0
    first_page_for_model = True

    while image_index < len(image_files):
        num_remaining = len(image_files) - image_index
        layout_mode = choose_layout(num_remaining, config)
        if layout_mode == "single":
            images_per_page = 1
        elif layout_mode == "double":
            images_per_page = 2
        elif layout_mode == "four":
            images_per_page = 4
        else:
            images_per_page = 1

        if first_page_for_model and target_page is not None:
            page = get_empty_page(doc, target_page)
            if page is None:
                page = doc.Pages.Add()
        else:
            page = get_empty_page(doc)
            if page is None:
                page = doc.Pages.Add()

        page_width = doc.DocumentPreferences.PageWidth
        page_height = doc.DocumentPreferences.PageHeight

        frames = []
        if images_per_page == 1:
            frames.append([0, 0, page_height, page_width])
        elif images_per_page == 2:
            frames.append([0, 0, page_height, page_width / 2])
            frames.append([0, page_width / 2, page_height, page_width])
        elif images_per_page == 4:
            frames.append([0, 0, page_height / 2, page_width / 2])
            frames.append([0, page_width / 2, page_height / 2, page_width])
            frames.append([page_height / 2, 0, page_height, page_width / 2])
            frames.append([page_height / 2, page_width / 2, page_height, page_width])

        for frame_bounds in frames:
            if image_index >= len(image_files):
                break
            image_path = os.path.join(model_folder, image_files[image_index])
            print(f"[INFO] Placing image: {image_path}")
            rect = page.Rectangles.Add()
            rect.GeometricBounds = frame_bounds
            try:
                rect.Place(image_path)
                # Removed the Fit() call so the image fills the frame:
                # rect.Graphics.Item(1).Fit()
            except Exception as e:
                print(f"[ERROR] Placing image {image_path}: {e}")
            image_index += 1

        # --- Before filling images on this page, click just outside the page region ---
        # Get page's top-left and bottom-right based on ratios in config
        region_top_left = config["text_frame_top_left_ratio"]
        region_bottom_right = config["text_frame_bottom_right_ratio"]

        # Convert ratios to actual pixel coordinates
        screenshot = pyautogui.screenshot()
        screen_w, screen_h = screenshot.size  # (width, height)

        left = int(region_top_left[0] * screen_w)
        top = int(region_top_left[1] * screen_h)
        right = int(region_bottom_right[0] * screen_w)
        bottom = int(region_bottom_right[1] * screen_h)

        # Adjust the click position to just outside the page boundary
        offset = 10  # You can adjust this offset as needed
        click_x = right + offset  # Click just to the right of the page
        click_y = top + offset  # Click just below the top-left corner of the page

        # Move the mouse to this adjusted position and click
        pyautogui.moveTo(click_x, click_y, duration=0.5)
        pyautogui.click()

        # --- Now press V to switch to the Selection Tool ---
        pyautogui.press('v')

        # Get the page selection region from config (if it exists)
        region = retrieve_ratio_region(config)
        if region is None:
            ratio_tl, ratio_br = get_region_from_opencv()
            config["text_frame_top_left_ratio"] = list(ratio_tl)
            config["text_frame_bottom_right_ratio"] = list(ratio_br)
            save_config(config, os.path.join(config["project_dir"], "config.json"))
            region = retrieve_ratio_region(config)
        (page_tl, page_br) = region

        # Adjust the corners by a few pixels (e.g., 10 pixels)
        selection_start = (page_tl[0] - offset, page_tl[1] - offset)
        selection_end = (page_br[0] + offset, page_br[1] + offset)

        # Drag to select all items on the page
        pyautogui.moveTo(selection_start[0], selection_start[1], duration=0.5)
        pyautogui.mouseDown()
        pyautogui.moveTo(selection_end[0], selection_end[1], duration=1)
        pyautogui.mouseUp()
        time.sleep(0.5)

        # Apply the fill command via the shortcut (ctrl+alt+shift+C)
        pyautogui.hotkey('ctrl', 'alt', 'shift', 'c')

        # --- Before adding text, click outside the page region again to deactivate any active frame ---
        pyautogui.moveTo(click_x, click_y)  # Click outside the page region
        pyautogui.click()

        # --- Now press T to switch to the Type Tool ---
        pyautogui.press('t')

        if first_page_for_model and credits_text:
            # Compute the text box coordinates based on the credits text.
            region = retrieve_ratio_region(config)
            if region is not None:
                region_top_left, region_bottom_right = region
            else:
                ratio_tl, ratio_br = get_region_from_opencv()
                config["text_frame_top_left_ratio"] = list(ratio_tl)
                config["text_frame_bottom_right_ratio"] = list(ratio_br)
                save_config(config, os.path.join(config["project_dir"], "config.json"))
                region = retrieve_ratio_region(config)
                region_top_left, region_bottom_right = region

            box_tl, box_br, click_pt = compute_text_box_coordinates(region_top_left, region_bottom_right, config, credits_text)
            insert_text_frame_and_type(credits_text, box_tl, box_br, click_pt, config, is_first_page=True)
            first_page_for_model = False



            

def cleanup_indd_files(project_dir, template_file, output_file):
    # Convert to lowercase for case-insensitive comparison.
    template_file_lower = template_file.lower()
    output_file_lower = output_file.lower()
    
    for filename in os.listdir(project_dir):
        if filename.lower().endswith(".indd") and filename.lower() not in [template_file_lower, output_file_lower]:
            file_path = os.path.join(project_dir, filename)
            try:
                os.remove(file_path)
                print(f"[INFO] Deleted intermediate file: {file_path}")
            except Exception as e:
                print(f"[ERROR] Could not delete {file_path}: {e}")

########################################
# MAIN AUTOMATION FUNCTION
########################################

def run_automation(config):
    """
    Open the InDesign template and process each model folder (subdirectories with a Credits.txt and JPG images).
    The template is first split into two documents based on a user-given empty page number.
    Automation is performed on the start document, and then the finish document is merged back.
    Finally, after all text boxes have been created, their positions are saved and the line spacing is adjusted.
    """

    project_dir   = config["project_dir"]
    template_file = config["template_file"]
    template_path = os.path.join(project_dir, template_file)
    temp_path    = os.path.join(project_dir, "temp.indd")
    start_file    = os.path.join(project_dir, "start.indd")
    finish_file   = os.path.join(project_dir, "finish.indd")
    split_page = config["split_page"]

    indd_files = [start_file, temp_path, finish_file]
    
    split_template(template_path, start_file, finish_file, split_page)

    try:
        indesign = win32com.client.Dispatch("InDesign.Application")
    except Exception as e:
        print("[ERROR] Unable to launch InDesign:", e)
        return

    try:
        doc = indesign.Documents.Add()
        doc.Save(temp_path)
    except Exception as e:
        print("[ERROR] Unable to open template:", e)
        return

    try:
        doc.DocumentPreferences.FacingPages = False
    except Exception:
        pass

    
    working_doc = doc

    target_page = config.get("target_page", None)

    model_folders = []
    for entry in os.listdir(project_dir):
        full_path = os.path.join(project_dir, entry)
        if os.path.isdir(full_path):
            if os.path.isfile(os.path.join(full_path, "Credits.txt")):
                jpg_files = [f for f in os.listdir(full_path) if f.lower().endswith(".jpg")]
                if jpg_files:
                    model_folders.append(full_path)

    if not model_folders:
        print("[WARN] No model folders found (folders with Credits.txt and JPG images).")
        return

    for i, model_folder in enumerate(model_folders):
        print(f"[INFO] Processing model folder: {model_folder}")
        tp = target_page if i == 0 and target_page is not None else None
        place_model_images(working_doc, model_folder, config, target_page=tp)
        

    # --- AFTER ALL TEXT IS WRITTEN, ADJUST THE LINE SPACING FOR ALL SAVED TEXT FRAMES ---
    try:
        new_leading_factor = config.get("leading_decrease_factor", 0.8)
        base_size = config.get("credits_font_size", 24)
        new_leading = base_size * new_leading_factor
        for tf in created_text_frames:
            tf.ParentStory.Leading = new_leading
        print(f"[INFO] Adjusted line spacing (leading) to {new_leading} for {len(created_text_frames)} text frames.")
    except Exception as e:
        print("[ERROR] Adjusting line spacing for saved text frames failed:", e)

    temp_path = os.path.join(project_dir, "temp.indd")
    output_path = os.path.join(project_dir, "output.indd")

    try:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        working_doc.Save(temp_path)
        print("[INFO] Document saved to:", temp_path)
        
    except Exception as e:
        print("[ERROR] Saving document:", e)

    
    merge_indd_files(indd_files, output_path)

    output_file = "output.indd"

    cleanup_indd_files(project_dir, template_file, output_file)


if __name__ == "__main__":
    # Sample configuration for testing.
    sample_config = {
        "project_dir": "C:/Users/swapn/Downloads/Magazine Automation/Projects/Project 1",
        "template_file": "template.indd",
        "credits_font": ["Blackadder ITC\tRegular", "Arial\tRegular"],
        "credits_font_size": 24,
        "text_box_position": ["bottom_center"],
        "layout_probabilities": {
            "single": 0.3,
            "double": 0.4,
            "four": 0.3
        },
        "target_page": 8,
        "split_page": 8,  # <-- User-provided empty page number for splitting.
        "leading_decrease_factor": 0.8,  # Factor by which line spacing will be decreased.
        "text_frame_top_left_ratio": [0.24791666666666667, 0.15185185185185185],
        "text_frame_bottom_right_ratio": [0.5916666666666667, 0.9194444444444444]
    }
    run_automation(sample_config)