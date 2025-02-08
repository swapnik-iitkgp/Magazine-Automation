#!/usr/bin/env python
# automation.py

import os
import time
import pyautogui
import cv2
import numpy as np
import win32com.client
import pywintypes

from config_module import load_config, save_config

########################################
# OPENCV CORNER SELECTION WITH RATIO
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
    1) Take a screenshot
    2) Show it in an OpenCV window
    3) Let user click TOP-LEFT, then BOTTOM-RIGHT
    4) Return ((ratio_left, ratio_top), (ratio_right, ratio_bottom)) 
       as screen-ratio corners
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

    print("[INFO] Click the TOP-LEFT corner, then the BOTTOM-RIGHT corner.")
    while True:
        cv2.imshow("Screenshot", clone)
        key = cv2.waitKey(1) & 0xFF
        if len(_selected_points) >= 2:
            break
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

    print(f"[INFO] Corner ratios: L={ratio_left:.3f}, T={ratio_top:.3f}, R={ratio_right:.3f}, B={ratio_bottom:.3f}")
    return (ratio_left, ratio_top), (ratio_right, ratio_bottom)

def retrieve_ratio_region(config):
    """
    If ratio-based corners exist in config, convert them to actual pixel coords
    for the current screen. Otherwise return None.
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
# INDESIGN + PYAUTOGUI AUTOMATION
########################################

def place_images(project_dir, template_file):
    """
    Open the InDesign template, place .jpg images, FILL the page proportionally,
    and save to output.indd. Returns (doc, output_path) with the doc left open.
    """
    template_path = os.path.join(project_dir, template_file)
    try:
        indesign = win32com.client.Dispatch("InDesign.Application")
    except Exception as e:
        print("[ERROR] Unable to launch InDesign:", e)
        return None, None

    try:
        doc = indesign.Open(template_path)
    except Exception as e:
        print("[ERROR] Unable to open template:", e)
        return None, None

    # Disable facing pages if desired
    try:
        doc.DocumentPreferences.FacingPages = False
    except:
        pass

    try:
        page_width = doc.DocumentPreferences.PageWidth
        page_height = doc.DocumentPreferences.PageHeight
    except Exception as e:
        print("[ERROR] Could not get page dimensions:", e)
        doc.Close()
        return None, None

    # Get or create a layer for images
    try:
        imageLayer = doc.Layers.Item("ImageLayer")
    except pywintypes.com_error:
        imageLayer = doc.Layers.Add()
        imageLayer.Name = "ImageLayer"

    # Gather .jpg images
    image_files = sorted([f for f in os.listdir(project_dir) if f.lower().endswith(".jpg")])
    if not image_files:
        print("[WARN] No .jpg files found in the project directory.")
        return doc, None

    for idx, image_name in enumerate(image_files):
        image_path = os.path.join(project_dir, image_name)
        print("[INFO] Placing image:", image_path)

        # Use page 1 for the first image, subsequent images on new pages
        if idx == 0 and doc.Pages.Count >= 1:
            page = doc.Pages.Item(1)
        else:
            page = doc.Pages.Add()

        # Create a rectangle that spans the entire page
        rect = page.Rectangles.Add()
        rect.GeometricBounds = [0, 0, page_height, page_width]

        try:
            rect.FrameFittingOptions.AutoFit = False
            rect.Place(image_path)
            # numeric for FILL_PROPORTIONALLY = 1818584694
            rect.Graphics[0].Fit(1818584697)
        except Exception as e:
            print("[ERROR] Placing/fitting image:", e)

        rect.ItemLayer = imageLayer

    # Save the document as output.indd
    output_path = os.path.join(project_dir, "output.indd")
    try:
        if os.path.exists(output_path):
            os.remove(output_path)
        doc.Save(output_path)
        print("[INFO] Document saved to:", output_path)
    except Exception as e:
        print("[ERROR] Saving document:", e)

    return doc, output_path

def scroll_to_first_page(doc):
    """
    Attempt to scroll or jump to the first page via COM; fallback with PyAutoGUI if needed.
    """
    try:
        if doc.Windows.Count > 0:
            window = doc.Windows.Item(1)
            window.ActivePage = doc.Pages.Item(1)
            print("[INFO] Scrolled to page 1 via COM.")
        else:
            print("[WARN] No doc window found. Using fallback keystrokes.")
            for _ in range(5):
                pyautogui.hotkey('ctrl', 'pageup')
                time.sleep(0.2)
    except Exception as e:
        print("[WARN] scroll_to_first_page error:", e)
        for _ in range(5):
            pyautogui.hotkey('ctrl', 'pageup')
            time.sleep(0.2)

def compute_text_box_coordinates(region_top_left, region_bottom_right, config):
    """
    Hard-code a 3-line text box region (for demonstration).
    """
    credits_file  = config.get("credits_file")
    project_dir   = config["project_dir"]
    path_credits = os.path.join(project_dir, credits_file)
    font_size = config.get("credits_font_size", 24)
    position  = config.get("text_box_position", "center")

    with open(path_credits, "r", encoding="utf-8") as f:
        line_count = sum(1 for line in f)

    print(line_count)

    width  = region_bottom_right[0] - region_top_left[0]
    height = region_bottom_right[1] - region_top_left[1]

    line_multiplier = 1.2
    box_h = int((line_count+1) * font_size * line_multiplier)  # 3 lines
    box_w = width

    if position == "center":
        tl_x = region_top_left[0]
        tl_y = region_top_left[1] + (height - box_h)//2
    elif position == "top_center":
        tl_x = region_top_left[0]
        tl_y = region_top_left[1]
    elif position == "bottom_center":
        tl_x = region_top_left[0]
        tl_y = region_bottom_right[1] - box_h
    else:
        # default to center
        tl_x = region_top_left[0]
        tl_y = region_top_left[1] + (height - box_h)//2

    box_tl = (tl_x, tl_y)
    box_br = (tl_x + box_w, tl_y + box_h)
    click_pt = ((box_tl[0] + box_br[0]) // 2, (box_tl[1] + box_br[1]) // 2)
    return box_tl, box_br, click_pt

def insert_text_frame_and_type(text_content, drag_start, drag_end, click_point, config):
    """
    Draw a text frame via PyAutoGUI & apply font/size using COM.
    """
    font_name = config.get("credits_font", "Arial\tRegular")
    font_size = config.get("credits_font_size", 24)

    print("[INFO] Creating text frame via PyAutoGUI...")

    time.sleep(1)
    # Press 't' to select Type Tool
    pyautogui.press('t')
    time.sleep(0.5)

    # Drag the text frame
    pyautogui.moveTo(drag_start[0], drag_start[1], duration=0.5)
    pyautogui.mouseDown()
    pyautogui.moveTo(drag_end[0], drag_end[1], duration=1)
    pyautogui.mouseUp()
    time.sleep(0.5)

    # Click inside the frame to set insertion point
    pyautogui.moveTo(click_point[0], click_point[1], duration=0.5)
    pyautogui.click()
    time.sleep(0.5)
    print("[INFO] Text frame created. Applying font...")

    # Apply the font/size using COM
    try:
        indesign = win32com.client.Dispatch("InDesign.Application")
        doc = indesign.ActiveDocument
        if doc.Selection.Count > 0:
            sel = doc.Selection[0]
            sel.AppliedFont = font_name
            sel.PointSize   = font_size
            print(f"[INFO] Applied font={font_name}, size={font_size}")
        else:
            print("[WARN] No selection in doc.")
    except Exception as e:
        print("[ERROR] Could not apply font/size:", e)

    # Type the text content
    pyautogui.typewrite(text_content, interval=0.05)
    print("[INFO] Finished typing text.")

def run_automation(config):
    """
    1) place_images (fill page proportionally)
    2) scroll_to_first_page
    3) retrieve or define text region
    4) compute_text_box_coordinates
    5) insert_text_frame_and_type
    """
    project_dir   = config["project_dir"]
    template_file = config["template_file"]
    credits_file  = config.get("credits_file")

    # 1) Place and fill images
    doc, out_path = place_images(project_dir, template_file)
    if not doc:
        print("[ERROR] Could not open doc or place images.")
        return

    # 2) Scroll to first page
    scroll_to_first_page(doc)

    # 3) Determine region for text (either from ratio or direct coords)
    region = retrieve_ratio_region(config)
    if region is not None:
        region_top_left, region_bottom_right = region
    else:
        if ("text_frame_top_left" in config) and ("text_frame_bottom_right" in config):
            region_top_left  = tuple(config["text_frame_top_left"])
            region_bottom_right = tuple(config["text_frame_bottom_right"])
        else:
            ratio_tl, ratio_br = get_region_from_opencv()
            config["text_frame_top_left_ratio"]    = list(ratio_tl)
            config["text_frame_bottom_right_ratio"] = list(ratio_br)
            save_config(config, os.path.join(project_dir, "config.json"))
            region = retrieve_ratio_region(config)
            if region is None:
                print("[ERROR] Could not compute region from ratio.")
                return
            region_top_left, region_bottom_right = region

    # 4) Compute box & click point
    box_tl, box_br, click_pt = compute_text_box_coordinates(region_top_left, region_bottom_right, config)

    # 5) Insert text
    text_content = "Hello from PyAutoGUI!"
    if credits_file:
        path_credits = os.path.join(project_dir, credits_file)
        if os.path.isfile(path_credits):
            with open(path_credits, "r", encoding="utf-8") as f:
                text_content = f.read().strip()

    insert_text_frame_and_type(text_content, box_tl, box_br, click_pt, config)
    print("[INFO] Automation complete.")

if __name__ == "__main__":
    # Minimal test
    sample_config = {
        "project_dir": "C:\\Temp\\MyProject",
        "template_file": "template.indd",
        "credits_file": "Credits.txt",
        "credits_font": "Arial\tRegular",
        "credits_font_size": 24
    }
    run_automation(sample_config)
