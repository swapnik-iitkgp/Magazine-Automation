#!/usr/bin/env python
# automation.py

import win32com.client
import os

def split_template(template_file, start_file, finish_file, split_page):
    """
    Splits the InDesign template into two documents:
      - Pages 1 to (split_page - 1) are duplicated into the 'start' document.
      - Pages (split_page + 1) to the end are duplicated into the 'finish' document.
      
    The new documents are saved to start_file and finish_file.
    """
    # Constant for duplicating a page after a reference page.
    AFTER = 1634104421  # Adjust if necessary.

    # Launch the InDesign application.
    app = win32com.client.Dispatch("InDesign.Application")

    # Open the template document invisibly.
    template_doc = app.Open(template_file, False)
    total_pages = template_doc.Pages.Count

    # Validate the split page number.
    if split_page < 1 or split_page > total_pages:
        print("Invalid split page number")
        doc_name = template_doc.Name
        close_script = "app.documents.itemByName('{}').close(SaveOptions.NO);".format(doc_name)
        app.DoScript(close_script, 1246973031)
        return

    # Create two new documents for the "start" and "finish" portions.
    start_doc = app.Documents.Add()
    finish_doc = app.Documents.Add()

    # Duplicate pages into the "start" document.
    # Copy pages 1 to (split_page - 1) from the template into start_doc.
    for i in range(1, split_page):
        src_page = template_doc.Pages.Item(i)
        last_page = start_doc.Pages.Item(start_doc.Pages.Count)
        src_page.Duplicate(AFTER, last_page)

    # Duplicate pages into the "finish" document.
    # Copy pages (split_page + 1) to total_pages from the template into finish_doc.
    for i in range(split_page + 1, total_pages + 1):
        src_page = template_doc.Pages.Item(i)
        last_page = finish_doc.Pages.Item(finish_doc.Pages.Count)
        src_page.Duplicate(AFTER, last_page)

    # Remove the default blank page from each new document if there is more than one page.
    if start_doc.Pages.Count > 1:
        start_doc.Pages.Item(1).Delete()
    if finish_doc.Pages.Count > 1:
        finish_doc.Pages.Item(1).Delete()

    # Save the new documents.
    start_doc.Save(start_file)
    finish_doc.Save(finish_file)

    # Helper function to close a document by its name using DoScript.
    def close_doc(doc):
        doc_name = doc.Name
        script = "app.documents.itemByName('{}').close(SaveOptions.NO);".format(doc_name)
        app.DoScript(script, 1246973031)

    # Close the documents without save prompts.
    close_doc(start_doc)
    close_doc(finish_doc)
    close_doc(template_doc)

    print("Split complete.")
    print("Start file saved to:", start_file)
    print("Finish file saved to:", finish_file)

# Example usage:
if __name__ == "__main__":
    # Update these file paths as needed.
    template_file = r"C:\Users\swapn\Downloads\Magazine Automation\Projects\Project 1\template.indd"
    start_file    = r"C:\Users\swapn\Downloads\Magazine Automation\Projects\Project 1\start.indd"
    finish_file   = r"C:\Users\swapn\Downloads\Magazine Automation\Projects\Project 1\finish.indd"
    split_page = 8  # The page number where the split occurs (the empty marker page).

    split_template(template_file, start_file, finish_file, split_page)
