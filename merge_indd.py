import win32com.client
import os

def merge_indd_files(indd_files, output_file):
    """
    Merges the InDesign documents specified in 'indd_files' into a single document,
    and saves it as 'output_file'. Pages from each file are appended in order.
    
    Parameters:
      indd_files (list of str): Full paths to the source InDesign documents.
      output_file (str): Full path for the merged output document.
    """
    # Constant for duplicating a page after a reference page.
    # (This value is commonly 1634104421; adjust if needed.)
    AFTER = 1634104421

    # Launch InDesign application.
    app = win32com.client.Dispatch("InDesign.Application")

    # Create a new document for merged content.
    merged_doc = app.Documents.Add()

    # Loop through each source document.
    for file_path in indd_files:
        # Open the source document invisibly.
        src_doc = app.Open(file_path, False)
        page_count = src_doc.Pages.Count

        # Loop over each page (COM collections are 1-indexed).
        for i in range(1, page_count + 1):
            src_page = src_doc.Pages.Item(i)
            # Get the last page in the merged document to serve as the insertion point.
            last_page = merged_doc.Pages.Item(merged_doc.Pages.Count)
            # Duplicate the source page into the merged document after the last page.
            src_page.Duplicate(AFTER, last_page)

        # Close the source document without saving changes.
        src_doc.Close()

    # If the merged document now has more than one page,
    # delete the original default blank page (assumed to be page 1).
    if merged_doc.Pages.Count > 1:
        merged_doc.Pages.Item(1).Delete()

    # Save the merged document to the specified file.
    merged_doc.Save(output_file)
    merged_doc.Close()

    print("Merged document saved to:", output_file)


# Example usage:
if __name__ == "__main__":
    # List of InDesign files to merge.
    indd_files = [
        r"C:\Users\swapn\Downloads\Magazine Automation\Projects\Project 1\start.indd",
        r"C:\Users\swapn\Downloads\Magazine Automation\Projects\Project 1\template.indd",
        r"C:\Users\swapn\Downloads\Magazine Automation\Projects\Project 1\finish.indd"
    ]

    # Output file path.
    output_file = r"C:\Users\swapn\Downloads\Magazine Automation\Projects\Project 1\merged_output.indd"

    merge_indd_files(indd_files, output_file)
