import win32com.client
from fileMaster import cwd, extract_filename, move_files_to_subdirectory, delete_folder



def word_to_html(full_path: str, filename: str = None, save_to_path: str = cwd) -> None:
    """Converts a Word document to HTML and saves it to a specified location.

    Args:
        full_path: The full path of the Word document to be converted.
        filename: The name to be given to the new HTML file.
        save_to_path: The path where the new HTML file should be saved (default is the current working directory).

    Returns:
        None
    """
    #Extract filename from path
    filename = extract_filename(full_path) if not filename else filename

    # Set full path for new file
    new_filename = save_to_path + '\\' + filename + '.htm'


    # Open Word
    word = win32com.client.Dispatch('Word.Application')

    # Open the document
    doc = word.Documents.Open(full_path)

    # Save the document as HTML
    try:
        doc.SaveAs(new_filename, FileFormat=8)
    finally:
        # Close the document and quit Word
        doc.Close()
        word.Quit()

    # Move Word Doc to respective Old-Emails subfolder 
    move_files_to_subdirectory([f"{filename}.docx"])
    
    # Delete Unecessary folder created
    delete_folder(filename + "_files")

if __name__ == '__main__':
    word_to_html(r"C:\Users\Eric Huang\Desktop\msg_2023-01-09.docx","msg_2023-01-09")