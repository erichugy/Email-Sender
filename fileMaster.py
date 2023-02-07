import os
import shutil
import datetime as dt

# Global Variables --------------------------------
cwd = os.path.abspath(os.getcwd())

## Date
DATE = dt.datetime.today()
DATE_FORMAT = r"%Y-%m-%d"
DATE_STR = DATE.strftime(DATE_FORMAT)

# Subdirectory
SUBDIRECTORY = r"C:\Users\Eric Huang\Desktop\Coding\Email Sender\Old-Emails" + f"\\{DATE_STR}"


def extract_filename(path: str) -> str:
    """Extracts the filename (without the file type) from a file path.

    Args:
        path: The file path.

    Returns:
        The filename without the file type.
    """
    return os.path.splitext(os.path.basename(path))[0]


def move_files_to_subdirectory(
    filenames: list[str],
    subdirectory: str = SUBDIRECTORY 
    ) -> None:
    """Moves specified files from the current working directory into a specified subdirectory.

    Args:
        subdirectory: The name of the subdirectory to move the files into.
        filenames: The names of the files to move as a list of strings.

    Returns:
        None
    """
    cwd = os.getcwd()  # Get the current working directory
    subdirectory_path = os.path.join(cwd, subdirectory)  # Construct the full path of the subdirectory
    if not os.path.exists(subdirectory_path):  # If the subdirectory does not exist, create it
        os.mkdir(subdirectory_path)

    # Move the specified files from the current working directory into the subdirectory
    for filename in filenames:
        file_path = os.path.join(cwd, filename)
        if os.path.isfile(file_path):  # Check if the item is a file
            os.rename(file_path, os.path.join(subdirectory_path, filename))


def move_folder_to_subdirectory(folder: str, subdirectory: str) -> None:
    """Moves a folder into a specified subdirectory in the current working directory.

    Args:
        folder: The name of the folder to move.
        subdirectory: The name of the subdirectory to move the folder into.

    Returns:
        None
    """
    cwd = os.getcwd()  # Get the current working directory
    subdirectory_path = os.path.join(cwd, subdirectory)  # Construct the full path of the subdirectory
    if not os.path.exists(subdirectory_path):  # If the subdirectory does not exist, create it
        os.mkdir(subdirectory_path)

    folder_path = os.path.join(cwd, folder)  # Construct the full path of the folder
    if os.path.isdir(folder_path):  # Check if the item is a folder
        os.rename(folder_path, os.path.join(subdirectory_path, folder))




def delete_folder(folder: str) -> None:
    """Deletes a folder with a specified name in the current working directory.

    Args:
        folder: The name of the folder to delete.

    Returns:
        None
    """
    cwd = os.getcwd()  # Get the current working directory
    folder_path = os.path.join(cwd, folder)  # Construct the full path of the folder
    if os.path.isdir(folder_path):  # Check if the item is a folder
        shutil.rmtree(folder_path)



def get_file_paths(folder_path:str):
    """
    Returns a list of paths to all the files in a folder.

    Parameters:
    folder_path (str): Path to the folder.

    Returns:
    list of str: List of paths to all the files in the folder.
    """
    if not folder_path:
        return [""]
    file_paths = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            file_paths.append(file_path)
    return file_paths
