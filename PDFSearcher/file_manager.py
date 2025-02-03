import os
import PyPDF2
import shutil

def move_file (results, source_folder, destination_folder):
    """
    Moves files with "Back Margin" to a specified folder.

    Args:
        results (list): The list of file names with "Back Margin" status.
        source_folder (str): The path to the source folder where PDFs are located.
        destination_folder (str): The path to the folder where "Back Margin" files should be moved.
    """
    if not os.path.exists(destination_folder):
        try:
            os.makedirs(destination_folder)
            print(f"Folder created at: {destination_folder}")  # Debugging line
        except Exception as e:
            print(f"Error creating folder: {e}")
            return  # Exit if folder creation fails

    for result in results:
        if result["Back Margin"] == "Yes":
            file_path = os.path.join(source_folder, result["File Name"])
            new_path = os.path.join(destination_folder, result["File Name"])
            
            try:
                shutil.copy(file_path, new_path)
                print(f"Moved file: {file_path} -> {new_path}")  # Debugging line
            except Exception as e:
                print(f"Error moving file {file_path}: {e}")

def move_file2 (results, source_folder, destination_folder):
    """
    Moves files with "Back Margin" to a specified folder.

    Args:
        results (list): The list of file names with "Back Margin" status.
        source_folder (str): The path to the source folder where PDFs are located.
        destination_folder (str): The path to the folder where "Back Margin" files should be moved.
    """
    if not os.path.exists(destination_folder):
        try:
            os.makedirs(destination_folder)
            print(f"Folder created at: {destination_folder}")  # Debugging line
        except Exception as e:
            print(f"Error creating folder: {e}")
            return  # Exit if folder creation fails

    for result in results:
        if result["Marketing Support"] == "Yes":
            file_path = os.path.join(source_folder, result["File Name"])
            new_path = os.path.join(destination_folder, result["File Name"])
            
            try:
                shutil.copy(file_path, new_path)
                print(f"Moved file: {file_path} -> {new_path}")  # Debugging line
            except Exception as e:
                print(f"Error moving file {file_path}: {e}")

