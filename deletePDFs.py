import os
import shutil

pdfs_folder = "pdfs"

# Check if the "pdfs" folder exists
if os.path.exists(pdfs_folder):
    # Remove the entire "pdfs" folder and its contents
    shutil.rmtree(pdfs_folder)
    print(f'Deleted the "{pdfs_folder}" folder and its contents.')
else:
    print(f'"{pdfs_folder}" folder does not exist.')

