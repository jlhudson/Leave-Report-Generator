import os


def list_excel_files(folder_path):
    """
    List all Excel files (.xlsx and .xls) in the specified folder.

    Args:
        folder_path (str): Path to the folder to scan
    """
    print(f"\nScanning for Excel files in: {folder_path}")
    print("-" * 50)

    excel_files = []
    for file in os.listdir(folder_path):
        if file.endswith(('.xlsx', '.xls')):
            excel_files.append(file)

    if excel_files:
        print("Found Excel files:")
        for idx, file in enumerate(sorted(excel_files), 1):
            print(f"{idx}. {file}")
    else:
        print("No Excel files found in the folder.")

    print("-" * 50)
    return excel_files
