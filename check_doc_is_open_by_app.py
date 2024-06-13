#dated 2023-07-13 , Working fine with given paths (modified for WordPad)
#Script developed by Naseem Amjad (urdujini@gmail.com)
import win32gui
import win32process
import win32api
import glob
import os
import win32con

def is_doc_file_open():
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    doc_paths = ['C:\\Docs', desktop_path]
    doc_extensions = ["*.doc", "*.docx"]
    print (f"monitoring {doc_paths} for {doc_extensions}")
    doc_files = []
    opened_files = []

    # Retrieve all .doc and .docx files in the specified paths
    for path in doc_paths:
        if os.path.exists(path):
            path_files = []
            for extension in doc_extensions:
                path_files.extend(glob.glob(os.path.join(path, extension)))
            doc_files.extend(path_files)
            #print(doc_files)

    # Iterate over the windows to find the file paths and associated process information
    def enum_windows_callback(hwnd, opened_files):
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            for file_path in doc_files:
                file_name = os.path.basename(file_path)
                if file_name.lower() in window_title.lower():
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    process_handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, pid)
                    process_info = win32process.GetModuleFileNameEx(process_handle, 0)
                    opened_files.append((file_path, process_info))
                    win32api.CloseHandle(process_handle)
                    break

    # Get all the windows
    win32gui.EnumWindows(enum_windows_callback, opened_files)

    return opened_files

# Example usage
opened_files = is_doc_file_open()

if opened_files:
    print("The following DOC or DOCX files are open in another application:")
    for file_path, process_info in opened_files:
        print(f"File: {file_path}")
        print(f"Opened by: {process_info}")
        print()
else:
    print("No DOC or DOCX file is open in another application.")
