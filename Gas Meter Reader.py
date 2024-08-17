


import os
import win32com.client

# Constants
ICON_PATH = r"C:\path\to\gas_meter_icon.ico"
SCRIPT_COMMAND = r"C:\path\to\your\run_pdf_extractor.bat"
SHORTCUT_NAME = "Run PDF Extractor.lnk"

def create_shortcut(icon_path, script_command, shortcut_name):
    """Creates a shortcut on the desktop with the specified icon and command."""
    try:
        # Verify that the icon and script paths exist
        if not os.path.isfile(icon_path):
            raise FileNotFoundError(f"Icon file not found: {icon_path}")
        if not os.path.isfile(script_command):
            raise FileNotFoundError(f"Script file not found: {script_command}")
        
        # Path to the desktop
        desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        shortcut_path = os.path.join(desktop, shortcut_name)
        
        # Initialize the WScript.Shell object
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_path)
        
        # Set shortcut properties
        shortcut.TargetPath = script_command
        shortcut.IconLocation = icon_path
        shortcut.save()
        
        print(f"Shortcut created successfully at {shortcut_path}")
    
    except FileNotFoundError as fnf_error:
        print(f"File error: {fnf_error}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Create the shortcut
create_shortcut(ICON_PATH, SCRIPT_COMMAND, SHORTCUT_NAME)





    








