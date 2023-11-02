import os
import subprocess
import ctypes
import sys


def run_as_admin():
    if os.name != 'nt':
        sys.exit("This script is for Windows only.")

    # Check if the script is already running with admin rights
    if ctypes.windll.shell32.IsUserAnAdmin():
        return True
    else:
        # Re-run the script with admin rights
        result = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        if result > 32:
            return True
        else:
            return False


if __name__ == "__main__":
    if run_as_admin():
        subprocess.Popen(["python", "misc\\main.py"], shell=True)
    else:
        print("Error: User Declined Admin Privileges so Registry keys could not be added.")
