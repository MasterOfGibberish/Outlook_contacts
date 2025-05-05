import os
import sys
import winreg
import subprocess
import ctypes
import platform
import tempfile
import site
import urllib.request

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def check_python_installed():
    try:
        python_version = platform.python_version()
        print(f"Python {python_version} detected")
        return True
    except:
        return False

def install_python():
    print("Python not detected. Attempting to download and install Python...")
    try:
        # Download Python installer
        python_url = "https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe"
        temp_dir = tempfile.gettempdir()
        installer_path = os.path.join(temp_dir, "python_installer.exe")
        
        print("Downloading Python installer...")
        urllib.request.urlretrieve(python_url, installer_path)
        
        print("Installing Python (this may take a few minutes)...")
        # Install Python with pip and adding to PATH
        subprocess.check_call([installer_path, "/quiet", "InstallAllUsers=1", "PrependPath=1", "Include_pip=1"])
        
        print("Python installed successfully!")
        return True
    except Exception as e:
        print(f"Error installing Python: {e}")
        print("Please install Python 3.7 or higher manually from https://www.python.org/downloads/")
        return False

def ensure_dependencies():
    print("Checking and installing required packages...")
    try:
        # First check if packages are already installed
        reqs = {"pywin32": False, "pandas": False, "openpyxl": False}
        installed_packages = [pkg.split('==')[0].lower() for pkg in subprocess.check_output([sys.executable, '-m', 'pip', 'freeze']).decode().split()]
        
        for pkg in installed_packages:
            if pkg == "pywin32":
                reqs["pywin32"] = True
            elif pkg == "pandas":
                reqs["pandas"] = True
            elif pkg == "openpyxl":
                reqs["openpyxl"] = True
        
        # Install missing packages
        to_install = []
        for pkg, installed in reqs.items():
            if not installed:
                to_install.append(pkg)
        
        if to_install:
            print(f"Installing missing packages: {', '.join(to_install)}")
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + to_install)
        else:
            print("All required packages are already installed!")
        
        return True
    except Exception as e:
        print(f"Error installing packages: {e}")
        print("Please run: pip install pywin32 pandas openpyxl")
        return False

def create_startup_shortcut():
    try:
        print("Creating startup shortcut for auto-loading...")
        startup_folder = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
        addin_path = os.path.abspath("addin.py")
        vbs_path = os.path.join(os.path.dirname(addin_path), "start_addin.vbs")
        
        # Create a VBS script to run the Python script hidden
        with open(vbs_path, "w") as f:
            f.write(f'CreateObject("WScript.Shell").Run "pythonw.exe {addin_path}", 0, False')
        
        # Create shortcut in startup folder
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(os.path.join(startup_folder, "OutlookContactExporter.lnk"))
        shortcut.Targetpath = vbs_path
        shortcut.WorkingDirectory = os.path.dirname(addin_path)
        shortcut.Description = "Outlook Contact Exporter"
        shortcut.save()
        
        print("Startup shortcut created!")
        return True
    except Exception as e:
        print(f"Error creating startup shortcut: {e}")
        print("The add-in will still work, but won't automatically start with Windows.")
        return False

def install_addin():
    if not check_python_installed():
        if not install_python():
            return False
    
    if not ensure_dependencies():
        return False
    
    # Get the full path of the add-in script
    addin_path = os.path.abspath("addin.py")
    
    # Register the add-in as a COM server
    print("Registering COM server...")
    try:
        subprocess.check_call([sys.executable, addin_path, "--register"])
    except Exception as e:
        print(f"Error registering COM server: {e}")
        print("Trying alternative registration method...")
        try:
            from win32com.client import makepy
            import win32com.server.register
            win32com.server.register.UseCommandLine(None)
            # Try to register more directly
            from addin import OutlookAddin
            win32com.server.register.RegisterClasses(OutlookAddin)
            print("Alternative registration successful!")
        except Exception as e2:
            print(f"Alternative registration also failed: {e2}")
            return False
    
    # Add registry entries for Outlook
    try:
        print("Adding registry entries...")
        key_path = r"Software\Microsoft\Office\Outlook\Addins\OutlookContactExporter.Addin"
        
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, "Contact Exporter Add-in for Outlook")
            winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "Contact Exporter")
            winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
            
        # Try to create HKLM entry as well if we have admin rights
        try:
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, "Contact Exporter Add-in for Outlook")
                winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "Contact Exporter")
                winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
        except:
            # It's OK if this fails - HKCU is sufficient
            pass
    except Exception as e:
        print(f"Error adding registry entries: {e}")
        return False
    
    # Try to create a startup shortcut
    create_startup_shortcut()
    
    print("✅ Installation completed successfully!")
    print("Please restart Outlook to see the 'Contact Tools' tab with the 'Save Contacts' button.")
    input("Press Enter to close...")
    return True

def uninstall_addin():
    # Unregister the COM server
    addin_path = os.path.abspath("addin.py")
    
    try:
        print("Unregistering COM server...")
        subprocess.check_call([sys.executable, addin_path, "--unregister"])
    except Exception as e:
        print(f"Error unregistering COM server: {e}")
        print("Trying alternative unregistration method...")
        try:
            # Try to unregister more directly
            from addin import OutlookAddin
            import win32com.server.register
            win32com.server.register.UnregisterClasses(OutlookAddin)
            print("Alternative unregistration successful!")
        except Exception as e2:
            print(f"Alternative unregistration also failed: {e2}")
    
    # Remove registry entries
    try:
        print("Removing registry entries...")
        # Remove from HKCU
        try:
            key_path = r"Software\Microsoft\Office\Outlook\Addins"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS) as key:
                winreg.DeleteKey(key, "OutlookContactExporter.Addin")
        except:
            pass
        
        # Remove from HKLM if possible
        try:
            key_path = r"Software\Microsoft\Office\Outlook\Addins"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_ALL_ACCESS) as key:
                winreg.DeleteKey(key, "OutlookContactExporter.Addin")
        except:
            pass
    except Exception as e:
        print(f"Error removing registry entries: {e}")
    
    # Remove startup shortcut
    try:
        print("Removing startup shortcut...")
        startup_folder = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
        shortcut_path = os.path.join(startup_folder, "OutlookContactExporter.lnk")
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
        
        vbs_path = os.path.join(os.path.dirname(addin_path), "start_addin.vbs")
        if os.path.exists(vbs_path):
            os.remove(vbs_path)
    except Exception as e:
        print(f"Error removing startup shortcut: {e}")
    
    print("✅ Uninstallation completed")
    input("Press Enter to close...")
    return True

if __name__ == "__main__":
    print("Outlook Contact Exporter Add-in Setup")
    print("=====================================")
    print("1. Install Add-in")
    print("2. Uninstall Add-in")
    
    try:
        choice = input("Enter your choice (1/2): ")
        
        if choice == "1":
            if not is_admin():
                print("This script needs administrator privileges to register COM components.")
                print("Please run this script as an administrator.")
                print("Right-click the script and select 'Run as administrator'")
                input("Press Enter to exit...")
                sys.exit(1)
            install_addin()
        elif choice == "2":
            if not is_admin():
                print("This script needs administrator privileges to unregister COM components.")
                print("Please run this script as an administrator.")
                print("Right-click the script and select 'Run as administrator'")
                input("Press Enter to exit...")
                sys.exit(1)
            uninstall_addin()
        else:
            print("Invalid choice.")
            input("Press Enter to exit...")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
        input("Press Enter to exit...") 