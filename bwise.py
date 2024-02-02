import csv
import winreg
import os
import datetime

def get_installed_programs():
    programs = []
    uninstall_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall")
    num_subkeys = winreg.QueryInfoKey(uninstall_key)[0]

    for i in range(num_subkeys):
        program = {}
        subkey_name = winreg.EnumKey(uninstall_key, i)
        subkey = winreg.OpenKey(uninstall_key, subkey_name)
        
        try:
            program["Name"] = winreg.QueryValueEx(subkey, "DisplayName")[0]
            program["Version"] = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
            program["Publisher"] = winreg.QueryValueEx(subkey, "Publisher")[0]
            program["InstallLocation"] = winreg.QueryValueEx(subkey, "InstallLocation")[0]
            install_date = winreg.QueryValueEx(subkey, "InstallDate")[0]
            program["InstallDate"] = datetime.datetime.strptime(install_date, "%Y%m%d").strftime("%Y-%m-%d")
            programs.append(program)
        except OSError:
            pass

    return programs

def save_to_csv(programs, filename, directory=None):
    fieldnames = ["Name", "Version", "Publisher", "InstallLocation", "InstallDate"]

    if directory is not None:
        filename = os.path.join(directory, filename)

    with open(filename, "w", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(programs)

installed_programs = get_installed_programs()
save_to_csv(installed_programs, "installed_programs.csv", directory=r"C:\Users\benjo\OneDrive\Desktop")