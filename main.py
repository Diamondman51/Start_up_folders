# import os
# from win32com.client import Dispatch
#
#
# def get_startup_folder():
#     shell = Dispatch("WScript.Shell")
#     startup_folder = shell.SpecialFolders("Startup")
#     return startup_folder
#
#
# def get_common_startup_folder():
#     shell = Dispatch("WScript.Shell")
#     common_startup_folder = shell.SpecialFolders("AllUsersStartup")
#     return common_startup_folder
#
#
# if __name__ == "__main__":
#     user_startup = get_startup_folder()
#     common_startup = get_common_startup_folder()
#
#     print(f"User Startup Folder: {user_startup}")
#     print(f"Common Startup Folder: {common_startup}")
#

# All options

import os
import winreg
from win32com.client import Dispatch


def get_startup_folder():
    shell = Dispatch("WScript.Shell")
    return shell.SpecialFolders("Startup")


def get_common_startup_folder():
    shell = Dispatch("WScript.Shell")
    return shell.SpecialFolders("AllUsersStartup")


def get_registry_startup_entries(root_key, sub_key):
    try:
        with winreg.OpenKey(root_key, sub_key) as key:
            i = 0
            entries = {}
            while True:
                try:
                    name, value, _ = winreg.EnumValue(key, i)
                    entries[name] = value
                    i += 1
                except OSError:
                    break
            return entries
    except FileNotFoundError:
        return {}


if __name__ == "__main__":
    # Get Startup Folder Paths
    user_startup = get_startup_folder()
    common_startup = get_common_startup_folder()

    print(f"User Startup Folder: {user_startup}")
    print(f"Common Startup Folder: {common_startup}")

    # Get Registry Startup Entries
    user_registry_startup = get_registry_startup_entries(winreg.HKEY_CURRENT_USER,
                                                         r"Software\Microsoft\Windows\CurrentVersion\Run")
    machine_registry_startup = get_registry_startup_entries(winreg.HKEY_LOCAL_MACHINE,
                                                            r"Software\Microsoft\Windows\CurrentVersion\Run")

    print("\nUser Registry Startup Entries:")
    for name, path in user_registry_startup.items():
        print(f"{name}: {path}")

    print("\nMachine Registry Startup Entries:")
    for name, path in machine_registry_startup.items():
        print(f"{name}: {path}")

