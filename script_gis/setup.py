import cx_Freeze
import sys

base = None

if sys.platform == 'win32':
    base = "Win32GUI"

executables = [cx_Freeze.Executable("main_2.py", base=base, icon="amazon-logo.ico",targetName="Shipments Auto Check Tool.exe")]

cx_Freeze.setup(
    name = "SeaofBTC-Client",
    options = {"build_exe": {"packages":["tkinter","selenium","credentials","time","os","datetime","pyautogui","pandas"], "include_files":["amazon-logo.ico","credentials.py"]}},
    version = "0.01",
    description ="Gisportal Application from 3P Integrations team",
    executables = executables
    )