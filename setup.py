import sys
from cx_Freeze import setup, Executable

# Build options
build_options = {
    'packages': ['os', 'tkinter', 'pandas', 'pyodbc', 'configparser', 'apscheduler'],
    'excludes': [],
    'include_files': ['config.ini'],  # Include the config.ini file
}

# Executable options
executables = [
    Executable(
        script='MW.py',  # This is the target script for the executable
        base="Win32GUI",  # Set base to Win32GUI for graphical application
        targetName='MW Sales Report.exe',  
    )
]

# Create the setup
setup(
    name='MW Sales Report',  # Name of the application
    version='1.0',
    description='MW Sales Report',
    options={'build_exe': build_options},
    executables=executables,
)
