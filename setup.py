import os
from cx_Freeze import setup, Executable
import shutil

# Replace 'main' with the name of your script (without the .py extension)
executables = [Executable('burndown_chart_generator.py')]

# Add the path to your PyQt5 installation directory
pyqt_path = r'C:\Python3\Lib\site-packages\PyQt5\Qt5\bin'

# Create the 'DLLs' directory in the build directory
build_dir = os.path.join(os.path.dirname(__file__), 'build')
dlls_dir = os.path.join(build_dir, 'DLLs')
os.makedirs(dlls_dir, exist_ok=True)

# Copy the required DLL files to the 'DLLs' directory
dll_files = [
    os.path.join(pyqt_path, 'Qt5Core.dll'),
    os.path.join(pyqt_path, 'Qt5Gui.dll'),
    os.path.join(pyqt_path, 'Qt5Widgets.dll'),
]
for file in dll_files:
    dest_file = os.path.join(dlls_dir, os.path.basename(file))
    if not os.path.exists(dest_file):
        shutil.copy(file, dest_file)

options = {
    'build_exe': {
        'include_msvcr': True,  # Include the MSVC runtime DLLs (Windows only)
        'packages': ['PyQt5'],  # Include the required PyQt5 package

    },
}

setup(
    name='Burndown Chart Generator',
    version='2.0',
    description='Generator Burndown Chart',
    options=options,
    executables=executables,
)
