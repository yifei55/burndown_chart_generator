import sys
from cx_Freeze import setup, Executable

# Replace 'main' with the name of your script (without the .py extension)
executables = [Executable('burndown_chart_generator.py')]

options = {
    'build_exe': {
        'include_msvcr': True,  # Include the MSVC runtime DLLs (Windows only)
        'packages': ['PyQt5'],  # Include the required PyQt5 package
    },
}

setup(
    name='Burndown Chart Generator',
    version='1.0',
    description='Generate the burndown chart based on your inputs',
    options=options,
    executables=executables,
)
