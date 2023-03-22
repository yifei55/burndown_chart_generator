@echo off
rem Check if pip is installed
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    rem If pip is not installed, download and install it
    echo Installing pip...
    curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
    python get-pip.py
    del get-pip.py
    echo Pip installed successfully.
) else (
    rem If pip is already installed, display a message saying so
    echo Pip is already installed.
)


:: This script checks if the required Python packages are installed.
:: If any package is not installed, it will be installed using pip.

:: Display a message about the installation process
echo Checking and installing required packages...

:: Call the check_install subroutine for each package
call :check_install matplotlib
call :check_install numpy
call :check_install pandas
call :check_install tk
call :check_install PyQt5
call :check_install tkcalendar

:: Display a message that all packages are installed or up-to-date
echo All packages are installed or up-to-date.

:: Pause the script to see the output before closing the window
pause

:: Exit the script
exit /b

:: Define the check_install subroutine
:check_install
:: Check if the package is already installed using pip show
pip show %1 >nul 2>nul

:: If the package is not installed, the errorlevel will be 1
if errorlevel 1 (
    echo Installing %1...
    pip install %1
) else (
    echo %1 is already installed.
)

:: Return to the caller of the subroutine
exit /b
