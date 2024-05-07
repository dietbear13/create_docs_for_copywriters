@echo off

IF EXIST "%cd%\Scripts\pip.exe" (
    echo Found pip.exe.
) ELSE (
    echo Could not find pip.exe in the Scripts folder.
    echo Make sure the Scripts folder is located in the current directory.
    echo Run the library installation after installing Python.
    pause
    exit
)

echo Installing necessary libraries...
"%cd%\Scripts\pip.exe" install pandas google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client openpyxl requests beautifulsoup4

echo Library installation completed.
pause
