@echo off
setlocal enabledelayedexpansion

set "directory=C:\Users\Amine Lahouani\PycharmProjects\Process_automation\Who_swho_phase_2\input_data\new_extract\Remaining missing picture 27th Nov\nov pics"

for %%f in ("%directory%\*.*") do (
    set "fullpath=%%f"
    for %%g in (!fullpath!) do (
        set "filename=%%~ng"
        set "extension=%%~xg"
        
        rem Rename the file to remove the extension
        ren "!fullpath!" "!filename!"
    )
)

echo File extensions removed.
pause
