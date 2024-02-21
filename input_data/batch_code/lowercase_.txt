@echo off
setlocal enabledelayedexpansion

cd "C:\Users\Amine Lahouani\PycharmProjects\Process_automation\Who_swho_phase_2\input_data\new_empl_pics"

for %%f in (*.*) do (
    set "filename=%%~nf"
    set "extension=%%~xf"
    set "newname=!filename!!extension!"
    
    for /f %%a in ('powershell -command "[System.IO.Path]::GetFileNameWithoutExtension('!newname!').ToLower()"') do (
        set "newname=%%a.png"
        ren "%%f" "!newname!"
    )
)

endlocal
