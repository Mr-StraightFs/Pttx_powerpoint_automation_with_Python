@echo off
setlocal enabledelayedexpansion

cd "C:\Users\Amine Lahouani\PycharmProjects\Process_automation\Who_swho_phase_2\input_data\new_extract\Remaining missing picture 27th Nov\nov pics"

for %%f in (*.*) do (
    set "filename=%%~nf"
    for %%F in ("!filename!") do (
        set "newname=%%~nxF"
        ren "%%f" "!newname!.png"
    )
)

for %%g in (*.*) do (
    set "newname=%%~ng"
    ren "%%g" "!newname!.png"
)

endlocal
