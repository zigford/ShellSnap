REM @Echo off
SET SOURCE=%~dp0
SET NAME=%~n0
SET PATH=%PATH%;%SOURCE%

REM Usage: SetBG.bat %_SMSTSBootImageID% relativepathtojpg

If %1==USC00002 (
SET WALLPAPEREXE=%SOURCE%x64\wallpaper.exe
) Else (
SET WALLPAPEREXE=%SOURCE%x86\wallpaper.exe
)

%WALLPAPEREXE% %SOURCE%%2