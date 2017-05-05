@echo off
cls

REM Wipe Drive script
REM Aaron Czechowski, Microsoft Consulting Services
REM 4 December 2007

REM display help if not parameter is specified
if !%1==! goto SYNTAX

REM initialize variable for total number of root files/folders
set /a intTotal=0

REM loop through each root file/folder and increment counter if it's not C:\_SMSTaskSequence
FOR /f "delims=" %%i in ('dir c:\ /b /a') do    if not %%i==_SMSTaskSequence    set /a intTotal+=1


echo wiping drive ...
echo.

REM initialize variable for current file/folder
set /a intCount=1

REM loop through each root file/folder,
REM   if it's not C:\_SMSTaskSequence, pass the object name to the PROCESSOR section
for /f "delims=" %%i in ('dir c:\ /b /a /o:n') do    if not %%i==_SMSTaskSequence    call :PROCESSOR "%%i"

echo.
echo Processing complete.
echo.
goto :EOF


:PROCESSOR
      REM %%i variable from above "call" is interpreted here as the first parameter (%~1)
      REM %~1 expands %1 removing any surrounding quotes (")
      REM !intCount! relies upon delayed environment variable expansion to be enabled
      echo          ... deleting !intCount! of %intTotal% (%~1^)
      del "c:\%~1" /q /f /a > NUL 2>&1
      rmdir "c:\%~1" /s /q > NUL 2>&1

      set /a intCount+=1

      REM return to just after the "call" above
      goto :EOF

:SYNTAX
     echo You must supply at least one argument to run this script.
     echo Delayed environment variable expansion (/v:on) should be
     echo    enabled to correctly display the progress.
     echo.
     echo    cmd /v:on /c wipe.cmd wipe