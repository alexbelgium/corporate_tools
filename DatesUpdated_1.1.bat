@echo off
SET /P folderPath="Press Enter to use the current directory or enter a new folder path: "
IF "%folderPath%"=="" SET folderPath=%~dp0

for /R "%folderPath%" %%G in (*.*) do (
    copy /b "%%G" +,, 
)

echo All files have been touched.
pause
