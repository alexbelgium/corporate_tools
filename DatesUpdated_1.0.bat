@Echo off 
rem : set path to folder of batch file, remove trailing /
set maxage=200
set mypath=%~dp0
set mypath=%mypath:~0,-1%             
echo."     _______ __                             __      __                    ___
echo."    / ____(_) /__  _____   __  ______  ____/ /___ _/ /____  _____   _   _<  /
echo."   / /_  / / / _ \/ ___/  / / / / __ \/ __  / __ \`/ __/ _\/ ___/  | | / / / 
echo."  / __/ / / /  __(__  )  / /_/ / /_/ / /_/ / /_/ / /_/  __/ /      | |/ / /  
echo." /_/   /_/_/\___/____/   \__,_/ .___/\__,_/\__,_/\__/\___/_/       |___/_/   
echo."                             /_/                                             
echo."
Echo Starting the script to update files older than %maxage% days in the current folder and its subfolders.
Echo The folder in which we will look is %mypath%
echo.
echo Ready to start ?
echo.
Pause

Echo Updating the timestamp for all files. It can be very long depending on your number of files. If you close this window it will resume when needed.
forfiles -s -p "%mypath%" -m *.* -d -%maxage% -c "cmd /c echo @file && copy /b @file +,, >nul"

echo. 
echo. All files were updated
echo. 
Pause
