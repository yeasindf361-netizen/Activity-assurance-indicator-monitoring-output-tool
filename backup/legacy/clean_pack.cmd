taskkill /f /im Tool_Build.exe >nul 2>nul
taskkill /f /im Tool_Build_debug.exe >nul 2>nul

if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
del /q *.spec 2>nul

for /d /r %%d in (__pycache__) do @rmdir /s /q "%%d" 2>nul
del /s /q *.pyc 2>nul
