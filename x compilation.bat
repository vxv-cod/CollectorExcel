start/w setversion.exe

pyinstaller -w -F -i "logo.ico" CollectorExcel.py

xcopy %CD%\*.xltx %CD%\dist /H /Y /C /R
xcopy %CD%\*.ico %CD%\dist /H /Y /C /R
xcopy %CD%\*.ini %CD%\dist /H /Y /C /R

xcopy C:\vxvproj\tnnc-Excel\collectorExcel\collectorApp\dist C:\vxvproj\tnnc-Excel\collectorExcel\ConsoleApp\ /H /Y /C /R



