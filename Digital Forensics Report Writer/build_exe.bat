@echo off
CD /D "%~dp0"

py -3 build_exe.py %*
if errorlevel 9009 goto :try_python
if errorlevel 1 goto :build_failed
goto :done

:try_python
python build_exe.py %*
if errorlevel 9009 goto :no_python
if errorlevel 1 goto :build_failed
goto :done

:no_python
echo.
echo Could not run build_exe.py. Install Python 3 and try again.
pause
exit /b 1

:build_failed
echo.
echo Build failed. See the messages above.
pause
exit /b 1

:done
echo.
pause
