@echo off
echo Current directory: %CD%
echo.
echo Files in current directory:
dir /b
echo.
echo Starting server on port 3000...
python -m http.server 3000
