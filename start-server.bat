@echo off
cd /d "C:\Users\britt\Desktop\Sales Dashboard"
echo Starting server from: %CD%
python -m http.server 8080
pause
