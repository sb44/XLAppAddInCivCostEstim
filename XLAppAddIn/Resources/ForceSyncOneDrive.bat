@echo off
taskkill /f /im onedrive.exe
start %localappdata%\Microsoft\OneDrive\OneDrive.exe /background