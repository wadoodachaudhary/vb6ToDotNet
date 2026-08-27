@echo off
rem Double-click entry point: runs run-bench.ps1 without changing the machine's
rem PowerShell execution policy.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run-bench.ps1" %*
