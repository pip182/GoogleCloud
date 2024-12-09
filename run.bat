@echo off

@REM Set the delay in seconds, 2 hours = 7200 seconds
set delay=7200

echo Delaying for %delay% seconds...
timeout /T %delay% /NOBREAK

python get_exported_data.py
