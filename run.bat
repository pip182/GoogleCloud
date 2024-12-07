@echo off

set delay=2
echo Delaying for %delay% seconds...
timeout /T %delay% /NOBREAK

python get_exported_data.py
