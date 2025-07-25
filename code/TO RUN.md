# TO RUN

## Copy files to Windows directory

```bash
cp -r /home/lucas/projects/active/daily_use/window_manager/code* /mnt/c/Users/lucas/Desktop/WindowsBuildsFromWSL/window-manager/
```

# then in your windows terminal

# 1
cd C:\Users\lucas\Desktop\WindowsBuildsFromWSL\window-manager

# 2
python -m PyInstaller --onefile --windowed --name="Window Manager" --icon=myicon.ico main.py

# 3 (optional)
in wsl
git rm -r --cached exe/
in windows
"C:\Users\lucas\Desktop\WindowsBuildsFromWSL\window-manager\dist\Window Manager.exe"
copy /Y "C:\Users\lucas\Desktop\WindowsBuildsFromWSL\window-manager\dist\Window Manager.exe" "\\wsl.localhost\Ubuntu\home\lucas\projects\active\daily_use\window_manager\exe\"

## Requirements
- customtkinter: `pip install customtkinter`
- psutil: `pip install psutil`
- pywin32: `pip install pywin32`