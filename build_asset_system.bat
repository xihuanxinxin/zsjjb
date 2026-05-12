@echo off
echo 正在打包资产管理系统...

cd /d "C:\Users\aa-u\Documents\trae_projects\Project-MMM"

python "C:\Users\aa-u\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.9_qbz5n2kfra8p0\LocalCache\local-packages\Python39\site-packages\PyInstaller\__main__.py" ^
    --onedir ^
    --windowed ^
    --name=资产管理系统 ^
    --add-data="C:\Program Files\WindowsApps\PythonSoftwareFoundation.Python.3.9_3.9.3568.0_x64__qbz5n2kfra8p0\python39.dll;." ^
    gui.py

echo 打包完成！
pause
