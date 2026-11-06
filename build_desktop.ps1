# 打包桌面客户端：dist\表计安装点填表器\ 和 zip
# 需要：pip install pyinstaller pywebview
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

python -m PyInstaller --noconfirm --clean --onedir --noconsole `
  --name "表计安装点填表器" `
  --icon "assets\app.ico" `
  --add-data "index.html;." `
  --add-data "assets;assets" `
  --add-data "src;src" `
  --add-data "vendor;vendor" `
  --collect-all webview `
  --hidden-import clr `
  --hidden-import pythonnet `
  --exclude-module PyQt5 --exclude-module PyQt6 `
  --exclude-module PySide2 --exclude-module PySide6 `
  --exclude-module numpy --exclude-module pandas `
  --exclude-module matplotlib --exclude-module scipy `
  --exclude-module PIL --exclude-module cryptography `
  --exclude-module tkinter `
  desktop.py

Copy-Item "使用说明.txt" "dist\表计安装点填表器\使用说明.txt" -Force
if (Test-Path "dist\表计安装点填表器.zip") { Remove-Item "dist\表计安装点填表器.zip" }
Compress-Archive -Path "dist\表计安装点填表器" -DestinationPath "dist\表计安装点填表器.zip" -CompressionLevel Optimal
Write-Host "完成：dist\表计安装点填表器.zip"


