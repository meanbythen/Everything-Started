@echo off
echo 正在清理旧的构建文件...
rmdir /s /q build
rmdir /s /q dist

echo 开始打包程序...
python -m PyInstaller invoice_gui.spec

echo 打包完成！
echo 可执行文件位于 dist 目录中
pause 