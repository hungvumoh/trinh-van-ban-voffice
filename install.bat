@echo off
title Cai dat Tro ly trinh van ban
echo.
echo   Dang cai "Tro ly trinh van ban". Vui long doi...
echo   (Tai Python + thu vien + chuong trinh. Lan dau co the mat vai phut.)
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex (iwr -useb 'https://raw.githubusercontent.com/hungvumoh/trinh-van-ban-voffice/main/setup.ps1').Content"
echo.
pause
