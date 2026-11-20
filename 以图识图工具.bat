@echo off
chcp 65001 >nul
title 以图识图工具 - 反向图片搜索
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0reverse_image_search.ps1"
