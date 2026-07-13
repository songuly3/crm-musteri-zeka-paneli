@echo off
cd /d "%~dp0"
title CRM Musteri Zeka Paneli V8
echo.
echo CRM Musteri Zeka Paneli V8 baslatiliyor...
echo Kaynak: final_musteri_skorlari_FINAL_3MODEL.xlsx
echo.
py -c "import streamlit,pandas,plotly,openpyxl" 2>nul
if errorlevel 1 (
    echo Eksik paketler kuruluyor...
    py -m pip install -r requirements_musteri_zeka_paneli.txt
)
echo.
echo Tarayici acilmazsa:
echo http://localhost:8506
echo.
py -m streamlit run musteri_zeka_paneli_v8.py --server.port 8506
pause
