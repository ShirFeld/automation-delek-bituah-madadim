@echo off
echo Installing required packages...
pip install -r requirements.txt

echo Running Fuel Price Scraper...
python fuel_scraper.py

pause

