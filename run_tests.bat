@echo off
echo ===========================================
echo     Excel to SQL Converter - Test Suite
echo ===========================================
echo.

echo Installing dependencies...
pip install -r requirements.txt
pip install -r requirements-dev.txt

echo.
echo Running tests with coverage...
python -m pytest test_excel_to_sql_converter.py -v --cov=excel_to_sql_converter --cov-report=term-missing --cov-report=html

echo.
echo ===========================================
echo Coverage report generated in htmlcov/
echo Open htmlcov/index.html to view detailed coverage
echo ===========================================
pause