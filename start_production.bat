@echo off
echo Starting Saleem Backend in Production Mode...
cd /d "%~dp0"
call .venv\Scripts\activate
:: We run without --reload and use --proxy-headers if behind a reverse proxy like Nginx/IIS
python -m uvicorn main:app --host 0.0.0.0 --port 8001 --proxy-headers
