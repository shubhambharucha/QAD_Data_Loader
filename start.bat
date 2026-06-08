@echo off
REM QAD Data Loader - Windows Quick Start Script
REM This script starts both the FastAPI backend and React frontend

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ╔═══════════════════════════════════════════════════════════╗
echo ║     QAD Data Loader - Backend & Frontend Launcher        ║
echo ╚═══════════════════════════════════════════════════════════╝
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python not found! Please install Python 3.8+ first.
    echo    Download from: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Check if Node.js is installed
node --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Node.js not found! Please install Node.js 16+ first.
    echo    Download from: https://nodejs.org/
    pause
    exit /b 1
)

echo ✓ Python and Node.js found!
echo.

REM Install backend dependencies
echo [1/4] Installing backend dependencies...
pip install fastapi uvicorn python-multipart >nul 2>&1
if errorlevel 1 (
    echo ⚠️  Some dependencies may have failed to install
)
echo ✓ Backend dependencies ready
echo.

REM Install frontend dependencies
echo [2/4] Installing frontend dependencies...
cd frontend
call npm install >nul 2>&1
if errorlevel 1 (
    echo ⚠️  npm install had some issues, but continuing...
)
cd ..
echo ✓ Frontend dependencies ready
echo.

echo [3/4] Starting FastAPI Backend...
echo.
echo ╔───────────────────────────────────────────────────────────╗
echo ║  Backend will start at: http://localhost:8000             ║
echo ║  Close this window to stop the backend                    ║
echo ╚───────────────────────────────────────────────────────────╝
echo.

REM Start backend in a new terminal window
start cmd /k "cd /d "%~dp0" && python main.py"

REM Wait a bit for backend to start
timeout /t 3 /nobreak

echo [4/4] Starting React Frontend...
echo.
echo ╔───────────────────────────────────────────────────────────╗
echo ║  Frontend will start at: http://localhost:5173            ║
echo ║  Close this window to stop the frontend                   ║
echo ║  Press Ctrl+C to stop the dev server                      ║
echo ╚───────────────────────────────────────────────────────────╝
echo.

REM Start frontend in a new terminal window
start cmd /k "cd /d "%~dp0\frontend" && npm run dev"

echo.
echo ✅ Both servers are starting...
echo.
echo 📍 Access the app at: http://localhost:5173
echo.
echo 💡 Press Enter to close this launcher window
echo    (The backend and frontend servers will continue running)
echo.

pause
