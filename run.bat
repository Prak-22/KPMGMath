@echo off
ECHO Setting up the Financial Statement PDF to Excel Converter...

REM Create a virtual environment
ECHO Creating Python virtual environment...
python -m venv venv

REM Activate the virtual environment
ECHO Activating virtual environment...
call venv\Scripts\activate

REM Install dependencies
ECHO Installing Python dependencies...
pip install -r requirements.txt

REM Create output folder if it doesn't exist
if not exist output_folder mkdir output_folder

REM Run the application
ECHO Starting Streamlit application...
streamlit run app.py

pause
