#!/bin/bash

echo "Setting up the Financial Statement PDF to Excel Converter..."


echo "Creating Python virtual environment..."
python -m venv venv


if [[ "$OSTYPE" == "darwin"* || "$OSTYPE" == "linux-gnu"* ]]; then
  source venv/bin/activate
elif [[ "$OSTYPE" == "msys" || "$OSTYPE" == "win32" ]]; then
  source venv/Scripts/activate
fi


echo "Installing Python dependencies..."
pip install -r requirements.txt


mkdir -p output_folder


echo "Starting Streamlit application..."
streamlit run app.py
