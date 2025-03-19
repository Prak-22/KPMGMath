#!/bin/bash

echo "Setting up the Financial Statement PDF to Excel Converter..."

# Create a virtual environment
echo "Creating Python virtual environment..."
python -m venv venv

# Activate the virtual environment
if [[ "$OSTYPE" == "darwin"* || "$OSTYPE" == "linux-gnu"* ]]; then
  source venv/bin/activate
elif [[ "$OSTYPE" == "msys" || "$OSTYPE" == "win32" ]]; then
  source venv/Scripts/activate
fi

# Install dependencies
echo "Installing Python dependencies..."
pip install -r requirements.txt

# Create output folder if it doesn't exist
mkdir -p output_folder

# Run the application
echo "Starting Streamlit application..."
streamlit run app.py
