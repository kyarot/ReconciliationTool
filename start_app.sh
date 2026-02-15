#!/bin/bash
# Launch script for Financial Reconciliation Web App

echo "Starting Financial Reconciliation Web App..."
echo "=================================="

# Activate virtual environment
source .venv/bin/activate

# Launch Streamlit
echo "Launching web interface..."
streamlit run app.py