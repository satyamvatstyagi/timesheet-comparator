#!/bin/bash
echo "======================================"
echo "Setting up Python venv for Jupyter"
echo "======================================"

# Step 1: Create virtual environment
echo "Creating virtual environment (.venv)..."
python3 -m venv .venv

# Step 2: Activate venv
echo "Activating virtual environment..."
source .venv/bin/activate

# Step 3: Upgrade pip
echo "Upgrading pip..."
python -m pip install --upgrade pip

# Step 4: Install Jupyter and notebook dependencies
echo "Installing Jupyter and required libraries..."
pip install jupyter ipykernel notebook pandas

# Step 5: Register kernel
echo "Registering venv as Jupyter kernel..."
python -m ipykernel install --user --name=.venv --display-name "Python (.venv)"

echo "======================================"
echo "Setup complete!"
echo "Open VS Code, install Python + Jupyter extensions,"
echo "then select kernel: Python (.venv)"
echo "======================================"