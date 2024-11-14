#!/bin/bash

# Function to install dependencies from requirements.txt
install_dependencies() {
    if [ -f "requirements.txt" ]; then
        echo "Installing dependencies from requirements.txt..."
        pip install -r requirements.txt
    else
        echo "requirements.txt file not found."
        exit 1
    fi
}

# Function to install the xlwings add-in
install_xlwings_addin() {
    echo "Installing xlwings add-in..."
    xlwings addin install
}

# Run 
install_dependencies
install_xlwings_addin
