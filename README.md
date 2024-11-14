# Meiragtx Project: Case Study Overview

This project provides an Excel-based application that integrates with Python via **xlwings** to process and analyze well data. It includes multiple Python scripts that are triggered by VBA macros in an Excel workbook, allowing users to process data in an automated and streamlined manner.

---

## Table of Contents

- [File Overview](#file-overview)
  - [Application File](#1-application-file)
  - [Python Script Files](#2-python-script-files)
  - [Function Files](#3-function-files)
- [Running the Excel File](#running-the-excel-file)
  - [Step 1: Install Python](#step-1-install-python)
  - [Step 2: Install the xlwings Library](#step-2-install-the-xlwings-library)
- [Using the Application](#using-the-application)

---

## File Overview

### 1. Application File
- **AnalysisExample_app.xlsm**  
  This is the main Excel application file that provides the user interface and contains VBA macros. These macros allow you to process data using the Python scripts provided. The macros are activated by buttons within the workbook.

### 2. Python Script Files
These Python scripts are linked to the VBA buttons in Excel. They are designed to perform specific data processing tasks.

- **`standard_analysis.py`**  
  Corresponds to the **Standard Analysis** button in Excel. This script processes data according to specifications for the wells. The input file must contain a "ct values" column and a "Well" column, which identifies the wells.

- **`sample_analysis.py`**  
  Corresponds to the **Import Sample Data** button. It processes the **ct values** and additional data in the imported file.

- **`process_samples.py`**  
  Linked to the **Process Non-Excluded Samples** button, this script calculates averages for non-excluded samples.

### 3. Function Files
These files include reusable functions for various tasks:

- **`General.py`**  
  Contains general functions for importing and exporting data, as well as for reading Excel files.

- **`Standards_class.py`**  
  Processes data related to the specification of filled wells.

- **`Samples_class.py`**  
  Contains functions to handle sample data and calculate averages for the samples.

---

## Running the Excel File

To run the **AnalysisExample_app.xlsm** file and use the associated Python scripts, make sure that **Python** is installed on your system and correctly configured in the system’s **PATH**.

### Step 1: Install Python

1. Download Python from the official website:  
   [Download Python](https://www.python.org/downloads/)

2. During the installation process, ensure that you check the box to **Add Python to PATH**. This will make Python accessible from the command line.

   ![Python Installation Path](https://github.com/user-attachments/assets/b24992c4-c9bd-43e9-acbe-5b1633d87a0e)

### Step 2: Install the xlwings Library

To interface Python with Excel, install the **xlwings** library through these instructions : https://docs.xlwings.org/en/stable/addin.html

## Using the Application

Once everything is set up, you can start using the **AnalysisExample_app.xlsm** Excel file:

1. **Open the Excel file** (AnalysisExample_app.xlsm) in Excel.
   
2. You will see buttons corresponding to different Python scripts (e.g., **Standard Analysis**, **Import Sample Data**, **Process Non-Excluded Samples**, etc.).
   
3. Clicking on a button triggers the associated Python script to process the data as defined in the script. Each button will perform a specific action based on the file you import into the application.
   
4. Follow the prompts in the application to upload files and process data.

5. The results of the processed data will be visible in the application, and depending on the script, may be saved to a new Excel file or displayed within the current workbook.

Each Python script is connected to a specific Excel macro. When you click a button in the Excel file, the relevant Python function is executed, processing your data directly from the Excel interface.

### Key VBA Buttons

- **Standard Analysis Button**: Processes well data based on specifications. Requires a file with a "Well" column, you can pick the ct collumn.
- **Import Sample Data Button**: Processes ct values and additional data from the imported sample file.
- **Process Non-Excluded Samples Button**: Calculates averages for non-excluded samples.







