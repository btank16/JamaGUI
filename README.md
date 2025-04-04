# JamaGUI
#### Developed by Blake Tanski
GUI to edit excel file outputs from Jama. This uses python to format excel files. Ther is functionality for general RTM's, dFMEAs and URRAs described below

## Getting Started
1. Check if you have python and pipinstalled by running the following in the command line:
    - python --version
    - pip --version
2. If you do not have python and pip installed, install python from https://www.python.org/downloads/
    - Then ensure you add python to your path through the following: https://realpython.com/add-python-to-path/
    - Restart your PC
3. Ensure pip is fully updated by running the following in the command line:
    - python -m pip install --upgrade pip
4. Install the required packages by running the following in the command line:
    - pip install openpyxl
    - pip install pandas
    - pip install tk
    - pip install ttkbootstrap
    - pip install ttkthemes
5. Run the application by double clicking run_app.bat
    - For easier access add a desktop shortcut to run_app.bat

## Directions for Use
Note: For all features we will be using a trace view export from jama.
Note: The "borders" option in all features adds a Thick Bottom Border to seoarate merged sections and improve readability.

### Exporting from Jama
1. Go to the trace view you want to export. Please ensure that you are starting from the top-level requirement.
2. Ensure that the "item type" field is added to the first column of the trace view at each level.
3. Click export (note the file saves as a .csv)
4. Open the file in excel and save as a .xlsx file. ALL features will ONLY work with .xlsx files.

### General Trace Matrix Merge
This will be used for RTMs or simple trace view exports.
1. Select the trace view export file
2. Select the output file name and location
3. Define the header row (this defaults to 4 if not changed)
4. Select the borders option
5. Click Generate
6. The file will be generated in the same location as the input file

### Pre-Mitigated dFMEA Formatting
This is a specialized format for dFMEAs. The cells will be merged like an RTM, but there are further inputs for risk scoring. The Risk Matrix follows the scoring as described in SPR-WI-7.1.1b, by default. Note that row inputs are integers and column inputs are letters as in excel.
1. Select the trace view export file
2. Select the output file name and location
3. Select the borders option
3. Define the following inputs:
    - Header Row (this defaults to 4 if not changed)
    - Occurrence Score Column
    - Harm ID Column
    - Severity Score Column
    - Risk Analysis Column (This should be an empty column to paste the risk analysis into)
4. Customize the Risk Matrix (per your Risk Management Plan)
5. Click Generate
6. The file will be generated in the same location as the input file

### Pre-Mitigated URRA Formatting
This is a specialized format for URRA. The cells will be merged like an RTM, but there are further inputs for risk scoring. The Risk Matrix follows the scoring as described in SPR-WI-7.1.1b, by default. Note that row inputs are integers and column inputs are letters as in excel.
1. Select the trace view export file
2. Select the output file name and location
3. Select the borders option
3. Define the following inputs:
    - Header Row (this defaults to 4 if not changed)
    - First Parent Column (User Task ID Column)
    - Second Parent Column (Use Error ID Column)
    - Occurrence Score Column
    - Harm ID Column
    - Severity Score Column
    - Risk Analysis Column (This should be an empty column to paste the risk analysis into)
4. Customize the Risk Matrix (per your Risk Management Plan)
5. Click Generate
6. The file will be generated in the same location as the input file

### Post-Mitigated dFMEA Formatting
This is a specialized format for dFMEAs. This template requires both pre and post mitigated occurrence scores to work.The cells will be merged like an RTM, but there are further inputs for risk scoring. The Risk Matrix follows the scoring as described in SPR-WI-7.1.1b, by default. Note that row inputs are integers and column inputs are letters as in excel.
1. Select the trace view export file
2. Select the output file name and location
3. Select the borders option
3. Define the following inputs:
    - Header Row (this defaults to 4 if not changed)
    - Pre-Mitigated Occurrence Score Column
    - Post-Mitigated Occurrence Score Column
    - Harm ID Column
    - Severity Score Column
    - Pre-Mitigated Risk Analysis Column (This should be an empty column to paste the risk analysis into)
    - Post-Mitigated Risk Analysis Column (This should be an empty column to paste the risk analysis into)
4. Customize the Risk Matrix (per your Risk Management Plan)
5. Click Generate
6. The file will be generated in the same location as the input file

### Post-Mitigated URRA Formatting
This is a specialized format for URRA. This template requires both pre and post mitigated occurrence scores to work.The cells will be merged like an RTM, but there are further inputs for risk scoring. The Risk Matrix follows the scoring as described in SPR-WI-7.1.1b, by default. Note that row inputs are integers and column inputs are letters as in excel.
1. Select the trace view export file
2. Select the output file name and location
3. Select the borders option
3. Define the following inputs:
    - Header Row (this defaults to 4 if not changed)
    - First Parent Column (User Task ID Column)
    - Second Parent Column (Use Error ID Column)
    - Pre-Mitigated Occurrence Score Column
    - Post-Mitigated Occurrence Score Column
    - Harm ID Column
    - Severity Score Column
    - Pre-Mitigated Risk Analysis Column (This should be an empty column to paste the risk analysis into)
    - Post-Mitigated Risk Analysis Column (This should be an empty column to paste the risk analysis into)
4. Customize the Risk Matrix (per your Risk Management Plan)
5. Click Generate
6. The file will be generated in the same location as the input file

### Risk Control Merge (Both dFMEAs and URRAs)
This tool is used to merge a risk document with its traced risk controls. You will need to do 2 exports from Jama. First, export either a dFMEA or URRA and then format it using either the pre-mitigated or post-mitigated formatting. Then, export the Fault/Failure or URRA traced to the risk controls and format it using the trace matrix merge. Note that row inputs are integers and column inputs are letters as in excel.
1. Export a dFMEA or URRA and format using either tool above
2. Export the Fault/Failure or URRA traced to the risk controls and format using the trace matrix merge
3. Select the Risk Document (either dFMEA or URRA)
4. Select the Control Document (risk control trace matrix)
5. Enter the output file name (this will be saved in the same location as the input files)
6. Add inputs for the Risk Document:
    - Header Row (this defaults to 4 if not changed)
    - Risk ID Column (column holding dFMEA or URRA ID)
    - Paste Column (an empty column to paste the risk controls into)
7. Add inputs for the Control Document:
    - Header Row (this defaults to 4 if not changed)
    - Risk ID Column (column holding dFMEA or URRA ID)
    - Risk Control Column (column holding the risk control IDs)
8. Click Generate
9. The file will be generated in the same location as the input files

## App File Structure
- UI/app contains all the file used in this application
    - app/screens contains all the screens used in this application
    - app/utils holds all the functions and holds the logic for the risk matrix (further comments are in each file)
    - main.py holds the main logic for the application and navigation between screens
- ExcelCrunch was the initial code for the functions and is not applicable (may be deleted)

