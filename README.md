# ESLint Compare Tool
A Python tool to compare ESLint HTML reports, identify new issues, and generate CSV/Excel outputs. Includes source code and a Windows executable.

## Overview
This tool parses ESLint HTML reports from old/ and new/ folders, compares them to detect new problems (errors and warnings), and saves results as eslint_new_issues.csv and eslint_new_issues.xlsx in the output/ folder.

## Features
- Compares ESLint HTML reports for new issues.
- Generates CSV and Excel outputs with detailed issue lists.
- Supports Windows with a standalone executable.
- Ignores build-work related files.

## Installation
### Prerequisites
Python 3.x (for source code)
Required Python packages (install via pip):
- beautifulsoup4
- pandas
- openpyxl

## Install Dependencies
Create a requirements.txt file with:
- textbeautifulsoup4==4.12.2
- pandas
- openpyxl
- 
Then run:
- pip install -r requirements.txt

## Download Executable
Download eslint_compare_tool.exe from the repository.
No installation needed for the executable.

## Usage with Source Code
Clone the repository:
git clone https://gitlab.com/charlotteorsalad/eslint-compare-tool.git
cd eslint-compare-tool

## Prepare folders:
Create old/ and new/ folders.
Place ESLint HTML report files (e.g., report.html) in both folders.

## Run the script:
python eslint_compare.py

Check output/ for results.

## With Executable
- Download eslint_compare_tool.exe.
- Place it in a directory with old/ and new/ folders containing HTML reports.
- Double-click eslint_compare_tool.exe to run.
  <img width="1717" height="866" alt="image" src="https://github.com/user-attachments/assets/adcdcf19-beb9-4656-bb8a-69070c9111ca" />

- Results will be in output/.

## Output
eslint_new_issues.csv: CSV file with new issues.
eslint_new_issues.xlsx: Excel file with merged cells and auto-adjusted columns.
<img width="1487" height="161" alt="image" src="https://github.com/user-attachments/assets/b29cf9bc-6a6e-4192-93ed-731ef1f7c0ad" />


## Notes
The tool assumes one HTML file per folder (uses the first file found).
Built for Windows; source code works cross-platform with Python.

## Contributing
Feel free to fork this repository, submit issues, or create pull requests. Suggestions for improvements (e.g., multi-file support, UI) are welcome!
