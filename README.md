# Cloudscape_ScopeAnalyzer
Script to automate analysis of the Export Scope report for Cloudscape asset discovery

## :white_check_mark: Prerequisites:

* Working version of [Python](https://www.python.org/downloads/)
* [PIP Package Manager](https://pip.pypa.io/en/stable/installing/) for Python
* [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/) Python Library installed via PIP

## :running: How to run:
Copy script locally and download the **Export Scope** report within Cloudscape from *Consume Intelligence -> Assets -> Assessment Coverage -> Export Scope*  
Save downloaded CSV file to same working directory as script and run the script with the name of the CSV file as an argument: exportscope.py *filename.csv*  
The script will copy the CSV file into an identically named XLSX file and create new tabs called **Bad Credentials** and **No Connectivity**.  The resulting XLSX file will include the original report data, a list of the devices with Bad Credentials (where the *Inaccessible* column = 1) and a list of the devices with No Connectivity (where the *Total IPs* column = 0) along with a running total for each type of error.
