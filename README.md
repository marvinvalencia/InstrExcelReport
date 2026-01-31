#INSTR Logger → Excel Report Converter

A lightweight Windows application that converts INSTR Data Logger CSV exports into a formatted Excel report with automatic calculations, summaries, and charts.
This tool is for Tito Ian.

What this tool does

Reads INSTR logger CSV exports (UTF-16, tab- or comma-delimited)

Parses metadata, channel definitions, and logged readings

Down-samples data to full elapsed minutes

Calculates temperature rise (ΔT) from ambient

Generates a ready-to-use Excel report with:

Raw data

Automatic summaries

Charts

Configurable thermocouple groupings

Observations sheet for notes

The output matches the existing manual Excel report format.

Requirements (for development)

Python 3.9 or later

openpyxl

pip install -r requirements.txt

Running the app (development)
python app.py


This opens a desktop window where you:

Select the INSTR CSV file

Review or adjust thermocouple grouping options

Generate the Excel report

The output file is saved in the same folder as the input CSV.

Running the app (end users)

End users receive a Windows installer / EXE.
Python is not required on their machine.

Note: Windows Defender may show a warning because the app is not digitally signed.

Project structure
InstrExcelReport/
├── app.py                 # Desktop UI (Tkinter)
├── logger_to_report.py    # CSV parsing and Excel report generation
├── requirements.txt
└── README.md

Building the Windows installer

Installer creation is for developers only.

Build the EXE using PyInstaller

Package the EXE using Inno Setup

(See build instructions provided separately.)

Notes

Thermocouple groupings can be adjusted in the Config sheet of the generated Excel file.

The Observations sheet is intentionally free-form for test notes.
