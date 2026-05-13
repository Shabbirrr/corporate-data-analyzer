# Corporate Data Analyzer

A desktop GUI tool to analyze CSV and Excel files —
group data, generate reports, and export charts.

![Demo 1](<img width="1919" height="1079" alt="Screenshot 2026-05-13 090631" src="https://github.com/user-attachments/assets/c390b4ac-e183-4bcc-960f-f9dc736bc5f1" />
)
![Demo 2](<img width="1919" height="1070" alt="Screenshot 2026-05-13 090450" src="https://github.com/user-attachments/assets/510f1830-e4af-4a88-a328-a94c98c93923" />
)

## Features
- Load CSV or Excel files
- Group by any column, aggregate (sum/avg/max/min)
- Preview results in a table
- Bar, Pie, Line, and Column charts
- Export reports as Excel or CSV
- Export charts as PNG

## Tech Stack
Python · Tkinter · Pandas · Matplotlib

## Project Structure
corporate-data-analyzer/
├── src/
│   └── File_analyzer.py
├── sample_data/
│   └── sample_sales.xlsx
└── README.md

## How to run
pip install -r requirements.txt
python src/File_analyzer.py

## Sample Data
A sample sales dataset is included in sample_data/
to test the app right away.
