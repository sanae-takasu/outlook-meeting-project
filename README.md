# Outlook Meeting Project

This project provides tools and a GUI to extract and analyze meeting data from Microsoft Outlook.
It retrieves meetings in a specified date range, categorizes them by month, calculates total durations,
and exports the results to Excel. A Tkinter-based GUI is included for easy operation.

---

## Project Structure
```
outlook-meeting-project
├── main.py
├── app
│ ├── **init**.py
│ ├── ui
│ │ ├── **init**.py
│ │ └── gui.py
│ └── services
│ ├── **init**.py
│ └── outlook_service.py
├── tests
│ └── test_outlook_service.py
├── requirements.txt
├── outlookmeetings.spec
├── .gitignore
└── README.md
```
---

## Features

- Retrieve Outlook calendar meetings within a specified date range
- Categorize meetings by month
- Filter by MeetingStatus
- Filter or exclude categories
- Calculate total durations (minutes/hours/days)
- Export results to Excel
- Interactive GUI

---

## Installation

1. git clone <repository-url>
2. cd outlook-meeting-project
3. pip install -r requirements.txt

Requirements:

- Windows
- Outlook installed
- Python 3.8+
- pywin32

---

## Usage

### Run GUI

python main.py

### Programmatically

from app.services.outlook_service import get_meetings

---

## Testing

pytest tests/

---

## Building EXE

pyinstaller outlookmeetings.spec

---

## License

MIT License
