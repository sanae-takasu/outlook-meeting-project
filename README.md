# Outlook Meeting Project

This project is designed to interact with Microsoft Outlook's calendar to extract meeting details and prepare data for output. It provides functionality to retrieve meetings within a specified date range, categorize them by month, and calculate total durations.

## Project Structure

```
outlook-meeting-project
├── outlookmeeting
│   ├── __init__.py
│   ├── gui.py
│   ├── outlookmeeting.py
├── tests
│   └── test_outlookmeeting.py
├── requirements.txt
├── .gitignore
└── README.md
```

## Features

- Retrieve meetings from Microsoft Outlook within a specified date range.
- Categorize meetings by month.
- Calculate total meeting durations.
- Export meeting data to an Excel file for further analysis.
- Interactive GUI for user-friendly operation.

## Installation

To set up the project, follow these steps:

1. Clone the repository:
   ```
   git clone <repository-url>
   ```
2. Navigate to the project directory:
   ```
   cd outlook-meeting-project
   ```
3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

**Note:** This project requires Python 3.8 or higher. Ensure you have the correct version installed before proceeding.

## Usage

### Running the GUI Application

To launch the GUI application, run the following command:

```
python src/gui.py
```

The GUI allows you to:

- Select a date range for retrieving meetings.
- Filter meetings by status and categories.
- Export the results to an Excel file.
- View the results directly within the application.

### Running the Script Programmatically

To use the script programmatically, you can call the `get_meetings` function from `outlookmeeting.py`. Example:

```python
from outlookmeeting import get_meetings
import datetime

start_date = datetime.datetime(2023, 1, 1)
end_date = datetime.datetime(2023, 1, 31)
download_folder = "path/to/output/folder"
meeting_types = [0, 1]  # Normal and Meeting
progress_callback = lambda progress: print(f"Progress: {progress}%")
category_filter = "Work, Personal"
exclude = False

file_path = get_meetings(start_date, end_date, download_folder, meeting_types, progress_callback, category_filter, exclude)
print(f"Meetings exported to: {file_path}")
```

## Testing

To run the tests, navigate to the project directory and execute:

```
pytest tests/
```

## License

This project is licensed under the MIT License. See the LICENSE file for details.
