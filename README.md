# Outlook Meeting Project

This project is designed to interact with Microsoft Outlook's calendar to extract meeting details and prepare data for output. It provides functionality to retrieve meetings within a specified date range, categorize them by month, and calculate total durations.

## Project Structure

```
outlook-meeting-project
├── outlookmeeting
│   ├── __init__.py
│   ├── outlookmeeting.py
├── tests
│   └── test_outlookmeeting.py
├── requirements.txt
├── .gitignore
└── README.md
```

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

## Usage

To run the main functionality of the project, execute the `outlookmeeting.py` script. This script will connect to Outlook, retrieve meeting data, and prepare it for output.

## Testing

Unit tests for the project are located in the `tests` directory. To run the tests, use the following command:
```
pytest tests/test_outlookmeeting.py
```

## Dependencies

This project requires the following Python packages:

- `pandas`
- `pywin32`

Make sure to install these packages using the `requirements.txt` file.

## Contributing

Contributions are welcome! Please submit a pull request or open an issue for any enhancements or bug fixes.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.