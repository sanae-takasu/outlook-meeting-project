import unittest
from datetime import datetime, timedelta
from outlookmeeting.outlookmeeting import get_meetings_by_month  # Assuming this function exists

class TestOutlookMeeting(unittest.TestCase):

    def setUp(self):
        # Setup code to initialize any required variables or mock objects
        self.start_date = datetime(2023, 4, 1)
        self.end_date = datetime(2023, 9, 30)

    def test_meeting_count(self):
        meetings = get_meetings_by_month(self.start_date, self.end_date)
        self.assertIsInstance(meetings, dict)
        self.assertGreater(len(meetings), 0)

    def test_meeting_details(self):
        meetings = get_meetings_by_month(self.start_date, self.end_date)
        for month, details in meetings.items():
            for subject, info in details.items():
                self.assertIn("count", info)
                self.assertIn("total_duration", info)
                self.assertIsInstance(info["total_duration"], timedelta)

if __name__ == '__main__':
    unittest.main()