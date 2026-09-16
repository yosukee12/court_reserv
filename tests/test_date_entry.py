import unittest
from datetime import date

from court_reserv.ui.date_entry import adjacent_month, parse_entry_date


class DateEntryTests(unittest.TestCase):
    def test_valid_date_and_whitespace(self):
        self.assertEqual(parse_entry_date(" 2028-02-29 "), "2028-02-29")

    def test_invalid_or_incomplete_dates(self):
        for value in ("", "2026-", "2026-02-29", "2026-04-31", "2026-13-01", "20261010"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_entry_date(value)

    def test_month_navigation_across_year(self):
        self.assertEqual(adjacent_month(date(2026, 12, 1), 1), date(2027, 1, 1))
        self.assertEqual(adjacent_month(date(2026, 1, 1), -1), date(2025, 12, 1))

    def test_month_navigation_from_last_day(self):
        self.assertEqual(adjacent_month(date(2028, 1, 31), 1), date(2028, 2, 1))


if __name__ == "__main__":
    unittest.main()
