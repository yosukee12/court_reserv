# -*- coding: utf-8 -*-
"""Build FFTC wiki text from a reservation-confirmed account CSV."""

from __future__ import annotations

import csv
import datetime
import re
from collections import defaultdict
from pathlib import Path


class FftcWikiTextService:
    """Create the fixed-format FFTC wiki text used by the GUI workflow."""

    PARK_NAME = "府中の森公園"
    COURT_NAME = "オムニコート"
    NOTE = "※受付時に当選者と受付者(当選者が本人でない場合)の利用者番号が必要になりました。"

    _DATE_PATTERN = re.compile(
        r"(?:(?P<year>\d{4})年)?\s*(?P<month>\d{1,2})月(?P<day>\d{1,2})日"
        r"(?:\s*[（(](?P<weekday>[月火水木金土日])(?:曜|曜日)?[）)])?"
    )
    _TIME_PATTERN = re.compile(
        r"(?P<start_hour>\d{1,2})時(?P<start_minute>\d{2})分\s*[～〜-]\s*"
        r"(?P<end_hour>\d{1,2})時(?P<end_minute>\d{2})分"
    )

    def build_from_csv(self, csv_path):
        """Return wiki text generated from a reservation-confirmed CSV."""
        grouped = defaultdict(lambda: defaultdict(list))
        with Path(csv_path).open(newline="", encoding="utf-8-sig") as source:
            for row in csv.reader(source):
                if len(row) < 5 or not re.fullmatch(r"\d{8}", row[0].strip()):
                    continue
                account_name = row[1].strip()
                user_id = row[0].strip()
                current_date = None
                for cell in row[4:]:
                    date_match = self._DATE_PATTERN.search(cell)
                    if date_match:
                        year_match = re.search(r"(\d{4})年", cell)
                        current_date = self._parse_date(
                            date_match,
                            year=int(year_match.group(1)) if year_match else None,
                        )
                    if current_date is None:
                        continue
                    for time_match in self._TIME_PATTERN.finditer(cell):
                        time_range = self._format_time(time_match)
                        grouped[current_date][time_range].append(
                            f"{account_name}({user_id})"
                        )

        if not grouped:
            return ""

        lines = []
        date_values = sorted(grouped)
        for index, date_value in enumerate(date_values):
            time_groups = grouped[date_value]
            time_parts = []
            names_by_time = []
            for time_range in sorted(time_groups, key=self._time_sort_key):
                names = time_groups[time_range]
                time_parts.append(f"{time_range} {len(names)}面")
                names_by_time.append("、".join(names))
            date_text = self._format_date(date_value)
            lines.append(
                f"**{date_text} {self.PARK_NAME}　{self.COURT_NAME} "
                f"{' '.join(time_parts)}　当選者名：{'//'.join(names_by_time)}"
            )
            lines.extend([self.NOTE, "", "-参加予定："])
            if index < len(date_values) - 1:
                lines.append("")
        return "\n".join(lines)

    @classmethod
    def _parse_date(cls, match, year=None):
        year = year or int(match.group("year") or datetime.date.today().year)
        month = int(match.group("month"))
        day = int(match.group("day"))
        return datetime.date(year, month, day)

    @staticmethod
    def _format_date(value):
        weekdays = "月火水木金土日"
        return f"{value.month}/{value.day}({weekdays[value.weekday()]})"

    @classmethod
    def _format_time(cls, match):
        return (
            f"{int(match.group('start_hour'))}:{match.group('start_minute')}～"
            f"{int(match.group('end_hour'))}:{match.group('end_minute')}"
        )

    @staticmethod
    def _time_sort_key(time_range):
        match = re.match(r"(\d+):(\d+)", time_range)
        return int(match.group(1)) * 60 + int(match.group(2))
