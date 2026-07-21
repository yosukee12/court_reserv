# -*- coding: utf-8 -*-
"""Lottery entry slot model."""

from dataclasses import dataclass


@dataclass(frozen=True)
class LotteryEntrySlot:
    """Represents one selectable lottery-entry slot on the current page."""

    park_name: str
    facility_name: str
    date: str
    weekday: str
    time_range: str
    start_time: str
    end_time: str
    field_number: str
    available_count: int | None = None
    applied_count: int | None = None
    raw_text: str | None = None
