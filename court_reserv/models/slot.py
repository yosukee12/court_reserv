# -*- coding: utf-8 -*-
"""Reservation slot model foundation."""

from dataclasses import dataclass

from .facility import Facility


@dataclass(frozen=True)
class Slot:
    """Represents an availability or reservation slot."""

    date: str
    time_range: str
    facility: Facility | None = None
    court_name: str | None = None
    status: str | None = None
    applied_count: str | None = None
