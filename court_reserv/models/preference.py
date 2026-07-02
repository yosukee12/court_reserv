# -*- coding: utf-8 -*-
"""Reservation preference model foundation."""

from dataclasses import dataclass, field

from .facility import Facility


@dataclass(frozen=True)
class ReservationPreference:
    """Represents lightweight future reservation preferences."""

    preferred_weekdays: list[str] = field(default_factory=list)
    preferred_time_ranges: list[str] = field(default_factory=list)
    preferred_facilities: list[Facility] = field(default_factory=list)
