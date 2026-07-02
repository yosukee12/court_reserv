# -*- coding: utf-8 -*-
"""Facility model foundation."""

from dataclasses import dataclass


@dataclass(frozen=True)
class Facility:
    """Represents a sport facility in a park."""

    park_id: str
    park_name: str
    facility_id: str
    facility_name: str
    sport_id: str = "130"
    sport_name: str = "テニス"
