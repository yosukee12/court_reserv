"""Service layer boundary.

This package is reserved for reservation, lottery, and vacancy-search
business logic that should eventually be separated from the UI and browser
layers.

Issue 0003 only documents the boundary and does not move existing logic.
"""

from .lottery import LotteryService
from .reservation import ReservationService

__all__ = ["LotteryService", "ReservationService"]
