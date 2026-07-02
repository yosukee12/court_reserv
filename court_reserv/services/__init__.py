"""Service layer boundary.

This package is reserved for reservation, lottery, and vacancy-search
business logic that should eventually be separated from the UI and browser
layers.

Issue 0003 only documents the boundary and does not move existing logic.
"""

from .lottery import LotteryService
from .reservation import ReservationService
from .availability import AvailabilityService
from .id_manager import IdManagerService
from .lottery_automation import LotteryAutomationDryRunService, SlotCollectionAdapter
from .auto_reservation import AutoReservationDryRunService
from .slot_ranking import RankedSlot, SlotRankingService

__all__ = [
    "LotteryService",
    "ReservationService",
    "AvailabilityService",
    "IdManagerService",
    "SlotCollectionAdapter",
    "SlotRankingService",
    "RankedSlot",
    "LotteryAutomationDryRunService",
    "AutoReservationDryRunService",
]
