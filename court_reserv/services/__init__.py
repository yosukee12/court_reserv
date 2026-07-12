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
from .lottery_entry_workflow import LotteryEntryWorkflowService
from .lottery_application_check_workflow import LotteryApplicationCheckWorkflowService
from .lottery_result_workflow import LotteryResultWorkflowService
from .reservation_confirmation_workflow import ReservationConfirmationWorkflowService
from .reservation_status_workflow import ReservationStatusWorkflowService
from .lottery_entry_slot_collector import LotteryEntrySlotCollector

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
    "LotteryEntryWorkflowService",
    "LotteryApplicationCheckWorkflowService",
    "LotteryResultWorkflowService",
    "ReservationConfirmationWorkflowService",
    "ReservationStatusWorkflowService",
    "LotteryEntrySlotCollector",
]
