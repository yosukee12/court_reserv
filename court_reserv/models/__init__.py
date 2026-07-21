"""Model layer boundary.

This package is reserved for user, facility, reservation slot, and result
data structures.

Issue 0003 only defines the responsibility boundary without changing
runtime behavior.
"""

from .account import Account
from .facility import Facility
from .lottery_entry_slot import LotteryEntrySlot
from .preference import ReservationPreference
from .slot import Slot

__all__ = [
    "Account",
    "Facility",
    "Slot",
    "ReservationPreference",
    "LotteryEntrySlot",
]
