"""UI layer boundary.

This package is reserved for Tkinter widgets, view composition, and user
interaction handling.

Issue 0003 keeps the current Tkinter implementation in court_reserv.py and
only documents the future separation point.
"""

from .app import main

__all__ = ["main"]
