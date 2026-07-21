"""Browser layer boundary.

This package is reserved for Selenium driver operations, login helpers,
page transitions, and DOM interaction wrappers.

Issue 0003 only defines the responsibility boundary. Existing Selenium
implementations remain in court_reserv.py until a later Issue moves them.
"""

from .login import LoginService
from .navigation import NavigationService
from .session import BrowserSession

__all__ = ["BrowserSession", "LoginService", "NavigationService"]
