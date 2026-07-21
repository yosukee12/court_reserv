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
    preferred_court_names: list[str] = field(default_factory=list)
    max_candidates: int = 10
    lottery_target_weekdays: list[str] = field(default_factory=lambda: ["土"])
    lottery_default_entries: list[dict] = field(default_factory=list)
    lottery_account_overrides: dict[str, list[dict]] = field(default_factory=dict)
    lottery_max_entries_per_account: int = 2
    lottery_search_weeks: int = 8
    lottery_dry_run: bool = True
    lottery_manual_final_submit: bool = False
    lottery_manual_preconfirm_submit: bool = False
    lottery_reuse_browser_session: bool = False
    lottery_human_sleep_enabled: bool = False
    lottery_human_sleep_min: float = 0.5
    lottery_human_sleep_max: float = 1.0
