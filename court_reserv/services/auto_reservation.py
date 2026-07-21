# -*- coding: utf-8 -*-
"""Compatibility shim for the renamed lottery automation module."""

from .lottery_automation import LotteryAutomationDryRunService, SlotCollectionAdapter

AutoReservationDryRunService = LotteryAutomationDryRunService

__all__ = ["SlotCollectionAdapter", "LotteryAutomationDryRunService", "AutoReservationDryRunService"]
