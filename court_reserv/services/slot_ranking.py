# -*- coding: utf-8 -*-
"""Slot ranking helpers for Phase 2 dry-run automation."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

from court_reserv.models import ReservationPreference, Slot


WEEKDAY_LABELS = ["月", "火", "水", "木", "金", "土", "日"]


@dataclass(frozen=True)
class RankedSlot:
    """Represents a ranked slot candidate for dry-run output."""

    slot: Slot
    score: int
    reasons: list[str] = field(default_factory=list)


class SlotRankingService:
    """Score and sort slot candidates with a simple additive policy."""

    def rank_slots(
        self,
        slots: list[Slot],
        preference: ReservationPreference,
    ) -> list[RankedSlot]:
        ranked = [self._score_slot(slot, preference) for slot in slots]
        return sorted(
            ranked,
            key=lambda item: (
                -item.score,
                item.slot.date or "9999-99-99",
                item.slot.time_range or "99:99-99:99",
                item.slot.raw_text or "",
            ),
        )

    def _score_slot(
        self,
        slot: Slot,
        preference: ReservationPreference,
    ) -> RankedSlot:
        score = 0
        reasons = []

        if self._matches_facility(slot, preference):
            score += 50
            reasons.append("facility match +50")

        if self._matches_weekday(slot, preference):
            score += 30
            reasons.append("weekday match +30")

        if self._matches_time_range(slot, preference):
            score += 20
            reasons.append("time range match +20")

        date_bonus = self._earlier_date_bonus(slot)
        if date_bonus:
            score += date_bonus
            reasons.append(f"earlier date +{date_bonus}")

        return RankedSlot(slot=slot, score=score, reasons=reasons)

    def _matches_facility(
        self,
        slot: Slot,
        preference: ReservationPreference,
    ) -> bool:
        haystacks = [
            slot.raw_text or "",
            slot.court_name or "",
            slot.facility.park_name if slot.facility else "",
            slot.facility.facility_name if slot.facility else "",
        ]
        searchable = " ".join(text for text in haystacks if text)

        for facility in preference.preferred_facilities:
            if facility.park_name and facility.park_name in searchable:
                return True
            if facility.facility_name and facility.facility_name in searchable:
                return True
            if facility.facility_id and facility.facility_id in searchable:
                return True

        for court_name in preference.preferred_court_names:
            if court_name and court_name in searchable:
                return True

        return False

    def _matches_weekday(
        self,
        slot: Slot,
        preference: ReservationPreference,
    ) -> bool:
        if not preference.preferred_weekdays:
            return False
        try:
            weekday = WEEKDAY_LABELS[datetime.strptime(slot.date, "%Y-%m-%d").weekday()]
        except Exception:
            return False
        return weekday in preference.preferred_weekdays

    def _matches_time_range(
        self,
        slot: Slot,
        preference: ReservationPreference,
    ) -> bool:
        if not preference.preferred_time_ranges or not slot.time_range:
            return False
        return slot.time_range in preference.preferred_time_ranges

    def _earlier_date_bonus(self, slot: Slot) -> int:
        try:
            days_ahead = (datetime.strptime(slot.date, "%Y-%m-%d").date() - datetime.today().date()).days
        except Exception:
            return 0
        if days_ahead < 0:
            return 0
        return max(0, 10 - min(days_ahead, 10))
