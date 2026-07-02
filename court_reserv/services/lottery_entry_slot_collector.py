# -*- coding: utf-8 -*-
"""Collect visible lottery-entry slots from the current page."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from court_reserv.models import LotteryEntrySlot


WEEKDAY_LABELS = ["月", "火", "水", "木", "金", "土", "日"]


class LotteryEntrySlotCollector:
    """Collect and filter currently visible lottery-entry slots."""

    def __init__(self, navigation_service, browser_session):
        self.navigation_service = navigation_service
        self.browser_session = browser_session

    def collect_slots(self, driver, target_weekdays=None):
        """Collect visible slots and filter by target weekdays."""
        if not driver:
            raise RuntimeError("Driver not started")

        wait = self.browser_session.get_wait(driver, 10)
        wait.until(lambda d: d.find_element("id", "usedate-table"))

        payload = self._collect_raw_payload(driver)
        target_labels = [str(label) for label in (target_weekdays or ["土"])]
        slots = []
        for item in payload.get("slots", []):
            ymd = str(item.get("ymd", "")).strip()
            date = self._ymd_to_date(ymd)
            weekday = self._weekday_label(ymd)
            if weekday not in target_labels:
                continue
            start_time = self._format_hhmm(item.get("stime", ""))
            end_time = self._format_hhmm(item.get("etime", ""))
            time_range = (
                f"{start_time}-{end_time}" if start_time and end_time else ""
            )
            field_number = str(item.get("field", "")).strip()
            applied_count = self._to_int(item.get("appliedCount"))
            available_count = self._to_int(item.get("availableCount"))
            raw_text = " ".join(
                part
                for part in (
                    ymd,
                    item.get("timeLabel", ""),
                    f"{item.get('stime', '')}-{item.get('etime', '')}",
                    f"fields:{field_number}" if field_number else "",
                    f"applied:{applied_count}" if applied_count is not None else "",
                )
                if part
            ).strip()
            slots.append(
                LotteryEntrySlot(
                    park_name=str(payload.get("park_name", "")).strip(),
                    facility_name=str(payload.get("facility_name", "")).strip(),
                    date=date,
                    weekday=weekday,
                    time_range=time_range,
                    start_time=start_time,
                    end_time=end_time,
                    field_number=field_number,
                    available_count=available_count,
                    applied_count=applied_count,
                    raw_text=raw_text,
                )
            )
        return slots

    def collect_slots_for_entries(
        self,
        driver,
        target_entries=None,
        target_weekdays=None,
        max_weeks=8,
    ):
        """Collect slots across weekly pages until target entries are found or the limit is hit."""
        target_entries = [entry for entry in (target_entries or []) if isinstance(entry, dict)]
        desired_keys = {
            self._entry_key(
                date=str(entry.get("date", "")).strip(),
                time_range=str(entry.get("time_range", "")).strip(),
                facility=str(entry.get("facility", "")).strip(),
            )
            for entry in target_entries
            if entry.get("date") and entry.get("time_range")
        }
        unique_slots = {}
        matched_keys = set()
        weekly_summaries = []
        stopped_reason = "max_weeks_reached"

        for week_index in range(max(1, int(max_weeks or 8))):
            slots = self.collect_slots(driver, target_weekdays=target_weekdays)
            for slot in slots:
                unique_slots[self._slot_identity(slot)] = slot
                if self._slot_matches_keys(slot, desired_keys):
                    matched_keys.add(
                        self._entry_key(
                            date=slot.date,
                            time_range=slot.time_range,
                            facility=self._slot_facility_label(slot),
                        )
                    )

            weekly_summaries.append(
                {
                    "week_index": week_index + 1,
                    "slot_count": len(slots),
                    "dates": sorted({slot.date for slot in slots if slot.date}),
                    "matched_target_count": len(matched_keys),
                }
            )
            self._save_week_debug(driver, week_index + 1)

            if desired_keys and desired_keys.issubset(matched_keys):
                stopped_reason = "all_targets_found"
                break

            if week_index + 1 >= max(1, int(max_weeks or 8)):
                stopped_reason = "max_weeks_reached"
                break

            next_week_result = self.navigation_service.go_to_lottery_next_week(driver)
            if not next_week_result.get("changed"):
                stopped_reason = "next_week_not_available"
                break

        return {
            "slots": list(unique_slots.values()),
            "weekly_summaries": weekly_summaries,
            "weeks_explored": len(weekly_summaries),
            "max_weeks": max(1, int(max_weeks or 8)),
            "matched_target_count": len(matched_keys),
            "target_count": len(desired_keys),
            "stopped_reason": stopped_reason,
        }

    def _collect_raw_payload(self, driver):
        script = """
        const payload = {
          park_name: "",
          facility_name: "",
          slots: []
        };
        try {
          const parkSelect = document.getElementById("bname");
          if (parkSelect && parkSelect.selectedIndex >= 0) {
            payload.park_name = parkSelect.options[parkSelect.selectedIndex].text.trim();
          }
          const facilitySelect = document.getElementById("iname");
          if (facilitySelect && facilitySelect.selectedIndex >= 0) {
            payload.facility_name = facilitySelect.options[facilitySelect.selectedIndex].text.trim();
          }
          const headers = Array.from(
            document.querySelectorAll('#usedate-table thead input[name="selectUseYMD"]')
          ).map((input) => input.value);
          const rows = Array.from(document.querySelectorAll("#usedate-table tbody tr"));
          rows.forEach((row) => {
            const timeHeader = row.querySelector("th");
            const timeLabel = timeHeader ? timeHeader.textContent.trim() : "";
            const cells = Array.from(row.querySelectorAll("td"));
            cells.forEach((td, idx) => {
              const stime = td.querySelector('input[name="selectStime"]');
              const etime = td.querySelector('input[name="selectEtime"]');
              if (!stime || !etime) {
                return;
              }
              const field = td.querySelector('input[name="selectField"]');
              const numbers = Array.from(td.childNodes)
                .map((node) => (node.textContent || "").trim())
                .filter(Boolean);
              const appliedNode = td.querySelector("span.font-weight-bold");
              const appliedCount = appliedNode ? appliedNode.textContent.trim() : "";
              let availableCount = "";
              if (numbers.length > 0) {
                availableCount = numbers[0];
              }
              payload.slots.push({
                ymd: headers[idx] || "",
                timeLabel: timeLabel,
                stime: stime.value || "",
                etime: etime.value || "",
                field: field ? field.value || "" : "",
                availableCount: availableCount,
                appliedCount: appliedCount,
              });
            });
          });
        } catch (error) {
          payload.error = String(error);
        }
        return JSON.stringify(payload);
        """
        raw = self.navigation_service.execute_script(driver, script)
        return json.loads(raw) if raw else {"slots": []}

    def _ymd_to_date(self, ymd):
        try:
            return datetime.strptime(str(ymd), "%Y%m%d").strftime("%Y-%m-%d")
        except Exception:
            return str(ymd)

    def _weekday_label(self, ymd):
        try:
            return WEEKDAY_LABELS[datetime.strptime(str(ymd), "%Y%m%d").weekday()]
        except Exception:
            return ""

    def _format_hhmm(self, value):
        text = str(value).strip()
        if len(text) == 3 and text.isdigit():
            return f"0{text[0]}:{text[1:]}"
        if len(text) == 4 and text.isdigit():
            return f"{text[:2]}:{text[2:]}"
        return text

    def _to_int(self, value):
        text = str(value).strip()
        if not text:
            return None
        try:
            return int(text)
        except Exception:
            return None

    def _slot_identity(self, slot):
        return (
            slot.date,
            slot.time_range,
            slot.park_name,
            slot.facility_name,
            slot.field_number,
        )

    def _slot_facility_label(self, slot):
        return " ".join(
            part for part in (slot.park_name, slot.facility_name) if part
        ).strip()

    def _entry_key(self, date, time_range, facility):
        return (
            str(date).strip(),
            str(time_range).strip(),
            str(facility).strip(),
        )

    def _slot_matches_keys(self, slot, desired_keys):
        if not desired_keys:
            return False
        slot_label = self._slot_facility_label(slot)
        exact_key = self._entry_key(slot.date, slot.time_range, slot_label)
        if exact_key in desired_keys:
            return True
        fallback_keys = {
            self._entry_key(slot.date, slot.time_range, ""),
            self._entry_key(slot.date, slot.time_range, slot.park_name),
            self._entry_key(slot.date, slot.time_range, slot.facility_name),
        }
        return any(key in desired_keys for key in fallback_keys)

    def _save_week_debug(self, driver, week_index):
        save_html = getattr(self.navigation_service, "save_debug_html", None)
        save_dom_summary = getattr(self.navigation_service, "save_dom_summary", None)
        if callable(save_html):
            save_html(driver, f"lottery_entry_week_{week_index:02d}.html")
        if callable(save_dom_summary):
            save_dom_summary(
                driver,
                f"lottery_entry_week_{week_index:02d}_dom_summary.json",
            )
