# -*- coding: utf-8 -*-
"""Dry-run lottery automation core for Phase 2."""

from __future__ import annotations

import csv
import json
import logging
import re
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

from court_reserv.models import Slot
from court_reserv.services.availability import AvailabilityService
from court_reserv.services.slot_ranking import RankedSlot, SlotRankingService


class SlotCollectionAdapter:
    """Bridge existing availability outputs into Slot models."""

    def __init__(self, availability_service: AvailabilityService | None = None):
        self.availability_service = availability_service

    def collect_from_service(
        self,
        driver,
        weeks_limit: int = 8,
        only_weekday: int | None = None,
    ) -> list[Slot]:
        if self.availability_service is None:
            raise ValueError("AvailabilityService is required for live collection")
        raw_slots = self.availability_service.collect_all_available_slots(
            driver,
            weeks_limit=weeks_limit,
            only_weekday=only_weekday,
        )
        return self.convert_raw_slots(raw_slots)

    def load_from_csv(self, csv_path: str | Path) -> list[Slot]:
        rows = []
        with open(csv_path, newline="", encoding="utf-8") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                slot_text = (row.get("slot") or "").strip()
                if slot_text:
                    rows.append(slot_text)
        return self.convert_raw_slots(rows)

    def find_latest_slots_csv(self, search_dirs: list[Path]) -> Path | None:
        candidates = []
        for directory in search_dirs:
            if not directory.exists():
                continue
            for path in directory.glob("available_slots_*.csv"):
                if path.is_file():
                    candidates.append(path)
        if not candidates:
            return None
        return sorted(candidates, key=lambda path: path.stat().st_mtime, reverse=True)[0]

    def convert_raw_slots(self, raw_slots: list[str]) -> list[Slot]:
        return [self._parse_slot(raw_text) for raw_text in raw_slots if raw_text]

    def _parse_slot(self, raw_text: str) -> Slot:
        date = ""
        time_range = ""
        court_name = None
        applied_count = None

        date_match = re.search(r"(?P<ymd>\d{8})", raw_text)
        if date_match:
            ymd = date_match.group("ymd")
            try:
                date = datetime.strptime(ymd, "%Y%m%d").strftime("%Y-%m-%d")
            except Exception:
                date = ymd

        time_match = re.search(r"(?P<start>\d{4})-(?P<end>\d{4})", raw_text)
        if time_match:
            time_range = (
                f"{time_match.group('start')[0:2]}:{time_match.group('start')[2:4]}-"
                f"{time_match.group('end')[0:2]}:{time_match.group('end')[2:4]}"
            )
        else:
            label_match = re.search(
                r"(?P<start>\d{1,2}:\d{2})-(?P<end>\d{1,2}:\d{2})",
                raw_text,
            )
            if label_match:
                time_range = f"{label_match.group('start')}-{label_match.group('end')}"

        field_match = re.search(r"fields:(?P<field>[^\s]+)", raw_text)
        if field_match:
            court_name = field_match.group("field")

        applied_match = re.search(r"applied:(?P<count>[^\s]+)", raw_text)
        if applied_match:
            applied_count = applied_match.group("count")

        return Slot(
            date=date,
            time_range=time_range,
            court_name=court_name,
            applied_count=applied_count,
            raw_text=raw_text,
        )


class LotteryAutomationDryRunService:
    """Coordinate preference loading, candidate collection, ranking, and output."""

    def __init__(
        self,
        slot_adapter: SlotCollectionAdapter,
        slot_ranking_service: SlotRankingService,
        logger=None,
    ):
        self.slot_adapter = slot_adapter
        self.slot_ranking_service = slot_ranking_service
        self.logger = logger or logging.getLogger(__name__)

    def run(
        self,
        preference,
        source_csv: str | Path | None = None,
        search_dirs: list[Path] | None = None,
    ) -> dict:
        resolved_csv = None
        if source_csv is not None:
            resolved_csv = Path(source_csv)
        elif search_dirs:
            resolved_csv = self.slot_adapter.find_latest_slots_csv(search_dirs)

        if resolved_csv is None:
            slots = []
        else:
            slots = self.slot_adapter.load_from_csv(resolved_csv)

        ranked_slots = self.slot_ranking_service.rank_slots(slots, preference)
        max_candidates = max(1, int(getattr(preference, "max_candidates", 10)))
        top_candidates = ranked_slots[:max_candidates]

        return {
            "source_csv": str(resolved_csv) if resolved_csv else None,
            "total_slots": len(slots),
            "ranked_candidates": top_candidates,
        }

    def print_result(self, result: dict) -> None:
        source_csv = result.get("source_csv")
        if source_csv:
            print(f"Candidate source: {source_csv}")
        print(f"Total candidates: {result.get('total_slots', 0)}")

        ranked_candidates: list[RankedSlot] = result.get("ranked_candidates", [])
        if not ranked_candidates:
            print("No ranked lottery candidates found.")
            return

        print("Ranked lottery candidates:")
        for index, ranked in enumerate(ranked_candidates, start=1):
            slot = ranked.slot
            reasons = ", ".join(ranked.reasons) if ranked.reasons else "no preference match"
            label = slot.raw_text or f"{slot.date} {slot.time_range}"
            print(f"{index}. score={ranked.score} slot={label} reasons={reasons}")

    def save_result(
        self,
        result: dict,
        output_dir: str | Path,
    ) -> tuple[Path, Path]:
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_path = output_path / f"lottery_automation_dry_run_{timestamp}.json"
        csv_path = output_path / f"lottery_automation_dry_run_{timestamp}.csv"

        ranked_candidates: list[RankedSlot] = result.get("ranked_candidates", [])
        payload = {
            "source_csv": result.get("source_csv"),
            "total_slots": result.get("total_slots", 0),
            "ranked_candidates": [
                {
                    "score": ranked.score,
                    "reasons": ranked.reasons,
                    "slot": asdict(ranked.slot),
                }
                for ranked in ranked_candidates
            ],
        }

        json_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        with open(csv_path, "w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(
                handle,
                fieldnames=[
                    "rank",
                    "score",
                    "date",
                    "time_range",
                    "court_name",
                    "applied_count",
                    "raw_text",
                    "reasons",
                ],
            )
            writer.writeheader()
            for index, ranked in enumerate(ranked_candidates, start=1):
                writer.writerow(
                    {
                        "rank": index,
                        "score": ranked.score,
                        "date": ranked.slot.date,
                        "time_range": ranked.slot.time_range,
                        "court_name": ranked.slot.court_name,
                        "applied_count": ranked.slot.applied_count,
                        "raw_text": ranked.slot.raw_text,
                        "reasons": " | ".join(ranked.reasons),
                    }
                )

        return json_path, csv_path
