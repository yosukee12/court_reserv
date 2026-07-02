# -*- coding: utf-8 -*-
"""Preference configuration loaders for Phase 2 dry-run automation."""

from __future__ import annotations

import json
from pathlib import Path

from court_reserv.models import Facility, ReservationPreference


def _strip_quotes(value: str) -> str:
    text = value.strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in ("'", '"'):
        return text[1:-1]
    return text


def _parse_scalar(value: str):
    text = _strip_quotes(value)
    if text.lower() in {"true", "false"}:
        return text.lower() == "true"
    if text.isdigit():
        return int(text)
    return text


def _load_simple_yaml(text: str) -> dict:
    """Parse a small YAML subset used by the sample preference file."""

    result: dict = {}
    current_key: str | None = None
    current_list: list | None = None
    current_item: dict | None = None

    for raw_line in text.splitlines():
        line = raw_line.rstrip()
        if not line.strip() or line.lstrip().startswith("#"):
            continue

        indent = len(raw_line) - len(raw_line.lstrip(" "))
        stripped = line.strip()

        if indent == 0 and ":" in stripped:
            key, value = stripped.split(":", 1)
            key = key.strip()
            value = value.strip()
            current_key = key
            current_item = None
            if not value:
                current_list = []
                result[key] = current_list
            else:
                current_list = None
                result[key] = _parse_scalar(value)
            continue

        if current_key is None or current_list is None:
            raise ValueError("Unsupported YAML structure in preferences file")

        if stripped.startswith("- "):
            item_text = stripped[2:].strip()
            if ":" in item_text:
                item_key, item_value = item_text.split(":", 1)
                current_item = {item_key.strip(): _parse_scalar(item_value.strip())}
                current_list.append(current_item)
            else:
                current_item = None
                current_list.append(_parse_scalar(item_text))
            continue

        if current_item is not None and ":" in stripped:
            item_key, item_value = stripped.split(":", 1)
            current_item[item_key.strip()] = _parse_scalar(item_value.strip())
            continue

        raise ValueError("Unsupported YAML structure in preferences file")

    return result


def load_preferences_data(path: str | Path) -> dict:
    prefs_path = Path(path)
    text = prefs_path.read_text(encoding="utf-8")
    suffix = prefs_path.suffix.lower()

    if suffix == ".json":
        return json.loads(text)

    if suffix in {".yaml", ".yml"}:
        try:
            import yaml  # type: ignore
        except Exception:
            yaml = None
        if yaml is not None:
            data = yaml.safe_load(text) or {}
            if not isinstance(data, dict):
                raise ValueError("Preference file must contain a mapping at the top level")
            return data
        return _load_simple_yaml(text)

    raise ValueError(f"Unsupported preference file format: {prefs_path}")


def load_reservation_preference(path: str | Path) -> ReservationPreference:
    data = load_preferences_data(path)

    preferred_facilities = []
    for item in data.get("preferred_facilities", []):
        if not isinstance(item, dict):
            continue
        preferred_facilities.append(
            Facility(
                park_id=str(item.get("park_id", "")),
                park_name=str(item.get("park_name", "")),
                facility_id=str(item.get("facility_id", "")),
                facility_name=str(item.get("facility_name", "")),
                sport_id=str(item.get("sport_id", "130")),
                sport_name=str(item.get("sport_name", "テニス")),
            )
        )

    max_candidates = data.get("max_candidates", 10)
    if not isinstance(max_candidates, int):
        try:
            max_candidates = int(max_candidates)
        except Exception:
            max_candidates = 10

    return ReservationPreference(
        preferred_weekdays=[str(item) for item in data.get("preferred_weekdays", [])],
        preferred_time_ranges=[
            str(item) for item in data.get("preferred_time_ranges", [])
        ],
        preferred_facilities=preferred_facilities,
        preferred_court_names=[
            str(item) for item in data.get("preferred_court_names", [])
        ],
        max_candidates=max_candidates,
    )
