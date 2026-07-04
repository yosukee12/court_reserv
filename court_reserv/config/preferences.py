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

    lines = []
    for raw_line in text.splitlines():
        if not raw_line.strip() or raw_line.lstrip().startswith("#"):
            continue
        indent = len(raw_line) - len(raw_line.lstrip(" "))
        lines.append((indent, raw_line.strip()))

    if not lines:
        return {}

    def parse_block(index, indent):
        container = None
        result_dict = {}
        result_list = []

        while index < len(lines):
            current_indent, stripped = lines[index]
            if current_indent < indent:
                break
            if current_indent > indent:
                raise ValueError("Unsupported YAML indentation in preferences file")

            if stripped.startswith("- "):
                if container is None:
                    container = "list"
                elif container != "list":
                    raise ValueError("Mixed YAML container types are not supported")

                item_text = stripped[2:].strip()
                if not item_text:
                    item_value, index = parse_block(index + 1, indent + 2)
                    result_list.append(item_value)
                    continue

                if ":" in item_text:
                    item_key, item_value_text = item_text.split(":", 1)
                    item_key = item_key.strip()
                    item_value_text = item_value_text.strip()
                    item_dict = {}
                    if item_value_text:
                        item_dict[item_key] = _parse_scalar(item_value_text)
                        index += 1
                    else:
                        nested_value, index = parse_block(index + 1, indent + 2)
                        item_dict[item_key] = nested_value
                    while index < len(lines):
                        child_indent, child_text = lines[index]
                        if child_indent < indent + 2:
                            break
                        if child_indent > indent + 2:
                            raise ValueError(
                                "Unsupported nested YAML indentation in preferences file"
                            )
                        if child_text.startswith("- "):
                            break
                        child_key, child_value_text = child_text.split(":", 1)
                        child_key = child_key.strip()
                        child_value_text = child_value_text.strip()
                        if child_value_text:
                            item_dict[child_key] = _parse_scalar(child_value_text)
                            index += 1
                        else:
                            nested_value, index = parse_block(index + 1, child_indent + 2)
                            item_dict[child_key] = nested_value
                    result_list.append(item_dict)
                    continue

                result_list.append(_parse_scalar(item_text))
                index += 1
                continue

            if container is None:
                container = "dict"
            elif container != "dict":
                raise ValueError("Mixed YAML container types are not supported")

            key, value_text = stripped.split(":", 1)
            key = key.strip()
            value_text = value_text.strip()
            if value_text:
                result_dict[key] = _parse_scalar(value_text)
                index += 1
            else:
                nested_value, index = parse_block(index + 1, indent + 2)
                result_dict[key] = nested_value

        if container == "list":
            return result_list, index
        return result_dict, index

    data, _ = parse_block(0, 0)
    if not isinstance(data, dict):
        raise ValueError("Preference file must contain a mapping at the top level")
    return data


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
    lottery_data = data.get("lottery", {})
    if not isinstance(lottery_data, dict):
        lottery_data = {}

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

    target_weekdays = lottery_data.get(
        "target_weekdays",
        data.get("preferred_weekdays", ["土"]),
    )
    if not isinstance(target_weekdays, list):
        target_weekdays = ["土"]

    max_entries_per_account = lottery_data.get("max_entries_per_account", 2)
    if not isinstance(max_entries_per_account, int):
        try:
            max_entries_per_account = int(max_entries_per_account)
        except Exception:
            max_entries_per_account = 2

    search_weeks = lottery_data.get("search_weeks", 8)
    if not isinstance(search_weeks, int):
        try:
            search_weeks = int(search_weeks)
        except Exception:
            search_weeks = 8

    dry_run = lottery_data.get("dry_run", True)
    if not isinstance(dry_run, bool):
        dry_run = bool(dry_run)

    default_entries = lottery_data.get("default_entries", [])
    if not isinstance(default_entries, list):
        default_entries = []

    raw_account_overrides = lottery_data.get("account_overrides", {})
    if not isinstance(raw_account_overrides, dict):
        raw_account_overrides = {}
    account_overrides = {}
    for account_key, override_value in raw_account_overrides.items():
        if isinstance(override_value, dict):
            entries = override_value.get("entries", [])
        else:
            entries = []
        if isinstance(entries, list):
            account_overrides[_strip_quotes(str(account_key))] = entries

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
        lottery_target_weekdays=[str(item) for item in target_weekdays] or ["土"],
        lottery_default_entries=default_entries,
        lottery_account_overrides=account_overrides,
        lottery_max_entries_per_account=max_entries_per_account,
        lottery_search_weeks=search_weeks,
        lottery_dry_run=dry_run,
    )


def save_preferences_data(path: str | Path, data: dict):
    prefs_path = Path(path)
    prefs_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        import yaml  # type: ignore
    except Exception:
        yaml = None

    if yaml is not None:
        text = yaml.safe_dump(
            data,
            allow_unicode=True,
            sort_keys=False,
            default_flow_style=False,
        )
        prefs_path.write_text(text, encoding="utf-8")
        return prefs_path

    def dump_value(value, indent=0):
        prefix = " " * indent
        if isinstance(value, dict):
            lines = []
            for key, child in value.items():
                if isinstance(child, (dict, list)):
                    lines.append(f"{prefix}{key}:")
                    lines.extend(dump_value(child, indent + 2))
                else:
                    lines.append(f"{prefix}{key}: {format_scalar(child)}")
            return lines
        if isinstance(value, list):
            lines = []
            for item in value:
                if isinstance(item, dict):
                    first = True
                    for key, child in item.items():
                        if isinstance(child, (dict, list)):
                            if first:
                                lines.append(f"{prefix}- {key}:")
                                lines.extend(dump_value(child, indent + 4))
                            else:
                                lines.append(f"{prefix}  {key}:")
                                lines.extend(dump_value(child, indent + 4))
                        else:
                            if first:
                                lines.append(
                                    f"{prefix}- {key}: {format_scalar(child)}"
                                )
                            else:
                                lines.append(
                                    f"{prefix}  {key}: {format_scalar(child)}"
                                )
                        first = False
                else:
                    lines.append(f"{prefix}- {format_scalar(item)}")
            return lines
        return [f"{prefix}{format_scalar(value)}"]

    def format_scalar(value):
        if isinstance(value, bool):
            return "true" if value else "false"
        if isinstance(value, int):
            return str(value)
        text = str(value)
        if not text:
            return '""'
        if any(char in text for char in [":", "#", "-", '"']) or text.strip() != text:
            return json.dumps(text, ensure_ascii=False)
        return text

    prefs_path.write_text("\n".join(dump_value(data)) + "\n", encoding="utf-8")
    return prefs_path
