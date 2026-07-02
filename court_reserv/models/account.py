# -*- coding: utf-8 -*-
"""Account model foundation."""

from dataclasses import dataclass


@dataclass(frozen=True)
class Account:
    """Represents a reserving user account."""

    user_id: str
    password: str | None = None
    name: str | None = None
    name_kana: str | None = None
    is_active: bool = True
