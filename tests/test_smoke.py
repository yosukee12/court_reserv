# -*- coding: utf-8 -*-

from pathlib import Path

import court_reserv


def test_package_importable():
    assert court_reserv is not None


def test_gui_entrypoint_exists():
    assert Path("court_reserv/court_reserv.py").exists()
