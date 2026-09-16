# -*- coding: utf-8 -*-

from court_reserv.config.preferences import load_reservation_preference


def test_load_preference_ignores_removed_lottery_debug_settings(tmp_path):
    preferences_path = tmp_path / "preferences.yaml"
    preferences_path.write_text(
        """lottery:
  manual_final_submit: true
  manual_preconfirm_submit: true
  reuse_browser_session: true
""",
        encoding="utf-8",
    )

    preference = load_reservation_preference(preferences_path)

    assert not hasattr(preference, "lottery_manual_final_submit")
    assert not hasattr(preference, "lottery_manual_preconfirm_submit")
    assert not hasattr(preference, "lottery_reuse_browser_session")
