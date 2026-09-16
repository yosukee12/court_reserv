# -*- coding: utf-8 -*-

from types import SimpleNamespace

from court_reserv.services.lottery_entry_workflow import LotteryEntryWorkflowService


class _DummyBrowserSession:
    def __init__(self):
        self.create_calls = 0
        self.quit_calls = []

    def create_driver(self):
        self.create_calls += 1
        return object()

    def safe_quit(self, driver):
        self.quit_calls.append(driver)


def test_run_reuses_one_browser_for_multiple_accounts_by_default():
    browser_session = _DummyBrowserSession()
    workflow = LotteryEntryWorkflowService(
        config=None,
        browser_session=browser_session,
        login_service=SimpleNamespace(),
        navigation_service=SimpleNamespace(),
        lottery_service=SimpleNamespace(),
        id_manager_service=None,
        slot_collector=None,
    )
    workflow._apply_sleep_policy = lambda preference: None
    workflow.resolve_accounts = lambda **kwargs: [
        {"user_id": "10000001", "password": "", "source": "id_csv"},
        {"user_id": "10000002", "password": "", "source": "id_csv"},
    ]
    workflow._logout_and_verify_login_page = lambda driver: {"success": True}
    workflow._prepare_next_account = lambda: None

    calls = []

    def run_account(**kwargs):
        calls.append(kwargs)
        return {"status": "completed"}

    workflow._run_account_with_retry = run_account

    result = workflow.run(
        preference=SimpleNamespace(lottery_target_weekdays=["土"]),
    )

    assert browser_session.create_calls == 1
    assert calls[0]["driver"] is calls[1]["driver"]
    assert calls[0]["session_reused"] is False
    assert calls[1]["session_reused"] is True
    assert browser_session.quit_calls == [calls[0]["driver"]]
    assert result["session_strategy"] == "reused_browser_session"
