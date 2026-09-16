# -*- coding: utf-8 -*-

from types import SimpleNamespace

from selenium.common.exceptions import TimeoutException

from court_reserv.services.lottery import LotteryService


class _Alert:
    text = "抽選申込処理を行います。よろしいですか？"

    def __init__(self, events):
        self.events = events

    def accept(self):
        self.events.append("accept")


class _Wait:
    def __init__(self, calls):
        self.calls = calls

    def until(self, predicate):
        del predicate
        self.calls += 1
        if self.calls > 1:
            raise TimeoutException()
        return True


def test_submission_alert_waits_before_accepting():
    events = []
    wait = _Wait(0)
    alert = _Alert(events)
    driver = SimpleNamespace(switch_to=SimpleNamespace(alert=alert))
    login_service = SimpleNamespace(detect_captcha=lambda driver: False)
    logger = SimpleNamespace(info=lambda *args, **kwargs: None)
    service = LotteryService(
        config=None,
        browser_session=None,
        login_service=login_service,
        navigation_service=None,
        logger=logger,
        show_info=lambda title, message: None,
        output_id_dict=None,
        sleep_func=lambda seconds: events.append(("sleep", seconds)),
    )
    service._get_wait = lambda driver, timeout: wait
    result = {"last_alert_text": "", "states": []}

    accepted = service._accept_submission_alerts(driver, result, timeout=1)

    assert accepted is True
    assert events == [("sleep", 0.3), "accept", ("sleep", 0.3)]
