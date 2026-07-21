# -*- coding: utf-8 -*-

from configparser import ConfigParser
from pathlib import Path

from court_reserv.services.id_manager import IdManagerService
from court_reserv.services import LotteryApplicationCheckWorkflowService


_TENNIS_HTML = """
<html><body>
  <table>
    <thead>
      <tr>
        <th>申込み</th><th>状態</th><th>分類</th><th>公園・施設</th><th>利用日</th><th>時刻</th><th>取消</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td>申込み番号：1</td>
        <td>状態：受付中</td>
        <td>分類：テニス（人工芝）</td>
        <td>公園・施設：府中の森公園 テニス（人工芝）</td>
        <td>利用日：5月18日(土曜) 2024年</td>
        <td>時刻：09時00分～11時00分</td>
        <td><button>取消</button></td>
      </tr>
    </tbody>
  </table>
</body></html>
"""


class _DummyBrowserSession:
    def __init__(self, wait_results=None):
        self._driver = _DummyDriver()
        self._wait_results = list(wait_results or [])

    def create_driver(self):
        return self._driver

    def safe_close(self, driver):
        return None

    def get_wait(self, driver, timeout=10):
        return _DummyWait(self._wait_results.pop(0) if self._wait_results else True)


class _DummyLoginService:
    def __init__(self, login_results=None):
        self._login_results = dict(login_results or {})

    def login(self, driver, user_id, password):
        driver.account_id = user_id
        driver.title = "抽選申込みの確認"
        driver.page_source = _TENNIS_HTML
        return self._login_results.get(user_id, True)


class _DummyNavigationService:
    def __init__(self):
        self.logout_calls = 0

    def go_to_lottery_cancel_list(self, driver):
        return None

    def logout(self, driver):
        self.logout_calls += 1
        return None


class _DummyIdManagerService:
    def __init__(self, accounts=None):
        self._accounts = accounts or {}

    def load_accounts(self, csv_file_path):
        return self._accounts

    def save_accounts(self, id_dict, output_file_path):
        Path(output_file_path).write_text("saved", encoding="utf-8")


class _DummyDriver:
    def __init__(self):
        self.account_id = ""
        self.title = "抽選申込みの確認"
        self.page_source = _TENNIS_HTML


class _DummyWait:
    def __init__(self, result):
        self.result = result

    def until(self, predicate):
        return self.result


def test_parse_application_rows_filters_baseball():
    html = """
    <html><body>
      <table>
        <thead>
          <tr>
            <th>申込み</th><th>状態</th><th>分類</th><th>公園・施設</th><th>利用日</th><th>時刻</th><th>取消</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>申込み番号：1</td>
            <td>状態：受付中</td>
            <td>分類：テニス（人工芝）</td>
            <td>公園・施設：府中の森公園 テニス（人工芝）</td>
            <td>利用日：5月18日(土曜) 2024年</td>
            <td>時刻：09時00分～11時00分</td>
            <td><button>取消</button></td>
          </tr>
          <tr>
            <td>申込み番号：2</td>
            <td>状態：受付中</td>
            <td>分類：野球（軟式）</td>
            <td>公園・施設：府中の森公園 野球場</td>
            <td>利用日：5月19日(日曜) 2024年</td>
            <td>時刻：11時00分～13時00分</td>
            <td><button>取消</button></td>
          </tr>
        </tbody>
      </table>
    </body></html>
    """

    service = LotteryApplicationCheckWorkflowService(
        config=ConfigParser(),
        browser_session=_DummyBrowserSession(),
        login_service=_DummyLoginService(),
        navigation_service=_DummyNavigationService(),
        id_manager_service=_DummyIdManagerService(),
    )

    rows, excluded, state = service.parse_application_rows(
        html,
        masked_user_id="12***78",
        account_label="sample",
    )

    assert len(rows) == 1
    assert rows[0]["sport"] == "テニス（人工芝）"
    assert rows[0]["application_no"] == "1"
    assert rows[0]["account"] == "12***78"
    assert excluded["baseball"] == 1
    assert state["table_exists"] is True
    assert state["row_count"] == 2
    assert state["bs_day_count"] >= 1


def test_save_result_uses_timestamped_csv_name(tmp_path):
    service = LotteryApplicationCheckWorkflowService(
        config=ConfigParser(),
        browser_session=_DummyBrowserSession(),
        login_service=_DummyLoginService(),
        navigation_service=_DummyNavigationService(),
        id_manager_service=_DummyIdManagerService(),
    )

    csv_path = service.save_result(
        {
            "output_id_dict": {
                "10000001": ["山田太郎", "ヤマダタロウ", "pass", "5月18日(土曜) 09時00分～11時00分"]
            }
        },
        tmp_path,
    )

    assert csv_path.parent == tmp_path
    assert csv_path.name.startswith("check_lottery_status_")
    assert csv_path.suffix == ".csv"
    assert csv_path.exists()


def test_resolve_output_dir_defaults_to_id_csv_parent(tmp_path):
    service = LotteryApplicationCheckWorkflowService(
        config=ConfigParser(),
        browser_session=_DummyBrowserSession(),
        login_service=_DummyLoginService(),
        navigation_service=_DummyNavigationService(),
        id_manager_service=_DummyIdManagerService(),
    )

    id_csv = tmp_path / "ids.csv"
    assert service.resolve_output_dir(id_csv=str(id_csv)) == tmp_path.resolve()


def test_run_keeps_failed_login_accounts_in_output_dict():
    service = LotteryApplicationCheckWorkflowService(
        config=ConfigParser(),
        browser_session=_DummyBrowserSession(),
        login_service=_DummyLoginService(login_results={"10000001": False}),
        navigation_service=_DummyNavigationService(),
        id_manager_service=_DummyIdManagerService(
            {"10000001": ["山田太郎", "ヤマダタロウ", "pass"]}
        ),
    )

    result = service.run(id_csv="ignored.csv")

    assert result["account_summaries"][0]["status"] == "login_failed"
    assert result["output_id_dict"]["10000001"] == [
        "山田太郎",
        "ヤマダタロウ",
        "pass",
        "",
        "",
    ]


def test_run_wait_failure_skips_account_and_continues():
    id_manager = _DummyIdManagerService(
        {
            "10000001": ["山田太郎", "ヤマダタロウ", "pass"],
            "10000002": ["佐藤花子", "サトウハナコ", "pass2"],
        }
    )
    browser_session = _DummyBrowserSession(wait_results=[False, True])
    login_service = _DummyLoginService()
    navigation_service = _DummyNavigationService()
    service = LotteryApplicationCheckWorkflowService(
        config=ConfigParser(),
        browser_session=browser_session,
        login_service=login_service,
        navigation_service=navigation_service,
        id_manager_service=id_manager,
    )

    result = service.run(id_csv="ignored.csv")

    assert result["account_summaries"][0]["status"] == "page_wait_failed"
    assert result["account_summaries"][1]["status"] == "completed"
    assert len(result["rows"]) == 1
    assert result["rows"][0]["account"] == "10***02"
    assert result["output_id_dict"]["10000001"] == [
        "山田太郎",
        "ヤマダタロウ",
        "pass",
        "",
        "",
    ]
    assert result["output_id_dict"]["10000002"][3] != ""


def test_save_accounts_skips_short_values(tmp_path):
    service = IdManagerService(config=ConfigParser())
    output_path = tmp_path / "accounts.csv"

    service.save_accounts({"10000001": ["山田太郎", "ヤマダタロウ"]}, output_path)

    assert output_path.read_text(encoding="utf-8") == ""
