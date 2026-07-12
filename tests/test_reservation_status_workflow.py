from court_reserv.services.reservation_status_workflow import (
    ReservationStatusWorkflowService,
)
import csv


class FakeDriver:
    title = "予約確認"
    current_url = "https://example.test/reservations"
    page_source = """
    <html><body>
      <h2>予約の確認</h2>
      <table id="rsvacceptlist"><tbody>
        <tr><td>予約番号</td><td>2026年7月18日</td><td>9時00分～11時00分</td><td>府中の森公園 テニス（人工芝）</td></tr>
        <tr><td>予約番号</td><td>2026年7月19日</td><td>9時00分～11時00分</td><td>府中の森公園 野球場</td></tr>
      </tbody></table>
      <button id="cancel" onclick="cancelReservation()">キャンセル</button>
    </body></html>
    """


def test_inspect_page_reports_cancel_control_without_clicking():
    service = ReservationStatusWorkflowService(
        reservation_service=None,
        login_service=None,
        browser_session=None,
        navigation_service=None,
        sleep_func=lambda _: None,
    )

    summary = service.inspect_page(FakeDriver())

    assert summary["title"] == "予約確認"
    assert "2026年7月18日" in summary["body_text"]
    assert summary["reservation_count"] == 1
    assert {
        "tag": "button",
        "label": "キャンセル",
        "id": "cancel",
        "name": "",
        "type": "",
    } in summary["controls"]


def test_save_remaining_accounts_csv_removes_only_cancelled_ids(tmp_path):
    input_path = tmp_path / "ID_list.csv"
    output_path = tmp_path / "ID_list_after.csv"
    with input_path.open("w", newline="", encoding="utf-8") as handle:
        csv.writer(handle).writerows(
            [
                ["10048008", "対象者", "カナ", "password"],
                ["10048009", "残す人", "カナ", "password"],
            ]
        )

    ReservationStatusWorkflowService.save_remaining_accounts_csv(
        input_path, ["10048008"], output_path
    )

    with output_path.open(newline="", encoding="utf-8-sig") as handle:
        assert list(csv.reader(handle)) == [
            ["10048009", "残す人", "カナ", "password"]
        ]
