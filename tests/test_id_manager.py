# -*- coding: utf-8 -*-

from court_reserv.services.id_manager import IdManagerService


def test_load_accounts_accepts_bom_and_reservation_datetime_column(tmp_path):
    csv_path = tmp_path / "reservation_check.csv"
    csv_path.write_text(
        "10061748,中川洋介,ナカガワヨウスケ,password,8月8日(土曜) 11時00分～ 13時00分\n",
        encoding="utf-8-sig",
    )

    accounts = IdManagerService({}).load_accounts(csv_path)

    assert accounts == {
        "10061748": [
            "中川洋介",
            "ナカガワヨウスケ",
            "password",
            "8月8日(土曜) 11時00分～ 13時00分",
        ]
    }
