# -*- coding: utf-8 -*-

import csv

from court_reserv.services.fftc_wiki_text import FftcWikiTextService


def test_build_from_confirmed_csv_groups_winners_by_date_and_time(tmp_path):
    csv_path = tmp_path / "reservation_confirmed.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        csv.writer(handle).writerows(
            [
                [
                    "10062787",
                    "川上大輔",
                    "カワカミダイスケ",
                    "password",
                    "7月18日(土曜) 2026年 9時00分～11時00分",
                ],
                [
                    "10081183",
                    "鈴木敦士",
                    "スズキアツシ",
                    "password",
                    "7月18日(土曜) 2026年 9時00分～11時00分",
                ],
                [
                    "10048008",
                    "小林秀太郎",
                    "コバヤシヒデタロウ",
                    "password",
                    "2026年7月18日 11時00分～13時00分",
                ],
                [
                    "10072185",
                    "馬場正裕",
                    "ババマサヒロ",
                    "password",
                    "2026年7月18日 11時00分～13時00分",
                ],
            ]
        )

    result = FftcWikiTextService().build_from_csv(csv_path)

    assert result == (
        "**7/18(土) 府中の森公園　オムニコート 9:00～11:00 2面 "
        "11:00～13:00 2面　当選者名：川上大輔(10062787)、鈴木敦士(10081183)//"
        "小林秀太郎(10048008)、馬場正裕(10072185)\n"
        "※受付時に当選者と受付者(当選者が本人でない場合)の利用者番号が必要になりました。\n"
        "\n"
        "-参加予定："
    )


def test_build_from_confirmed_csv_repeats_note_for_each_date(tmp_path):
    csv_path = tmp_path / "reservation_confirmed.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        csv.writer(handle).writerows(
            [
                ["10062787", "川上大輔", "", "password", "2026年7月18日 9時00分～11時00分"],
                ["10072185", "馬場正裕", "", "password", "2026年7月18日 11時00分～13時00分"],
                ["10061748", "中川洋介", "", "password", "2026年8月8日 11時00分～13時00分"],
                ["10062051", "度会公司", "", "password", "2026年8月8日 11時00分～13時00分"],
            ]
        )

    result = FftcWikiTextService().build_from_csv(csv_path)

    assert result == (
        "**7/18(土) 府中の森公園　オムニコート 9:00～11:00 1面 "
        "11:00～13:00 1面　当選者名：川上大輔(10062787)//馬場正裕(10072185)\n"
        "※受付時に当選者と受付者(当選者が本人でない場合)の利用者番号が必要になりました。\n"
        "\n"
        "-参加予定：\n"
        "\n"
        "**8/8(土) 府中の森公園　オムニコート 11:00～13:00 2面　当選者名："
        "中川洋介(10061748)、度会公司(10062051)\n"
        "※受付時に当選者と受付者(当選者が本人でない場合)の利用者番号が必要になりました。\n"
        "\n"
        "-参加予定："
    )
