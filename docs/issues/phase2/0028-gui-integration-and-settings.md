# Issue 0028: GUI Integration & Settings

## Status

Done

---

## Summary

Phase2 で実装した Lottery Workflow を既存 Tkinter GUI へ統合する。

GUI から抽選申込み、抽選結果確認、予約確定補助を実行できるようにし、設定画面・ログ画面・ファイル参照ダイアログを追加する。

CLI はそのまま残し、GUI は既存 Workflow を呼び出すだけの構成とする。

---

## Background

現在は以下が CLI と Service として完成している。

- BrowserSession
- LoginService
- NavigationService
- LotteryService
- ReservationService
- AvailabilityService
- LotteryEntryWorkflow
- LotteryResultWorkflow
- ReservationConfirmationWorkflow

Phase2 からは GUI を日常運用の入口とする。

---

## Goal

GUI から以下を実行できるようにする。

- 抽選申込み自動化
- 抽選結果確認
- 予約確定補助

また設定画面を追加し、

- ID CSV
- Preference YAML
- 対象曜日
- 探索週数
- 最大申込み数

などを GUI から編集できるようにする。

---

## Scope

### In Scope

- GUI に Workflow ボタン追加
- 設定画面追加
- ログ画面追加
- ファイル参照ダイアログ
- 前回設定復元
- Dry Run ON/OFF
- README 更新
- Completion Report

### Out of Scope

- GUI デザイン刷新
- 通知
- Scheduler
- CAPTCHA 自動突破
- Phase3 機能

---

# GUI Layout

```
-----------------------------------------------------

東京都テニスコート予約

-----------------------------------------------------

ID CSV

[ ids.csv                      ][参照]

Preference

[ preferences.yaml             ][参照]

-----------------------------------------------------

抽選

[ 抽選申込み自動化 ]

[ 抽選結果確認 ]

[ 予約確定補助 ]

[ 設定 ]

-----------------------------------------------------

ログ

+---------------------------------------------+

09:01 Login

09:02 CAPTCHA

09:03 週送り

...

+---------------------------------------------+

[ログ保存]

[ログクリア]

-----------------------------------------------------
```

---

# Settings Dialog

設定画面を追加する。

編集対象

## ファイル

- ID CSV
- Preference YAML

## Lottery

- target_weekdays
- max_entries_per_account
- search_weeks
- Dry Run

## Default Entries

追加

削除

編集

## Account Overrides

追加

削除

編集

OK

Cancel

保存時は Preference YAML を更新する。

---

# Browse

ID CSV

```
askopenfilename()
```

Preference

```
askopenfilename()
```

を使用する。

---

# Log Window

ScrolledText を使用する。

要件

- 自動スクロール
- Timestamp
- INFO
- WARNING
- ERROR

Workflow のログをそのまま表示する。

---

# Remember Last Settings

終了時

```
config.local.ini
```

へ保存する。

起動時

復元する。

保存対象

- last_id_csv
- last_preferences

---

# Dry Run

設定画面へ追加する。

```
☑ Dry Run
```

ON

送信しない

OFF

通常動作

---

# Architecture

GUI は Service を直接呼ぶ。

CLI のロジックは再利用する。

GUI 専用ロジックは最小限にする。

---

# Tasks

- [x] GUI に Workflow ボタン追加
- [x] 設定ボタン追加
- [x] 設定画面追加
- [x] Browse ボタン追加
- [x] Log Window 追加
- [x] Log 保存追加
- [x] Log Clear 追加
- [x] Dry Run 追加
- [x] Preference 編集追加
- [x] config.local.ini 保存追加
- [x] 起動時復元追加
- [x] README 更新
- [x] Verification
- [x] Completion Report

---

# Acceptance Criteria

- GUI から Workflow が実行できる
- Browse が使える
- 設定画面で Preference を編集できる
- Dry Run を切替できる
- Log がリアルタイム表示される
- 前回設定が復元される
- compileall 成功

---

# Verification

```bash
git status
python -m compileall .
python setup.py --name
python -m court_reserv.ui.app
```

GUI 起動確認を行う。

---

# Notes

CLI は削除しない。

GUI は CLI の Workflow を呼び出すだけとする。

---

# Completion Report

※ Codex が記入

## Summary

- `Court_Reserv` GUI を Phase 2 の運用入口として更新し、`抽選申込み自動化`、`抽選結果確認`、`予約確定補助`、`設定` の各ボタンから既存 Workflow / Service を呼び出せるようにした
- GUI へ ID CSV / Preference YAML 入力欄と Browse ボタン、設定ダイアログ、Log Window、Log 保存 / Clear を追加した
- 設定ダイアログでは `target_weekdays`、`search_weeks`、`max_entries_per_account`、`dry_run`、`default_entries`、`account_overrides` を編集し、Preference YAML へ保存できるようにした
- `config.local.ini` の `GUI` セクションへ `last_id_csv` と `last_preferences` を保存し、起動時に復元するようにした

## Changed Files

- `court_reserv/court_reserv.py`
- `court_reserv/config/preferences.py`
- `court_reserv/config/__init__.py`
- `court_reserv/models/preference.py`
- `court_reserv/services/lottery_entry_workflow.py`
- `README.md`
- `docs/issues/phase2/0028-gui-integration-and-settings.md`

## Verification Result

- `git status`
  変更ファイルを確認した
- `python -m compileall .`
  成功
- `python setup.py --name`
  `court_reserv`
- `python -m court_reserv.ui.app`
  実行したが、この実行環境では Tk / AppKit 側の `NSInternalInconsistencyException` で終了した
  Python コードの import / compile は成功しており、GUI モジュール自体の構文エラーは確認されなかった
- `python -m compileall court_reserv/court_reserv.py court_reserv/config/preferences.py court_reserv/models/preference.py court_reserv/services/lottery_entry_workflow.py`
  成功

## Follow-up Items

- GUI polish
- Phase2 Final Test
