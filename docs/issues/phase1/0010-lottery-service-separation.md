# Issue 0010: Lottery Service Separation

## Status

Done

---

## Summary

`Court_Reserv` クラスから抽選関連処理を `court_reserv/services/lottery.py` に切り出す。

このIssueでは抽選申込み・抽選申込み状況確認・抽選当選結果確認に関する責務分離のみを行い、予約確定・空き確認・自動予約機能は変更しない。

---

## Background

Issue 0007 で BrowserSession、Issue 0008 で LoginService、Issue 0009 で NavigationService を切り出した。

現在も `Court_Reserv` には抽選関連の業務処理が残っている。

代表的な対象は以下。

- 抽選申込み
- 半自動抽選申込み
- 抽選申込み状況確認
- 抽選当選結果確認
- 抽選画面への遷移後の操作

これらは UI ではなく service 層の責務であるため、今後の自動予約・戦略エンジン実装に向けて分離する。

---

## Goal

- 抽選関連処理を `services/lottery.py` に分離する
- `Court_Reserv` からは `LotteryService` を呼び出す形にする
- 既存の抽選申込み・抽選確認フローを維持する
- 予約確定・空き確認・UI挙動は変更しない
- 次Issueで Reservation Service を切り出しやすくする

---

## Scope

### In Scope

- `court_reserv/services/lottery.py` の追加
- `LotteryService` または同等クラスの追加
- 抽選申込み関連処理の最小限の切り出し
- 抽選申込み状況確認処理の最小限の切り出し
- 抽選当選結果確認処理の最小限の切り出し
- `Court_Reserv` 側の最小限の呼び出し変更
- docs/ARCHITECTURE.md 更新
- docs/DEVELOPMENT.md 更新
- Completion Report 更新

### Out of Scope

- 予約確定処理の切り出し
- 空き確認処理の切り出し
- ID管理処理の切り出し
- UI変更
- Selenium 4 移行
- `find_element_by_*` 置換
- 自動予約機能追加
- CAPTCHA / reCAPTCHA 処理変更

---

## Proposed Design

`court_reserv/services/lottery.py` に以下のような責務を持たせる。

```text
LotteryService

- start_lottery_entry(...)
- run_semi_auto_lottery(...)
- check_lottery_entries(...)
- check_lottery_results(...)
```

ただし、既存コードとの整合を優先する。  
このIssueではインターフェースの美しさより、挙動を変えずに責務分離することを優先する。

---

## Automation Policy

- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない
- CAPTCHA / reCAPTCHA が表示された場合は手動認証を待つ
- 手動認証後は既存フローを継続する
- 外部CAPTCHAサービスは利用しない

---

## Implementation Plan

1. `court_reserv/services/lottery.py` を追加する
2. `LotteryService` を追加する
3. `Court_Reserv` 内の抽選関連処理を特定する
4. 抽選関連処理を `LotteryService` に最小限切り出す
5. `Court_Reserv` は UI イベントから `LotteryService` を呼び出す形にする
6. 予約確定・空き確認・ID管理処理は変更しない
7. docs/ARCHITECTURE.md を更新する
8. docs/DEVELOPMENT.md を更新する
9. Verification を実行する
10. Completion Report を記入する

---

## Target Files

- court_reserv/services/lottery.py
- court_reserv/services/__init__.py
- court_reserv/court_reserv.py
- docs/ARCHITECTURE.md
- docs/DEVELOPMENT.md
- docs/issues/phase1/0010-lottery-service-separation.md

---

## Tasks

- [x] `services/lottery.py` を追加する
- [x] `LotteryService` を追加する
- [x] 抽選申込み処理を最小限切り出す
- [x] 抽選申込み状況確認処理を最小限切り出す
- [x] 抽選当選結果確認処理を最小限切り出す
- [x] `Court_Reserv` 側を最小限修正する
- [x] docs を更新する
- [x] compileall を実行する
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] `court_reserv/services/lottery.py` が追加されている
- [x] 抽選関連処理が `LotteryService` から利用されている
- [x] 既存の抽選申込みフローを変更していない
- [x] 既存の抽選申込み状況確認フローを変更していない
- [x] 既存の抽選当選結果確認フローを変更していない
- [x] 予約確定処理を変更していない
- [x] 空き確認処理を変更していない
- [x] UI挙動を変更していない
- [x] Selenium 4 移行を行っていない
- [x] `find_element_by_*` 置換を行っていない
- [x] CAPTCHA / reCAPTCHA 方針を変更していない
- [x] `python -m compileall .` が成功する

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
```

可能なら以下も実行する。

```bash
python court_reserv/court_reserv.py
python -m court_reserv.ui.app
```

実ブラウザ確認ができる場合は、以下も確認する。

```text
抽選申込み画面へ遷移できること
抽選申込み状況確認画面へ遷移できること
抽選当選結果確認画面へ遷移できること
```

GUI起動確認や実ブラウザ確認が難しい場合は、理由を Completion Report に記載する。

---

## Notes

- このIssueは抽選関連処理の分離のみを対象とする。
- 予約確定処理は次Issue以降で扱う。
- 空き確認処理は次Issue以降で扱う。
- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない。

---

# Completion Report

※ Codex が記入

## Summary

- `court_reserv/services/lottery.py` を追加し、抽選申込み、抽選申込み状況確認、抽選当選結果確認、および関連する自動選択補助を `LotteryService` に集約した。
- `court_reserv/court_reserv.py` では `LotteryService` への委譲に留め、予約確定処理、空き確認処理、ID管理処理、Tkinter UI は変更していない。
- `docs/ARCHITECTURE.md` と `docs/DEVELOPMENT.md` に Lottery Service 分離方針を追記した。

---

## Changed Files

- `court_reserv/services/__init__.py`
- `court_reserv/services/lottery.py`
- `court_reserv/court_reserv.py`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase1/0010-lottery-service-separation.md`

---

## Verification Result

- `git status` 実行
- `python -m compileall .` 成功
- `python setup.py --name` で `court_reserv` を確認
- GUI起動確認 (`python court_reserv/court_reserv.py`, `python -m court_reserv.ui.app`) は実行環境依存のため未実施
- 実ブラウザ確認は手元環境依存のため未実施

---

## Follow-up Items

- Reservation Service の切り出し
- 空き確認処理のサービス分離
- 予約確定処理のサービス分離
