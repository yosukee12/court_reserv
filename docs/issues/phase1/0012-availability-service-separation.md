# Issue 0012: Availability Service Separation

## Status

Done

---

## Summary

`Court_Reserv` クラスから空き確認・空き枠収集関連処理を `court_reserv/services/availability.py` に切り出す。

このIssueでは空き確認系処理の責務分離のみを行い、抽選処理・予約確定処理・ID管理処理・自動予約機能は変更しない。

---

## Background

Issue 0007 で BrowserSession、Issue 0008 で LoginService、Issue 0009 で NavigationService、Issue 0010 で LotteryService、Issue 0011 で ReservationService を切り出した。

現在も `Court_Reserv` には空き確認・空き枠収集関連の業務処理が残っている。

これらは UI ではなく service 層の責務であるため、今後の予約戦略エンジンや自動予約機能に向けて分離する。

---

## Goal

- 空き確認・空き枠収集関連処理を `services/availability.py` に分離する
- `Court_Reserv` からは `AvailabilityService` を呼び出す形にする
- 既存の空き確認フローを維持する
- 抽選処理・予約確定処理・UI挙動は変更しない
- 次Issue以降で ID 管理やモデル整理を行いやすくする

---

## Scope

### In Scope

- `court_reserv/services/availability.py` の追加
- `AvailabilityService` または同等クラスの追加
- 空き確認処理の最小限の切り出し
- 空き枠収集処理の最小限の切り出し
- `Court_Reserv` 側の最小限の呼び出し変更
- docs/ARCHITECTURE.md 更新
- docs/DEVELOPMENT.md 更新
- Completion Report 更新

### Out of Scope

- 抽選処理の変更
- 予約確定処理の変更
- ID管理処理の切り出し
- UI変更
- Selenium 4 移行
- `find_element_by_*` 置換
- 自動予約機能追加
- 予約戦略エンジン追加
- CAPTCHA / reCAPTCHA 処理変更

---

## Proposed Design

`court_reserv/services/availability.py` に以下のような責務を持たせる。

```text
AvailabilityService

- check_availability(...)
- collect_available_slots(...)
- save_available_slots(...)
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

1. `court_reserv/services/availability.py` を追加する
2. `AvailabilityService` を追加する
3. `Court_Reserv` 内の空き確認・空き枠収集関連処理を特定する
4. 空き確認関連処理を `AvailabilityService` に最小限切り出す
5. `Court_Reserv` は UI イベントから `AvailabilityService` を呼び出す形にする
6. 抽選処理・予約確定処理・ID管理処理は変更しない
7. docs/ARCHITECTURE.md を更新する
8. docs/DEVELOPMENT.md を更新する
9. Verification を実行する
10. Completion Report を記入する

---

## Target Files

- court_reserv/services/availability.py
- court_reserv/services/__init__.py
- court_reserv/court_reserv.py
- docs/ARCHITECTURE.md
- docs/DEVELOPMENT.md
- docs/issues/phase1/0012-availability-service-separation.md

---

## Tasks

- [x] `services/availability.py` を追加する
- [x] `AvailabilityService` を追加する
- [x] 空き確認処理を最小限切り出す
- [x] 空き枠収集処理を最小限切り出す
- [x] `Court_Reserv` 側を最小限修正する
- [x] docs を更新する
- [x] compileall を実行する
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] `court_reserv/services/availability.py` が追加されている
- [x] 空き確認・空き枠収集関連処理が `AvailabilityService` から利用されている
- [x] 既存の空き確認フローを変更していない
- [x] 抽選処理を変更していない
- [x] 予約確定処理を変更していない
- [x] ID管理処理を変更していない
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
空き確認画面へ遷移できること
空き確認処理が既存どおり実行できること
```

GUI起動確認や実ブラウザ確認が難しい場合は、理由を Completion Report に記載する。

---

## Notes

- このIssueは空き確認・空き枠収集関連処理の分離のみを対象とする。
- 抽選処理は変更しない。
- 予約確定処理は変更しない。
- ID管理処理は次Issue以降で扱う。
- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない。

---

# Completion Report

※ Codex が記入

## Summary

- `court_reserv/services/availability.py` を追加し、空き確認と空き枠収集の業務フローを `AvailabilityService` に集約した。
- `court_reserv/court_reserv.py` では `AvailabilityService` への委譲に留め、抽選処理、予約確定処理、ID 管理処理、Tkinter UI は変更していない。
- `docs/ARCHITECTURE.md` と `docs/DEVELOPMENT.md` に Availability Service 分離方針を追記した。

---

## Changed Files

- `court_reserv/services/__init__.py`
- `court_reserv/services/availability.py`
- `court_reserv/court_reserv.py`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase1/0012-availability-service-separation.md`

---

## Verification Result

- `git status` 実行
- `python -m compileall .` 成功
- `python setup.py --name` で `court_reserv` を確認
- GUI起動確認 (`python court_reserv/court_reserv.py`, `python -m court_reserv.ui.app`) は実行環境依存のため未実施
- 実ブラウザ確認は手元環境依存のため未実施

---

## Follow-up Items

- ID 管理処理の分離
- モデル整理と責務境界の明確化
- 自動予約機能や予約戦略エンジンは別 Issue で扱う
