# Issue 0013: ID Management Separation

## Status

Done

---

## Summary

ID管理・CSV入出力・ID有効確認関連処理を `court_reserv/services/id_manager.py` に整理する。

このIssueでは、既存のID管理処理の責務分離のみを行い、抽選・予約・空き確認・GUI挙動は変更しない。

---

## Background

Issue 0007〜0012 により、BrowserSession、LoginService、NavigationService、LotteryService、ReservationService、AvailabilityService を切り出した。

現在、ID管理は主に `court_reserv/manage_id.py` に残っており、CSV入出力、ID一覧管理、有効期限確認、設定参照が混在している。

今後の自動予約・複数ID運用・予約戦略エンジンに向けて、ID管理の責務を service 層へ整理する。

---

## Goal

- ID管理処理を `services/id_manager.py` に整理する
- 既存の `Manage_Id` 互換性を維持する
- CSV読み込み・書き込み・有効確認処理の責務を明確化する
- `Court_Reserv` からのID管理利用を整理する
- 既存挙動は変更しない

---

## Scope

### In Scope

- `court_reserv/services/id_manager.py` の追加
- `IdManagerService` または同等クラスの追加
- `manage_id.py` との役割整理
- 既存 `Manage_Id` の互換ラッパー化、または service への委譲
- ID CSV 読み込み処理の整理
- ID CSV 書き込み処理の整理
- ID有効確認処理の整理
- `Court_Reserv` 側の最小限の呼び出し変更
- docs/ARCHITECTURE.md 更新
- docs/DEVELOPMENT.md 更新
- Completion Report 更新

### Out of Scope

- 抽選処理の変更
- 予約確定処理の変更
- 空き確認処理の変更
- Selenium 4 移行
- UI変更
- CSVフォーマット変更
- ID管理画面の新規追加
- 自動予約機能追加
- CAPTCHA / reCAPTCHA 処理変更

---

## Proposed Design

`court_reserv/services/id_manager.py` に以下のような責務を持たせる。

```text
IdManagerService

- load_accounts()
- save_accounts()
- check_account_validity()
- get_active_accounts()
```

既存互換のため、`manage_id.py` の `Manage_Id` は当面残してよい。

推奨方針：

```text
court_reserv/manage_id.py
    既存互換ラッパー

court_reserv/services/id_manager.py
    新しい実処理
```

このIssueでは既存挙動維持を優先する。

---

## Automation Policy

- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない
- CAPTCHA / reCAPTCHA が表示された場合は手動認証を待つ
- 手動認証後は既存フローを継続する
- 外部CAPTCHAサービスは利用しない

---

## Implementation Plan

1. `court_reserv/services/id_manager.py` を追加する
2. `IdManagerService` を追加する
3. `manage_id.py` の既存処理を確認する
4. 可能な範囲で処理を `IdManagerService` に移す
5. `Manage_Id` は互換性維持のため残し、service へ委譲する
6. `Court_Reserv` 側は最小限の変更に留める
7. CSVフォーマットは変更しない
8. docs/ARCHITECTURE.md を更新する
9. docs/DEVELOPMENT.md を更新する
10. Verification を実行する
11. Completion Report を記入する

---

## Target Files

- court_reserv/services/id_manager.py
- court_reserv/services/__init__.py
- court_reserv/manage_id.py
- court_reserv/court_reserv.py
- docs/ARCHITECTURE.md
- docs/DEVELOPMENT.md
- docs/issues/phase1/0013-id-management-separation.md

---

## Tasks

- [x] `services/id_manager.py` を追加する
- [x] `IdManagerService` を追加する
- [x] `manage_id.py` を互換ラッパー化または service 委譲にする
- [x] ID CSV 読み込み処理を整理する
- [x] ID CSV 書き込み処理を整理する
- [x] ID有効確認処理を整理する
- [x] `Court_Reserv` 側を最小限修正する
- [x] docs を更新する
- [x] compileall を実行する
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] `court_reserv/services/id_manager.py` が追加されている
- [x] ID管理処理が `IdManagerService` に整理されている
- [x] `Manage_Id` の既存互換性が維持されている
- [x] CSVフォーマットを変更していない
- [x] 抽選処理を変更していない
- [x] 予約確定処理を変更していない
- [x] 空き確認処理を変更していない
- [x] UI挙動を変更していない
- [x] Selenium 4 移行を行っていない
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

ID CSV を使った動作確認ができる場合は、以下も確認する。

```text
ID一覧CSVを読み込めること
ID有効確認の入口が動くこと
```

実データ確認が難しい場合は、理由を Completion Report に記載する。

---

## Notes

- このIssueはID管理の責務整理のみを対象とする。
- CSVフォーマット変更は行わない。
- `Manage_Id` は既存互換のため当面残してよい。
- 自動予約機能は別Issueで扱う。
- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない。

---

# Completion Report

※ Codex が記入

## Summary

- `court_reserv/services/id_manager.py` を追加し、ID CSV 読み込み、CSV 書き出し、ID 有効確認を `IdManagerService` に整理した。
- `court_reserv/manage_id.py` は既存互換のため残し、`IdManagerService` への委譲ラッパーに変更した。
- `court_reserv/court_reserv.py` では ID 読み込み・書き出し・有効確認の入口だけを `IdManagerService` 経由に変更し、抽選処理、予約処理、空き確認処理、Tkinter UI は変更していない。
- `docs/ARCHITECTURE.md` と `docs/DEVELOPMENT.md` に IdManager Service 整理方針を追記した。

---

## Changed Files

- `court_reserv/services/__init__.py`
- `court_reserv/services/id_manager.py`
- `court_reserv/manage_id.py`
- `court_reserv/court_reserv.py`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase1/0013-id-management-separation.md`

---

## Verification Result

- `git status` 実行
- `python -m compileall .` 成功
- `python setup.py --name` で `court_reserv` を確認
- GUI起動確認 (`python court_reserv/court_reserv.py`, `python -m court_reserv.ui.app`) は実行環境依存のため未実施
- ID CSV を使った実動作確認は手元データ依存のため未実施

---

## Follow-up Items

- モデル整理と型の明確化
- 旧 `Manage_Id` API の将来的な縮退方針整理
- 自動予約機能や予約戦略エンジンは別 Issue で扱う
