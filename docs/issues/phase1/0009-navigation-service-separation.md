# Issue 0009: Navigation Service Separation

## Status

Done

---

## Summary

`Court_Reserv` クラスから画面遷移・JavaScript 呼び出し・共通ナビゲーション処理を `court_reserv/browser/navigation.py` に切り出す。

このIssueでは画面遷移の責務分離のみを行い、ログイン・抽選・予約処理のロジックは変更しない。

---

## Background

Issue 0007 で BrowserSession、Issue 0008 で LoginService を切り出した。

現在も `Court_Reserv` には以下が残っている。

- メニュー画面遷移
- JavaScript 呼び出し
- ボタン押下
- 共通画面遷移
- 抽選画面への遷移
- 予約画面への遷移

これらは業務ロジックではなく Browser 層の責務である。

---

## Goal

- Navigation 関連を Browser 層へ集約する
- JavaScript 呼び出しを一元管理する
- 画面遷移を共通化する
- 既存動作を維持する
- 次Issueで Lottery Service を切り出しやすくする

---

## Scope

### In Scope

- `court_reserv/browser/navigation.py` の追加
- `NavigationService` の追加
- JavaScript 実行処理の切り出し
- 共通画面遷移処理の切り出し
- `Court_Reserv` 側の最小限の変更
- docs 更新
- Completion Report 更新

### Out of Scope

- ログイン処理変更
- 抽選処理変更
- 予約処理変更
- Selenium 4 移行
- `find_element_by_*` の置換
- UI変更
- CAPTCHA / reCAPTCHA 処理変更
- 自動予約機能追加

---

## Proposed Design

```text
NavigationService

- open_menu()
- execute_script()
- go_to_lottery()
- go_to_reservation()
- go_to_result()
- click_button()
```

必要以上の抽象化は行わず、既存コードを安全に移動する。

---

## Automation Policy

- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない
- CAPTCHA / reCAPTCHA が表示された場合は手動認証を待つ
- 手動認証後は既存フローを継続する

---

## Implementation Plan

1. `browser/navigation.py` を追加
2. JavaScript 呼び出しを切り出す
3. 共通画面遷移を切り出す
4. Browser 層へ責務を集約する
5. `Court_Reserv` は NavigationService を利用する
6. docs 更新
7. Verification
8. Completion Report 更新

---

## Target Files

- court_reserv/browser/navigation.py
- court_reserv/browser/__init__.py
- court_reserv/court_reserv.py
- docs/ARCHITECTURE.md
- docs/DEVELOPMENT.md
- docs/issues/phase1/0009-navigation-service-separation.md

---

## Tasks

- [x] navigation.py 作成
- [x] NavigationService 作成
- [x] JavaScript 呼び出し切り出し
- [x] 共通画面遷移切り出し
- [x] Court_Reserv 修正
- [x] docs 更新
- [x] compileall 実施
- [x] Completion Report 更新

---

## Acceptance Criteria

- [x] browser/navigation.py が追加されている
- [x] JavaScript 呼び出しが NavigationService に集約されている
- [x] 共通画面遷移が NavigationService に集約されている
- [x] ログイン処理を変更していない
- [x] 抽選処理を変更していない
- [x] 予約処理を変更していない
- [x] UI変更を行っていない
- [x] Selenium 4 移行を行っていない
- [x] CAPTCHA / reCAPTCHA 方針を変更していない
- [x] python -m compileall . が成功する

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
```

可能なら

```bash
python court_reserv/court_reserv.py
python -m court_reserv.ui.app
```

GUI起動確認が難しい場合は Completion Report に理由を記載する。

---

## Notes

- このIssueは Navigation の責務分離のみを対象とする。
- Lottery Service の切り出しは次Issueで行う。
- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない。

---

# Completion Report

※ Codex が記入

## Summary

- `court_reserv/browser/navigation.py` を追加し、JavaScript 実行、共通画面遷移、抽選画面・確認画面・空き確認画面への共通ナビゲーションを `NavigationService` に集約した。
- `court_reserv/court_reserv.py` では `NavigationService` の呼び出しに置き換える最小変更のみを行い、ログイン処理、抽選処理、予約処理、空き確認処理、Tkinter UI は変更していない。
- `docs/ARCHITECTURE.md` と `docs/DEVELOPMENT.md` に Navigation Service 分離方針を追記した。

---

## Changed Files

- `court_reserv/browser/__init__.py`
- `court_reserv/browser/navigation.py`
- `court_reserv/court_reserv.py`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase1/0009-navigation-service-separation.md`

---

## Verification Result

- `git status` 実行
- `python -m compileall .` 成功
- `python setup.py --name` で `court_reserv` を確認
- GUI起動確認 (`python court_reserv/court_reserv.py`, `python -m court_reserv.ui.app`) は実行環境依存のため未実施

---

## Follow-up Items

- Lottery Service の切り出し
- Reservation Service の切り出し
- 空き確認処理のサービス分離
