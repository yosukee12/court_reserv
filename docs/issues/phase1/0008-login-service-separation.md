# Issue 0008: Login Service Separation

## Status

Done

---

## Summary

`Court_Reserv` クラスからログイン関連処理を `court_reserv/browser/login.py` に切り出す。

このIssueでは、ログイン処理の責務分離のみを行い、予約処理・抽選処理・画面遷移・Selenium 4移行は行わない。

---

## Background

Issue 0007 で WebDriver 生成・ChromeOptions・WebDriverWait・終了処理を `BrowserSession` に切り出した。

次に、複数フローで再利用されるログイン処理を `LoginService` として分離し、今後の抽選・予約・空き確認サービス分割に備える。

現在、ログイン処理は `Court_Reserv` 内に残っており、UI・業務ロジック・Selenium操作と密結合している。

---

## Goal

- ログイン処理を `court_reserv/browser/login.py` に分離する
- `Court_Reserv` からは `LoginService` を呼び出す形にする
- reCAPTCHA / CAPTCHA 手動対応方針を維持する
- 既存のログイン挙動を変更しない
- 次Issue以降で画面遷移や抽選処理を分離しやすくする

---

## Scope

### In Scope

- `court_reserv/browser/login.py` の追加
- `LoginService` または同等クラスの追加
- 既存ログイン処理の最小限の切り出し
- CAPTCHA / reCAPTCHA 検知時の手動待機方針の維持
- `Court_Reserv` 側の最小限の呼び出し変更
- docs/ARCHITECTURE.md 更新
- docs/DEVELOPMENT.md 更新
- Completion Report 更新

### Out of Scope

- 予約処理の切り出し
- 抽選処理の切り出し
- 画面遷移処理の切り出し
- Selenium 4 移行
- `find_element_by_*` 置換
- UI変更
- 自動予約機能追加
- CAPTCHA / reCAPTCHA 回避・突破・自動認証

---

## Proposed Design

`court_reserv/browser/login.py` に以下の責務を持たせる。

```text
LoginService
    - login(driver, user_id, password)
    - wait_for_manual_captcha_if_needed(driver)
```

または既存コードとの整合上、以下のような形でもよい。

```python
login_service = LoginService(wait_factory=session.get_wait)
login_service.login(driver, user_id, password)
```

このIssueでは既存挙動を優先し、インターフェースは最小限でよい。

---

## Automation Policy

- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない
- CAPTCHA / reCAPTCHA が表示された場合は、ユーザーによる手動認証を待つ
- 手動認証完了後は既存フローを継続する
- 外部CAPTCHAサービスは利用しない

---

## Implementation Plan

1. `court_reserv/browser/login.py` を追加する
2. 既存のログイン処理を `LoginService` に移す
3. CAPTCHA / reCAPTCHA 手動待機方針を維持する
4. `Court_Reserv` から `LoginService` を呼び出すように最小修正する
5. ログイン以外の処理は変更しない
6. docs/ARCHITECTURE.md を更新する
7. docs/DEVELOPMENT.md を更新する
8. Verification を実行する
9. Completion Report を記入する

---

## Target Files

- court_reserv/browser/login.py
- court_reserv/browser/__init__.py
- court_reserv/court_reserv.py
- docs/ARCHITECTURE.md
- docs/DEVELOPMENT.md
- docs/issues/phase1/0008-login-service-separation.md

---

## Tasks

- [x] `browser/login.py` を追加する
- [x] `LoginService` を追加する
- [x] ログイン処理を最小限切り出す
- [x] CAPTCHA / reCAPTCHA 手動待機方針を維持する
- [x] `Court_Reserv` 側を最小限修正する
- [x] docs を更新する
- [x] 動作確認を行う
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] `court_reserv/browser/login.py` が追加されている
- [x] ログイン処理が `LoginService` から利用されている
- [x] 既存ログイン挙動を変更していない
- [x] 予約処理を変更していない
- [x] 抽選処理を変更していない
- [x] UI挙動を変更していない
- [x] Selenium 4 移行を行っていない
- [x] `find_element_by_*` 置換を行っていない
- [x] CAPTCHA / reCAPTCHA 回避・突破・自動認証を実装していない
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

実ブラウザログイン確認ができる場合は、以下も確認する。

```text
ログイン画面へ遷移できること
CAPTCHA / reCAPTCHA が表示された場合は手動待機になること
```

GUI起動確認や実ログイン確認が難しい場合は、理由を Completion Report に記載する。

---

## Notes

- このIssueはログイン処理の分離のみを扱う。
- Navigation Service の切り出しは次Issue以降で行う。
- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない。

---

# Completion Report

※ Codex が記入

## Summary

- `court_reserv/browser/login.py` を追加し、既存のログイン処理と CAPTCHA / reCAPTCHA 手動待機方針を `LoginService` に集約した。
- `court_reserv/court_reserv.py` では `LoginService` への委譲に留め、予約処理・抽選処理・画面遷移・Tkinter UI は変更していない。
- `docs/ARCHITECTURE.md` と `docs/DEVELOPMENT.md` に Login Service 分離方針を追記した。

---

## Changed Files

- `court_reserv/browser/__init__.py`
- `court_reserv/browser/login.py`
- `court_reserv/court_reserv.py`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase1/0008-login-service-separation.md`

---

## Verification Result

- `git status` 実行
- `python -m compileall .` 成功
- `python setup.py --name` で `court_reserv` を確認
- GUI起動確認 (`python court_reserv/court_reserv.py`, `python -m court_reserv.ui.app`) は実行環境依存のため未実施
- 実ブラウザログイン確認は手元環境依存のため未実施

---

## Follow-up Items

- Navigation Service の切り出し
- 抽選申込み系サービスの分離
- 予約確定系サービスの分離
