# Issue 0015: Selenium 4 Migration

## Status

Done

---

## Summary

Selenium 3系の旧APIや不安定な待機処理を整理し、Selenium 4互換の実装へ移行する。

このIssueでは Selenium API の移行と待機処理の整理のみを行い、抽選・予約・空き確認・ID管理・UIの仕様変更は行わない。

---

## Background

Issue 0007〜0014 により、Browser / Service / Model 層の基本的な分離が進んだ。

現在のコードには以下のような課題が残っている。

- `find_element_by_*` など Selenium 3系の旧APIが残っている
- `time.sleep()` に依存する待機処理が多い
- `except TimeoutException or UnexpectedAlertPresentException:` のような不正確な例外処理が残っている可能性がある
- Selenium操作が複数サービスに分散しているため、今後の自動予約前に安定化が必要

---

## Goal

- `find_element_by_*` を Selenium 4 の `find_element(By.*)` に置き換える
- 可能な範囲で `time.sleep()` を `WebDriverWait` に置き換える
- 例外処理を正しい書き方に修正する
- 既存の抽選・予約・空き確認・ログイン挙動を維持する
- 自動予約機能に入る前の Selenium 基盤を安定化する

---

## Scope

### In Scope

- Selenium 3系旧APIの置換
- `By` の import 整理
- 明らかに置換可能な `time.sleep()` の `WebDriverWait` 化
- Selenium例外処理の修正
- `BrowserSession` の待機補助の活用
- `docs/ARCHITECTURE.md` 更新
- `docs/DEVELOPMENT.md` 更新
- `docs/issues/phase1/0015-selenium4-migration.md` Completion Report 更新

### Out of Scope

- 抽選処理の仕様変更
- 予約確定処理の仕様変更
- 空き確認処理の仕様変更
- ID管理処理の仕様変更
- UI変更
- 自動予約機能追加
- 予約戦略エンジン追加
- CAPTCHA / reCAPTCHA 処理方針変更
- 大規模なサービス再分割

---

## Implementation Plan

1. `find_element_by_*` の残存箇所を確認する
2. Selenium 4形式へ置換する
3. `By` import を整理する
4. 明らかに待機条件が分かる `time.sleep()` を `WebDriverWait` へ置換する
5. 置換が危険な `time.sleep()` は無理に変更せず、理由を Completion Report に記載する
6. 不正確な例外処理を修正する
7. `python -m compileall .` を実行する
8. 可能なら `pytest` または `make test` を実行する
9. docs を更新する
10. Completion Report を記入する

---

## Target Files

- court_reserv/court_reserv.py
- court_reserv/browser/session.py
- court_reserv/browser/login.py
- court_reserv/browser/navigation.py
- court_reserv/services/lottery.py
- court_reserv/services/reservation.py
- court_reserv/services/availability.py
- court_reserv/services/id_manager.py
- docs/ARCHITECTURE.md
- docs/DEVELOPMENT.md
- docs/issues/phase1/0015-selenium4-migration.md

---

## Tasks

- [ ] `find_element_by_*` の残存確認
- [ ] Selenium 4 APIへ置換
- [ ] `By` import 整理
- [ ] 明らかに安全な `time.sleep()` を `WebDriverWait` 化
- [ ] 例外処理を修正
- [ ] docs を更新
- [ ] compileall 実行
- [ ] pytest / make test 実行
- [ ] Completion Report 更新

---

## Acceptance Criteria

- [ ] `find_element_by_*` が残っていない
- [ ] Selenium 4 の `By.*` 形式が使われている
- [ ] 不正確な Selenium 例外処理が修正されている
- [ ] 既存の抽選処理の仕様を変更していない
- [ ] 既存の予約処理の仕様を変更していない
- [ ] 既存の空き確認処理の仕様を変更していない
- [ ] 既存のログイン処理の仕様を変更していない
- [ ] UI挙動を変更していない
- [ ] CAPTCHA / reCAPTCHA 方針を変更していない
- [ ] `python -m compileall .` が成功する

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
make test
```

残存確認：

```bash
grep -R "find_element_by_" -n court_reserv || true
grep -R "except TimeoutException or" -n court_reserv || true
grep -R "except .* or .*Exception" -n court_reserv || true
```

可能なら以下も実行する。

```bash
python court_reserv/court_reserv.py
python -m court_reserv.ui.app
```

GUI起動確認や実ブラウザ確認が難しい場合は、理由を Completion Report に記載する。

---

## Notes

- このIssueは Selenium 4 移行のみを対象とする。
- `time.sleep()` はすべて無理に置換しなくてよい。
- 画面遷移後の安定待ちなど、意図がある `sleep` は残してよい。
- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない。
- reCAPTCHA が表示された場合は、ユーザーによる手動認証待機を維持する。

---

# Completion Report

※ Codex が記入

## Summary

- `services/availability.py` に残っていた `find_element_by_name()` を `find_element(By.NAME, ...)` へ置換した
- `services/lottery.py` と `legacy/old_court_reserv.py` の不正確な `except TimeoutException or UnexpectedAlertPresentException:` を tuple 形式へ修正した
- `browser/navigation.py` の `select_lottery_tennis_park()` にあった固定 `sleep(1)` を、`bname` 要素の出現待ちに置き換えた
- `legacy/old_court_reserv.py` に残っていた Selenium 3 形式も、検索残りが出ないよう互換的に置換した
- `docs/ARCHITECTURE.md` と `docs/DEVELOPMENT.md` に Selenium 4 移行方針を追記した

## Changed Files

- `court_reserv/browser/navigation.py`
- `court_reserv/services/availability.py`
- `court_reserv/services/lottery.py`
- `court_reserv/legacy/old_court_reserv.py`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase1/0015-selenium4-migration.md`

## Verification Result

- `git status --short`
  変更ファイルのみ表示された
- `python -m compileall .`
  成功
- `python setup.py --name`
  `court_reserv`
- `make test`
  失敗。実行環境に `pytest` が入っておらず `No module named pytest` で停止
- `grep -R "find_element_by_" -n court_reserv || true`
  `__pycache__` のバイナリ一致のみ検出。ソース確認のため `rg -n "find_element_by_" court_reserv` を補助実行し、ソース一致なしを確認
- `grep -R "except TimeoutException or" -n court_reserv || true`
  一致なし
- `grep -R "except .* or .*Exception" -n court_reserv || true`
  一致なし
- GUI 起動確認
  この環境では未確認のため、可能なら `python court_reserv/court_reserv.py` と `python -m court_reserv.ui.app` を別途実施する

## Follow-up Items
- `time.sleep()` が多数残っているが、画面更新完了条件がコード上で明確でない箇所は今回無理に変更していない
- 具体例として、ログイン送信直前の短時間待機、抽選・予約画面での手動操作前後待機、ID有効確認画面の遷移待機、空き枠HTML収集のフォールバック待機は、画面状態だけで安全に条件化しづらいため残した
- Selenium 待機の追加整理は、抽選・予約・空き確認それぞれの専用 Issue で挙動確認付きで段階的に進める
