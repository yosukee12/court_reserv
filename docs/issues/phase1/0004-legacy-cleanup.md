# Issue 0004: Legacy Cleanup

## Status

- [ ] Draft
- [ ] Ready
- [ ] In Progress
- [ ] Review
- [x] Done

---

## Summary

現在のプロジェクトに残っている旧実装・検証用スクリプト・テンプレート由来ファイルを整理し、今後の責務分離に向けて不要なノイズを減らす。

---

## Background

Issue 0001〜0003 により、プロジェクト基盤、設定管理、アーキテクチャ方針は整備された。

一方で、現在のリポジトリには以下の整理対象が残っている。

- `run_collect_slots.py`
- `old_court_reserv.py`
- `setup.py`
- テンプレート由来の `tests/`
- debug出力先の重複
- `docs/_build/` など生成物

これらは今後のリファクタリング時に混乱の原因になるため、先に扱いを明確化する。

---

## Goal

- 旧実装・検証用ファイルの扱いを明確にする
- 不要なテンプレート由来ファイルを整理する
- debug出力先を整理する
- package / setup 周りの不整合を解消する
- 既存のGUIアプリ挙動は変更しない

---

## Scope

### In Scope

- `run_collect_slots.py` の扱い整理
  - 不要なら削除
  - 残す場合は `scripts/` または `legacy/` へ移動
  - 正式機能化はしない
- `old_court_reserv.py` の扱い整理
  - 削除せず、必要なら `court_reserv/legacy/` へ移動
- `setup.py` の `README.rst` 参照不整合修正
  - `README.md` 参照に変更
  - または不要なら削除
- テンプレート由来の無効な `tests/` 整理
- debug出力先を `output/debug_pages/` に一本化
- `.gitignore` の必要な見直し
- docs/DEVELOPMENT.md の更新
- Completion Report の記入

### Out of Scope

- `Court_Reserv` クラスの分割
- Selenium 4 移行
- ログイン処理変更
- 予約処理変更
- GUI変更
- 自動予約機能追加
- CAPTCHA / reCAPTCHA 処理変更

---

## Implementation Plan

1. `run_collect_slots.py` の参照有無を確認する
2. 正式運用で不要な検証用スクリプトであれば削除、または `legacy/` へ移動する
3. `old_court_reserv.py` を削除せず `court_reserv/legacy/old_court_reserv.py` へ移動する
4. `setup.py` の `README.rst` 参照を `README.md` に修正する
5. テンプレート由来の無効なテストを削除または置き換える
6. debug出力先を `output/debug_pages/` に一本化する
7. `.gitignore` を必要に応じて更新する
8. docs/DEVELOPMENT.md に legacy / scripts の扱いを追記する
9. `python -m compileall .` を実行する
10. Completion Report を記入する

---

## Target Files

- run_collect_slots.py
- scripts/
- court_reserv/legacy/
- court_reserv/old_court_reserv.py
- setup.py
- tests/
- .gitignore
- docs/DEVELOPMENT.md
- docs/issues/phase1/0004-legacy-cleanup.md

---

## Tasks

- [x] `run_collect_slots.py` の扱いを決める
- [x] `old_court_reserv.py` を legacy に移動する
- [x] `setup.py` の README 参照不整合を修正する
- [x] 無効なテンプレートテストを整理する
- [x] debug出力先を `output/debug_pages/` に整理する
- [x] `.gitignore` を見直す
- [x] docs/DEVELOPMENT.md を更新する
- [x] `python -m compileall .` を実行する
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] 既存のGUIアプリ起動パスを壊していない
- [x] Selenium / Tkinter / 予約ロジックを変更していない
- [x] `old_court_reserv.py` の扱いが明確になっている
- [x] `run_collect_slots.py` の扱いが明確になっている
- [x] `setup.py` が存在しない `README.rst` を参照していない
- [x] テンプレート由来の無効なテストが残っていない
- [x] debug出力先が整理されている
- [x] `python -m compileall .` が成功する
- [x] CAPTCHA / reCAPTCHA 方針を変更していない

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
pytest
```

`pytest` はテストを整理した結果、テストがない場合は無理に通す必要はない。  
その場合は Completion Report に理由を記載する。

---

## Notes

- このIssueは不要ファイル整理とlegacy整理が目的。
- `old_court_reserv.py` は削除せず legacy へ移動することを優先する。
- `run_collect_slots.py` は検証用であれば削除または legacy 移動でよい。
- 正式なCLI機能化は次Issue以降で検討する。
- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない。

---

# Completion Report

※ Codex が記入

## Summary

旧実装と検証用スクリプトの扱いを明確化し、`old_court_reserv.py` を `court_reserv/legacy/old_court_reserv.py` へ移動した。`run_collect_slots.py` は正式機能ではない検証用スクリプトとして `scripts/run_collect_slots.py` へ移動し、GUI の正式起動パスとは切り離した。

あわせて `setup.py` の `README.rst` 参照不整合を修正し、テンプレート由来で無効だったテストを削除して最小の smoke test に整理した。debug 出力先は `output/debug_pages/` に一本化し、関連ドキュメントと ignore 設定を更新した。

---

## Changed Files

- `scripts/run_collect_slots.py`
- `court_reserv/legacy/__init__.py`
- `court_reserv/legacy/old_court_reserv.py`
- `court_reserv/court_reserv.py`
- `court_reserv/config/loader.py`
- `README.md`
- `setup.py`
- `tests/test_smoke.py`
- `.gitignore`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase1/0004-legacy-cleanup.md`

---

## Verification Result

- `git status` を実行し、対象ファイルの変更を確認
- `python -m compileall .` を実行し、構文エラーがないことを確認
- `python setup.py --name` を実行し、`court_reserv` が返ることを確認
- `pytest` はコマンドが環境に存在せず未実行のため、`tests/test_smoke.py` を `compileall` 対象として構文確認した

---

## Follow-up Items

- `scripts/run_collect_slots.py` を将来も維持するか削除するかは後続 Issue で判断
- `Makefile` や `requirements.txt` に残る template / legacy 由来設定の整理は後続 Issue で対応
