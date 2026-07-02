# Issue 0005: Project Metadata Cleanup

## Status

- [ ] Draft
- [ ] Ready
- [ ] In Progress
- [ ] Review
- [x] Done

---

## Summary

分割リファクタリングに入る前に、プロジェクトメタデータ、README、Makefile、requirements、config.ini 管理方針、不要な検証用スクリプトを整理する。

---

## Background

Issue 0001〜0004 により、基盤・設定管理・アーキテクチャ方針・legacy整理は完了した。

一方で、以下が残っている。

- `scripts/run_collect_slots.py` は検証用であり、今後は不要
- README の Current Status が古い
- `Makefile` / `requirements.txt` の実態確認が必要
- `court_reserv/config.ini` のGit管理方針が曖昧
- `.gitignore` と実ファイル管理の整合確認が必要

---

## Goal

- 不要な検証用スクリプトを削除する
- README / CHANGELOG / docs を現在状態に合わせる
- Makefile / requirements.txt を現行プロジェクトに合わせる
- config.ini の管理方針を明確化する
- 次Issueから `Court_Reserv` 分割に入れる状態にする

---

## Scope

### In Scope

- `scripts/run_collect_slots.py` の削除
- `README.md` の Current Status 更新
- `CHANGELOG.md` の更新
- `Makefile` の確認・整理
- `requirements.txt` の確認・整理
- `court_reserv/config.ini` のGit管理方針整理
- `.gitignore` の整合確認
- `docs/DEVELOPMENT.md` 更新
- `docs/issues/phase1/0005-project-metadata-cleanup.md` Completion Report 更新

### Out of Scope

- `Court_Reserv` クラス分割
- Selenium 4 移行
- ログイン処理変更
- 予約処理変更
- GUI変更
- 自動予約機能追加
- CAPTCHA / reCAPTCHA 処理変更

---

## Implementation Plan

1. `scripts/run_collect_slots.py` を削除する
2. `README.md` の Current Status を更新する
3. `CHANGELOG.md` に Issue 0001〜0005 の進捗を追記する
4. `Makefile` を現行構成に合わせて確認・整理する
5. `requirements.txt` を現行構成に合わせて確認・整理する
6. `court_reserv/config.ini` をデフォルト設定として管理するか判断する
7. `.gitignore` と実ファイル管理の整合を取る
8. `docs/DEVELOPMENT.md` に metadata / config 管理方針を追記する
9. Verification を実行する
10. Completion Report を記入する

---

## Target Files

- scripts/run_collect_slots.py
- README.md
- CHANGELOG.md
- Makefile
- requirements.txt
- court_reserv/config.ini
- config.example.ini
- .gitignore
- docs/DEVELOPMENT.md
- docs/issues/phase1/0005-project-metadata-cleanup.md

---

## Tasks

- [x] `scripts/run_collect_slots.py` を削除する
- [x] README の Current Status を更新する
- [x] CHANGELOG を更新する
- [x] Makefile を整理する
- [x] requirements.txt を整理する
- [x] config.ini の管理方針を整理する
- [x] .gitignore を整理する
- [x] docs/DEVELOPMENT.md を更新する
- [x] 動作確認を行う
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] `scripts/run_collect_slots.py` が削除されている
- [x] README の Current Status が現状と一致している
- [x] CHANGELOG が更新されている
- [x] Makefile が存在する場合、現行構成と矛盾していない
- [x] requirements.txt が現行構成と矛盾していない
- [x] config.ini の管理方針が明確になっている
- [x] .gitignore と実ファイル管理が整合している
- [x] 既存のGUIアプリ起動パスを壊していない
- [x] Selenium / Tkinter / 予約ロジックを変更していない
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
make test
```

`pytest` が未インストールの場合は、Completion Report に記載する。

---

## Notes

- このIssueは分割リファクタリング前の最終整理である。
- `scripts/run_collect_slots.py` は検証用スクリプトのため削除する。
- 自動予約機能は今後、正式な service 層として再設計する。
- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない。

---

# Completion Report

※ Codex が記入

## Summary

不要になった検証用スクリプト `scripts/run_collect_slots.py` を削除し、README のエントリーポイントと Current Status、CHANGELOG、Makefile、requirements.txt を現行構成に合わせて整理した。GUI の正式起動パス `python court_reserv/court_reserv.py` は維持している。

また、`court_reserv/config.ini` を秘密情報を含まない Git 管理のベース設定として位置づけ直し、`.gitignore`、`config.example.ini`、`docs/DEVELOPMENT.md`、`docs/SECURITY.md` の方針を整合させた。Selenium / Tkinter / 予約ロジック、ログイン処理、CAPTCHA / reCAPTCHA 処理には変更を加えていない。

---

## Changed Files

- `README.md`
- `CHANGELOG.md`
- `Makefile`
- `requirements.txt`
- `court_reserv/config.ini`
- `config.example.ini`
- `.gitignore`
- `docs/DEVELOPMENT.md`
- `docs/SECURITY.md`
- `docs/issues/phase1/0005-project-metadata-cleanup.md`
- `scripts/run_collect_slots.py` (deleted)

---

## Verification Result

- `git status` を実行し、対象ファイルの変更を確認
- `python -m compileall .` を実行し、構文エラーがないことを確認
- `python setup.py --name` を実行し、`court_reserv` が返ることを確認
- `make test` を実行し、Makefile が `python -m pytest` を呼ぶことを確認
- この環境では `pytest` モジュール未導入のため `make test` は未通過

---

## Follow-up Items

- `Court_Reserv` 分割開始時に、README の Current Status と CHANGELOG を再度見直す
- テスト実行環境を整えて `make test` が常に実行可能な状態にする
