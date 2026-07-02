# Issue 0001: Project Foundation

## Status

Done

---

## Summary

Court Reserv App の今後のリファクタリングと自動予約機能追加に向けて、プロジェクトの基本構成・ドキュメント構成・Issue運用・安全な設定管理の土台を整備する。

---

## Background

現在のアプリは、東京都スポーツ施設予約システムにログインし、テニスコートの抽選申込・予約確定・空き確認を行う Python / Selenium / Tkinter ベースの個人利用ツールである。

今後、以下を段階的に進める予定。

- Selenium 4 対応
- 設定管理の改善
- Tkinter UI と Selenium 操作の責務分離
- 空き枠検索・予約戦略エンジンの追加
- 半自動 / 全自動予約モードの追加
- docs / Issue 駆動による開発運用

その前段として、まずは Photo Storage App と同様の Issue 駆動開発に対応できるプロジェクト基盤を作る。

---

## Goal

- 今後のリファクタリングに耐えられるディレクトリ構成を作る
- `docs/issues/` 配下に Issue 管理の仕組みを作る
- 秘密情報や個人情報をコミットしないための最低限のガードを作る
- README / CHANGELOG / docs の初期ファイルを用意する
- 既存アプリの動作は原則変更しない

---

## Scope

### In Scope

- ディレクトリ作成
- Python パッケージ用 `__init__.py` の作成
- `docs/issues/README.md` の作成
- `docs/issues/TEMPLATE.md` の作成
- README / CHANGELOG / docs の初期ファイル作成
- `.gitignore` の整備
- `.env.example` / `config.example.ini` の作成
- 既存コードを壊さない範囲でのファイル配置整理
- 個人情報・認証情報を含むファイルを Git 管理対象から外す準備

### Out of Scope

- Selenium 4 への本格移行
- Tkinter UI と Selenium 処理の分離
- ログイン処理の変更
- 予約処理の変更
- 全自動予約の実装
- CAPTCHA / reCAPTCHA の回避処理
- 大規模な既存コードの分割
- 既存機能の仕様変更

---

## Implementation Plan

1. 以下のディレクトリを作成する。

```text
court_reserv/
├── browser/
├── config/
├── models/
├── services/
├── ui/
└── utils/

docs/
└── issues/
    ├── phase1/
    ├── phase2/
    └── tech-debt/

scripts/
tests/
logs/
output/
```

2. Python パッケージとして扱うため、必要な `__init__.py` を作成する。

3. `docs/issues/README.md` を作成し、Issue運用ルールを定義する。

4. `docs/issues/TEMPLATE.md` を作成し、今後のIssueテンプレートを定義する。

5. 以下のドキュメント初期ファイルを作成する。

- `README.md`
- `CHANGELOG.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/ROADMAP.md`
- `docs/SECURITY.md`

6. `.gitignore` を整備する。

7. `.env.example` と `config.example.ini` を作成する。

8. 既存機能に影響がないことを確認する。

---

## Target Files

- `.gitignore`
- `.env.example`
- `config.example.ini`
- `README.md`
- `CHANGELOG.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/ROADMAP.md`
- `docs/SECURITY.md`
- `docs/issues/README.md`
- `docs/issues/TEMPLATE.md`
- `docs/issues/phase1/0001-project-foundation.md`
- `court_reserv/__init__.py`
- `court_reserv/browser/__init__.py`
- `court_reserv/config/__init__.py`
- `court_reserv/models/__init__.py`
- `court_reserv/services/__init__.py`
- `court_reserv/ui/__init__.py`
- `court_reserv/utils/__init__.py`

---

## Tasks

- [x] ディレクトリ構成を作成する
- [x] Python パッケージ用の `__init__.py` を作成する
- [x] `docs/issues/README.md` を作成する
- [x] `docs/issues/TEMPLATE.md` を作成する
- [x] README / CHANGELOG / docs の初期ファイルを作成する
- [x] `.gitignore` を整備する
- [x] `.env.example` を作成する
- [x] `config.example.ini` を作成する
- [x] 既存コードの動作に影響がないことを確認する
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] 指定されたディレクトリ構成が作成されている
- [x] 必要な `__init__.py` が作成されている
- [x] `docs/issues/README.md` にIssue運用ルールが記載されている
- [x] `docs/issues/TEMPLATE.md` が今後のIssue作成に使える内容になっている
- [x] `.gitignore` に秘密情報・ログ・出力ファイル・debug HTML が含まれている
- [x] `.env.example` に実値ではないサンプル設定が記載されている
- [x] `config.example.ini` に実値ではないサンプル設定が記載されている
- [x] 認証情報・個人情報を新たに追加していない
- [x] CAPTCHA / reCAPTCHA の回避処理を実装していない
- [x] 既存機能を変更していない、または変更が必要な場合は理由が記録されている

---

## Verification

実施した確認内容を記載する。

推奨確認コマンド：

```bash
git status
find . -maxdepth 3 -type d | sort
find docs/issues -maxdepth 3 -type f | sort
python -m compileall .
```

既存アプリが起動できる場合：

```bash
python court_reserv/court_reserv.py
```

---

## Notes

- このIssueでは大規模リファクタリングは行わない。
- 既存ファイル移動が必要な場合も、import修正を伴う大きな変更は次Issue以降に回す。
- `debug_pages/` に個人情報が含まれる可能性があるため、Git管理対象から外す。
- CAPTCHA / reCAPTCHA は手動対応前提とし、回避処理は実装しない。

---

# Completion Report

※Codexが実装完了時に記入

## Summary

プロジェクト基盤整備として、Issue 運用文書、初期ドキュメント、サンプル設定、Git 管理対象の見直し、空ディレクトリの整備を実施した。既存の Selenium / Tkinter / 予約処理ロジックには変更を加えていない。

追加修正として、`.env.example` のサンプル項目更新、README のローカル絶対パスリンク修正と Current Status 追加、`.gitignore` の HTML 除外見直し、`docs/ROADMAP.md` のフェーズ追記を行った。

---

## Changed Files

- `.gitignore`
- `.env.example`
- `README.md`
- `CHANGELOG.md`
- `config.example.ini`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/ROADMAP.md`
- `docs/SECURITY.md`
- `docs/issues/README.md`
- `docs/issues/TEMPLATE.md`
- `docs/issues/phase1/0001-project-foundation.md`
- `docs/issues/phase2/.gitkeep`
- `docs/issues/tech-debt/.gitkeep`
- `logs/.gitkeep`
- `output/.gitkeep`
- `scripts/.gitkeep`

---

## Verification Result

- `git status` で変更対象を確認
- `find . -maxdepth 3 -type d | sort` でディレクトリ構成を確認
- `find docs/issues -maxdepth 3 -type f | sort` で Issue 管理ファイルを確認
- `python -m compileall .` で構文確認を実施

---

## Follow-up Items

- `setup.py` が `README.rst` を参照している不整合は本Issueの対象外のため未対応
- 空の `browser / config / models / services / ui / utils` への実際の責務分離は後続Issueで対応
