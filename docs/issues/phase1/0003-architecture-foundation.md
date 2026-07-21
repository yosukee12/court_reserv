# Issue 0003: Architecture Foundation

## Status

- [ ] Draft
- [ ] Ready
- [ ] In Progress
- [ ] Review
- [x] Done

---

## Summary

今後の大規模リファクタリングに向けて、Court Reserv App のアーキテクチャ基盤を整備する。

この Issue では設計・ドキュメント・責務整理を中心に行い、既存の Selenium / Tkinter / 予約ロジックの挙動は変更しない。

---

## Background

現在、`Court_Reserv` クラスに以下の責務が集中している。

- Tkinter UI
- Selenium 操作
- ログイン
- 抽選申込
- 抽選結果確認
- 予約確定
- 空き枠収集
- CSV 出力

Issue 0001 でプロジェクト基盤を整備し、Issue 0002 で設定管理を統一した。

次の段階として、今後の責務分離に向けたアーキテクチャと開発ルールを明確化する。

---

## Goal

- 将来の責務分離方針を明文化する
- 各ディレクトリの役割を定義する
- Automation Policy を整備する
- reCAPTCHA の取り扱いをプロジェクト共通ルールとして定義する
- 既存挙動を変更せず、次Issue以降のリファクタリング準備を行う

---

## Scope

### In Scope

- docs/ARCHITECTURE.md 更新
- docs/DEVELOPMENT.md 更新
- docs/issues/README.md 更新
- browser/
- services/
- models/
- ui/
- utils/

各ディレクトリの責務整理

- run_collect_slots.py の固定出力パス整理
- Completion Report 更新

### Out of Scope

- Court_Reserv クラス分割
- Selenium 4 移行
- setup.py 整理
- old_court_reserv.py 整理
- tests 整理
- config.ini 整理
- GUI変更
- ログイン処理変更
- 予約処理変更
- 自動予約機能追加
- CAPTCHA / reCAPTCHA の仕様変更

---

## Architecture Direction

今後の責務分離方針は以下とする。

```text
court_reserv/

browser/
    Selenium操作
    ログイン
    画面遷移

services/
    業務ロジック
    抽選
    予約
    空き確認

models/
    利用者
    施設
    予約枠

ui/
    Tkinter

utils/
    共通ユーティリティ
```

この Issue では上記を実装するのではなく、責務を定義する。

---

## Automation Policy

本プロジェクトでは以下を共通ルールとする。

### CAPTCHA / reCAPTCHA

以下は禁止する。

- CAPTCHA 回避
- reCAPTCHA 回避
- 自動認証
- 外部CAPTCHAサービス利用

reCAPTCHA が表示された場合は

1. ユーザーへ認証を促す
2. 手動認証完了まで待機する
3. 認証完了後は処理を自動再開する

今後実装するすべての機能は、この方針を前提とする。

---

## Implementation Plan

1. docs/ARCHITECTURE.md 更新
2. docs/DEVELOPMENT.md 更新
3. docs/issues/README.md 更新
4. Automation Policy を追加
5. 各ディレクトリの責務を記載
6. run_collect_slots.py の固定出力パス整理
7. Completion Report 更新

---

## Target Files

- docs/ARCHITECTURE.md
- docs/DEVELOPMENT.md
- docs/issues/README.md
- docs/issues/phase1/0003-architecture-foundation.md
- run_collect_slots.py

---

## Tasks

- [x] ARCHITECTURE 更新
- [x] DEVELOPMENT 更新
- [x] Issues README 更新
- [x] Automation Policy 追加
- [x] 各ディレクトリ責務整理
- [x] run_collect_slots.py 固定出力パス整理
- [x] 動作確認
- [x] Completion Report 更新

---

## Acceptance Criteria

- [x] 既存の Selenium の挙動を変更していない
- [x] Tkinter の挙動を変更していない
- [x] 予約ロジックを変更していない
- [x] docs/ARCHITECTURE.md が更新されている
- [x] docs/DEVELOPMENT.md が更新されている
- [x] docs/issues/README.md が更新されている
- [x] Automation Policy が追加されている
- [x] reCAPTCHA の方針が明文化されている
- [x] run_collect_slots.py の固定出力パスが整理されている
- [x] python -m compileall . が成功する

---

## Verification

```bash
git status

python -m compileall .
```

---

## Notes

この Issue は「設計の土台作り」である。

大規模リファクタリングは次Issue以降で実施する。

---

# Completion Report

※ Codex が記入

## Summary

アーキテクチャの土台として、現在構成と将来構成、各ディレクトリの責務、サービス分割方針、Automation Policy をドキュメントへ反映した。`browser/`, `services/`, `models/`, `ui/`, `utils/` には責務境界を示す薄いモジュールコメントを追加し、既存の Selenium 実装は移動していない。

また、`run_collect_slots.py` に残っていた固定出力パスを整理し、共通設定経由の出力先解決へ変更した。既存の Tkinter UI、Selenium の操作順序、予約ロジック、CAPTCHA / reCAPTCHA の処理には変更を加えていない。

---

## Changed Files

- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/README.md`
- `docs/issues/phase1/0003-architecture-foundation.md`
- `court_reserv/browser/__init__.py`
- `court_reserv/services/__init__.py`
- `court_reserv/models/__init__.py`
- `court_reserv/ui/__init__.py`
- `court_reserv/utils/__init__.py`
- `court_reserv/config/__init__.py`
- `court_reserv/config/loader.py`
- `run_collect_slots.py`

---

## Verification Result

- `git status` を実行し、対象ファイルの変更を確認
- `python -m compileall .` を実行し、構文エラーがないことを確認

---

## Follow-up Items

- `Court_Reserv` クラスの責務分離は次 Issue 以降で段階的に対応
- `run_collect_slots.py` の補助スクリプトとしての位置づけ整理は後続 Issue で判断
