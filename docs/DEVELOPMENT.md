# Development

## Issue Driven Development

このリポジトリでは Issue Driven Development を採用します。

1. `docs/issues/` に Issue を作成する
2. Scope / Out of Scope / Acceptance Criteria を明記する
3. 実装担当は Issue に書かれた内容だけを変更する
4. 実装後は Completion Report を記入する

## Roles

- ChatGPT
  設計判断、方針整理、Issue 化の支援を担当する
- Codex
  Issue に記載された範囲の実装、検証、Completion Report 記入を担当する

## Completion Report

- 実装後は対象 Issue の Completion Report を必ず更新する
- 実装できなかった点や未確認事項は仮実装せず、Completion Report に残す
- Follow-up Items には次 Issue で扱うべき事項だけを記載する

## Implementation Rules

- Issue 外の仕様変更は行わない
- 既存動作は Issue に明記がない限り変更しない
- 大規模リファクタリングは Issue 単位で行う
- 不明点は仮実装せず、Issue または Completion Report に残す
- 小さな PR / 小さな差分を優先する
- Selenium の実装変更、GUI 変更、予約ロジック変更は明示的な Issue がある場合のみ行う
- CAPTCHA / reCAPTCHA の回避や自動認証は実装しない
- reCAPTCHA が表示された場合は手動認証完了まで待機し、その後に処理を継続する前提で実装する
- 既存の正式起動パスを置き換える場合は、互換エントリーポイントを維持するか、専用 Issue で明示的に廃止する
- 基盤切り出し Issue では、WebDriver 生成や待機処理のような共通部だけを分離し、ログイン・予約・画面遷移の中身は変更しない
- ログイン分離 Issue では、ログイン処理と CAPTCHA 手動待機だけを切り出し、抽選・予約・画面遷移ロジックは変更しない

## Local Setup

- `court_reserv/config.ini` は Git 管理する安全なベース設定として扱う
- 環境依存の設定や秘密情報は `config.local.ini` または `.env` で上書きする
- 個人情報や認証情報を含むファイルは Git に追加しない
- 出力物は `logs/`, `output/`, `output/debug_pages/` を利用する

## Legacy And Scripts

- `scripts/` には検証用・補助用スクリプトのみを置く
- `court_reserv/legacy/` には旧実装を退避し、現行実装から参照しない
- 検証用スクリプトは正式機能化しない限り、GUI の正式起動パスとして扱わない
- 不要になった検証用スクリプトは削除し、履歴は Git と `legacy/` で追跡する

## Entrypoints

- GUI の正式起動パスは `python court_reserv/court_reserv.py`
- 新しい module entrypoint を追加する場合でも、互換起動パスは段階的分割が完了するまで維持する
- entrypoint 整理だけを行う Issue では、`Court_Reserv` の中身を大規模に分割しない

## Browser Session

- Selenium の `WebDriver` 生成と終了処理は `court_reserv/browser/session.py` に集約する
- `WebDriverWait` の生成は共通ヘルパーを経由し、タイムアウト値だけを各処理側から渡す
- Browser Session 分離 Issue では、操作順序や DOM 選択式は変更しない

## Login Service

- Selenium のログイン処理と CAPTCHA / reCAPTCHA 手動待機は `court_reserv/browser/login.py` に集約する
- `Court_Reserv` 側には最小限の委譲メソッドだけを残し、既存の予約系フローからの呼び出し形は大きく変えない
- `find_element_by_*` の置換やログイン後ナビゲーションの整理は別 Issue で扱う

## Metadata Policy

- `README.md`, `CHANGELOG.md`, `Makefile`, `requirements.txt` は実態に合わせて保つ
- テンプレート由来の設定やコマンドは残さず、現行構成に合わせて更新する
- `court_reserv/config.ini` には秘密情報を書かず、共有可能な既定値だけを置く

## Useful Commands

```bash
git status
find . -maxdepth 3 -type d | sort
find docs/issues -maxdepth 3 -type f | sort
python -m compileall .
```
