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
- Navigation 分離 Issue では、JavaScript 呼び出しと共通画面遷移だけを切り出し、抽選・予約・空き確認ロジックは変更しない
- Lottery Service 分離 Issue では、抽選申込み系フローだけを切り出し、予約確定・空き確認・ID 管理は変更しない
- Reservation Service 分離 Issue では、予約確定・予約確認だけを切り出し、抽選・空き確認・ID 管理は変更しない
- Availability Service 分離 Issue では、空き確認・空き枠収集だけを切り出し、抽選・予約確定・ID 管理は変更しない
- IdManager Service 整理 Issue では、ID管理・CSV 入出力・ID有効確認だけを整理し、抽選・予約確定・空き確認は変更しない
- Model Foundation Issue では、基本 dataclass の追加だけを行い、既存サービスへの大規模適用は行わない

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

## Navigation Service

- JavaScript 実行と共通画面遷移は `court_reserv/browser/navigation.py` に集約する
- `Court_Reserv` 側には最小限の委譲だけを残し、業務ロジックの順序や条件分岐は維持する
- 画面要素の詳細なラップやページ単位の再設計は別 Issue で扱う

## Lottery Service

- 抽選申込み、抽選申込み状況確認、抽選当選結果確認は `court_reserv/services/lottery.py` に集約する
- `Court_Reserv` 側には UI イベントハンドラと最小限の委譲だけを残す
- 予約確定処理や空き確認処理の切り出しは別 Issue で扱う

## Reservation Service

- 予約確定、予約確認は `court_reserv/services/reservation.py` に集約する
- `Court_Reserv` 側には UI イベントハンドラと最小限の委譲だけを残す
- 空き確認処理や ID 管理処理の切り出しは別 Issue で扱う

## Availability Service

- 空き確認、空き枠収集は `court_reserv/services/availability.py` に集約する
- `Court_Reserv` 側には UI イベントハンドラと最小限の委譲だけを残す
- 自動予約機能や予約戦略エンジンの追加は別 Issue で扱う

## IdManager Service

- ID管理、CSV 入出力、ID有効確認は `court_reserv/services/id_manager.py` に整理する
- `manage_id.py` は既存互換のため残し、当面はサービスへの委譲ラッパーとして扱う
- CSV フォーマット変更は別 Issue がない限り行わない

## Models

- `court_reserv/models/` には `Account`、`Facility`、`Slot`、`ReservationPreference` などの基本 dataclass を置く
- モデル基盤 Issue では、既存コードへの全面適用ではなく import 可能な土台作りを優先する
- 既存辞書構造や CSV フォーマットは別 Issue がない限り維持する

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
