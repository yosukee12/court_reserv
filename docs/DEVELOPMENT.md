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

## Phase Status

- Phase 1 は Issue 0016 で完了とする
- Phase 2 では、既存挙動を維持したまま抽選申込み自動化を段階的に追加する
- 空き施設予約は低優先度とし、Phase 2 では主対象にしない
- Phase 2 の Issue でも、仕様変更を伴う場合は Scope と Out of Scope を明示する

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

## Selenium 4 Migration

- Selenium 4 移行では `find_element_by_*` を `find_element(By.*)` へ置換する
- 例外を複数扱う場合は `except (A, B):` を使い、`except A or B:` は使わない
- `time.sleep()` は安全に条件化できる箇所だけ `WebDriverWait` へ置換する
- ページ描画や外部システム都合で意図が不明な待機は残し、理由を Completion Report に記載する

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

## Phase 2 Working Rules

- Phase 2 では抽選申込み自動化を主対象として小さく進める
- 既存サービス層の安定動作を優先し、仕様変更と構造整理を同じ Issue で混在させない
- CAPTCHA / reCAPTCHA の手動認証待機方針は継続する
- 空き施設予約は低優先度として扱う

## Lottery Guide Rules

- 抽選申込みは利用前月 1 日 0 時から 10 日 23 時 59 分まで
- 抽選結果確認は利用前月 14 日 0 時以降
- 当選した場合は利用前月 20 日 23 時 59 分までに確認および当選施設の利用申込みが必要
- 1 回の抽選につき、種目ごとに 2 件まで
- 空き施設予約は利用前月 22 日から利用開始時刻までだが、Phase 2 では低優先度
- 当選後の最終確定はユーザー判断を挟み、完全自動確定は行わない

## Speed Priority Policy

- Phase 2 はスピード重視で進め、完璧な抽象化より早く動く構成を優先する
- 新規レイヤを増やしすぎず、既存 `LoginService`、`AvailabilityService`、`LotteryService`、`NavigationService` の再利用を優先する
- dry-run や安全な段階導入を優先し、live 実行は明示的な Issue と確認を前提に進める
- 通知機能とスケジューラ機能は Phase 2 の対象外とする
- CAPTCHA / reCAPTCHA は引き続き手動認証待機で扱う
- Issue 0021 では実際の抽選申込み送信を行わず、候補抽出と順位付けだけを dry-run で確認する

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
