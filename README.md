# Court Reserv

東京都スポーツ施設予約システム向けの個人利用ツールです。現在の実装は Python + Selenium + Tkinter を中心に構成されており、抽選申込み、抽選結果確認、予約確定、空き枠確認を補助します。

Phase 1 では、設定管理、ドキュメント整備、Browser Layer / Service Layer / Model Foundation、Selenium 4 移行までを完了しました。Phase 2 では、空き施設予約よりも先に、抽選申込み自動化の dry-run、抽選申込み画面での枠収集と候補選択、確認付き送信、抽選結果確認、予約確定補助を優先して進めます。

## Scope

- 個人利用を前提とした補助ツールです
- 既存の Selenium / Tkinter ベース実装を段階的に整理していきます
- 開発は `docs/issues/` 配下の Issue Driven Development で進めます

## Current Entry Points

- GUI: `python court_reserv/court_reserv.py`
- GUI module entrypoint: `python -m court_reserv.ui.app`

## Current Architecture Snapshot

- `court_reserv/browser/`
  Selenium セッション、ログイン、共通画面遷移を担当
- `court_reserv/services/`
  抽選、予約、空き確認、ID 管理の業務フローを担当
- `court_reserv/models/`
  Phase 2 に向けた基本モデルを提供
- `court_reserv/ui/`
  Tkinter UI のエントリーポイントと表示制御を担当
- `court_reserv/legacy/`
  旧実装の退避先。現行フローの主経路では使用しない

## Directory Overview

```text
court_reserv/
  browser/    Selenium 関連の分離先
  config/     設定関連の分離先
  models/     データ構造の分離先
  services/   業務ロジックの分離先
  ui/         Tkinter UI の分離先
  utils/      共通処理の分離先
docs/
  issues/     Issue 管理
tests/        テスト
logs/         ログ出力先
output/       CSV 等の出力先
scripts/      将来の補助スクリプト置き場
```

## Setup

1. Python と Google Chrome を用意します
2. `requirements.txt` を元に依存関係をインストールします
3. リポジトリ同梱の `court_reserv/config.ini` を安全なベース設定として確認します
4. 必要に応じて `config.local.ini` または `.env` を作成し、ローカル固有の設定や認証情報を上書きします
5. WebDriver は Selenium Manager に委譲するため、ChromeDriver の固定パス設定は不要です

## Launch

- 既存 GUI 起動: `python court_reserv/court_reserv.py`
- 新しい module 起動: `python -m court_reserv.ui.app`
- GUI では `抽選申込み自動化` / `抽選結果確認` / `予約確定補助` / `設定` を利用できます
- GUI 下部の Log Window では Workflow の標準出力と logging をリアルタイム表示します
- Phase 2 dry-run: `python scripts/lottery_automation_dry_run.py --preferences config/preferences.example.yaml --dry-run`
- Lottery entry workflow: `python scripts/lottery_entry_workflow.py --preferences config/preferences.example.yaml`
  指定曜日の枠と現在申込数を収集し、IDごとの申込み予定枠を表示したうえで `yes` と入力した場合のみ最終送信します。送信後に reCAPTCHA が表示された場合は手動認証を待ち、保存済み選択情報から申込み枠選択を復元して最大1回だけ再送信します
- Lottery result workflow: `python scripts/lottery_result_workflow.py`
  抽選結果を一覧表示し、`output/lottery_automation/` に JSON / CSV 保存します
- Reservation confirmation assist: `python scripts/reservation_confirmation_workflow.py`
  当選一覧を表示し、選択したアカウントだけを `yes` 確認後に予約確定します

`.env` では次のような値を設定できます。

```bash
COURT_RESERV_USER_ID=
COURT_RESERV_PASSWORD=
COURT_RESERV_LOG_PATH=
COURT_RESERV_OUTPUT_CSV_PATH=
COURT_RESERV_TOP_URL=
COURT_RESERV_LOG_LEVEL=INFO
```

Selenium 4.6 以降では Selenium Manager が ChromeDriver の検出と取得を担当します。通常は `webdriver.Chrome(options=...)` のみで起動します。

## Current Status

- [x] Project foundation
- [x] Security and configuration cleanup
- [x] Architecture foundation
- [x] Legacy cleanup
- [x] Project metadata cleanup
- [x] Court_Reserv split phase 1
- [x] Browser layer extraction
- [x] Service layer extraction
- [x] ID management separation
- [x] Model foundation
- [x] Selenium 4 migration
- [x] Phase 1 wrap-up
- [x] Lottery automation dry-run and entry selection workflow
- [x] Lottery submission confirmation workflow
- [x] Lottery result workflow
- [x] Reservation confirmation assist workflow
- [x] Lottery entry slot collection workflow
- [x] Lottery submission recovery
- [x] GUI integration and settings
- [ ] Vacant facility reservation improvements
- [ ] Notification and operations

## Phase 2 Focus

- 抽選申込み自動化の dry-run と候補抽出
- 希望条件と順位付けの整備
- 抽選申込みワークフローの候補自動選択
- 抽選申込み前の確認付き送信
- 抽選申込み画面の枠収集と current entry count 表示
- 送信後 reCAPTCHA の手動 recovery
- 抽選結果確認 workflow
- 予約確定補助 workflow
- 空き施設予約は低優先度
- 通知とスケジューラは Phase 2 の対象外

## Authentication Priority

抽選申込み自動化系の CLI では、認証情報を次の優先順位で解決します。

1. `IdManagerService` から読む ID CSV
2. `config.local.ini`
3. `.env`

希望条件ファイルには ID / password を含めません。

`config/preferences.example.yaml` では、抽選申込み用に次の設定を使えます。

- `lottery.target_weekdays`
  未指定時は `土` を対象にします
- `lottery.default_entries`
  全ID共通の申込み予定枠です
- `lottery.account_overrides`
  指定IDだけ申込み予定枠を上書きします
- `lottery.max_entries_per_account`
  1IDあたり最大2枠までに制限します

## Reservation Confirmation Note

既存 `ReservationService` の制約により、予約確定補助はアカウント単位で動作します。

- 当選一覧は行単位で表示します
- ただし確定実行は選択アカウント単位です
- 同じアカウントの当選を一部だけ確定する操作は、この段階では対応しません

## Lottery Guide

- 抽選申込みは利用前月 1 日 0 時から 10 日 23 時 59 分まで
- 抽選結果確認は利用前月 14 日 0 時以降
- 当選した場合は利用前月 20 日 23 時 59 分までに確認と当選施設の利用申込みが必要
- 1 回の抽選につき、種目ごとに 2 件まで
- 空き施設予約は利用前月 22 日から利用開始時刻までだが、Phase 2 では低優先度

## Development Policy

- 設計判断は ChatGPT が担当します
- Codex は Issue に記載された範囲だけを実装します
- Issue に書かれていない仕様変更は提案までとし、勝手に実装しません
- 既存動作は Issue に明記がない限り変更しません

詳細は [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) と [docs/issues/README.md](docs/issues/README.md) を参照してください。
