# Court Reserv

東京都スポーツ施設予約システム向けの個人利用ツールです。現在の実装は Python + Selenium + Tkinter を中心に構成されており、抽選申込み、抽選結果確認、予約確定、空き枠確認を補助します。

## Scope

- 個人利用を前提とした補助ツールです
- 既存の Selenium / Tkinter ベース実装を段階的に整理していきます
- 開発は `docs/issues/` 配下の Issue Driven Development で進めます

## Current Entry Points

- GUI: `python court_reserv/court_reserv.py`
- GUI module entrypoint: `python -m court_reserv.ui.app`

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
5. ChromeDriver のパスなど、環境依存の設定をローカルで補完します

## Launch

- 既存 GUI 起動: `python court_reserv/court_reserv.py`
- 新しい module 起動: `python -m court_reserv.ui.app`

`.env` では次のような値を設定できます。

```bash
COURT_RESERV_USER_ID=
COURT_RESERV_PASSWORD=
COURT_RESERV_CHROME_DRIVER_PATH=
COURT_RESERV_LOG_PATH=
COURT_RESERV_OUTPUT_CSV_PATH=
COURT_RESERV_TOP_URL=
COURT_RESERV_LOG_LEVEL=INFO
```

## Current Status

- [x] Project foundation
- [x] Security and configuration cleanup
- [x] Architecture foundation
- [x] Legacy cleanup
- [x] Project metadata cleanup
- [ ] Selenium 4 migration
- [ ] Court_Reserv split
- [ ] Service layer extraction
- [ ] Reservation strategy engine
- [ ] Auto reservation
- [ ] Notification and operations

## Development Policy

- 設計判断は ChatGPT が担当します
- Codex は Issue に記載された範囲だけを実装します
- Issue に書かれていない仕様変更は提案までとし、勝手に実装しません
- 既存動作は Issue に明記がない限り変更しません

詳細は [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) と [docs/issues/README.md](docs/issues/README.md) を参照してください。
