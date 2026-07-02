# Architecture

## Current State

現状の実装は `court_reserv/court_reserv.py` に以下の責務が集中しています。

- Tkinter UI
- Selenium 操作
- ログインと画面遷移
- 抽選申込み、結果確認、予約確定
- 空き枠収集
- CSV 出力

補助的に `court_reserv/manage_id.py` が ID CSV の読込・書出しと有効確認を担っています。検証用補助スクリプトは Issue 0005 で削除され、現在の正式な起動パスは GUI のみです。

主なエントリーポイント:

- `court_reserv/court_reserv.py`

## Future Architecture

将来的には以下の責務分離を目指します。

```text
court_reserv/
├── browser/   Selenium操作、ログイン、画面遷移
├── services/  抽選、予約、空き確認などの業務ロジック
├── models/    利用者、施設、予約枠などのデータ構造
├── ui/        Tkinter UI
├── utils/     CSV、日付処理、共通ユーティリティ
└── config/    設定読込、ローカル設定、環境変数連携
```

この Issue では構成方針を明文化するだけに留め、既存の Selenium 実装や Tkinter UI、予約ロジックは移動しません。

## Directory Responsibilities

- `browser/`
  Selenium ドライバ操作、ログイン、画面遷移、DOM 操作の分離先
- `services/`
  抽選申込み、抽選結果確認、予約確定、空き枠収集などの業務ロジックの分離先
- `models/`
  利用者、施設、予約枠、抽選結果などのデータ構造の分離先
- `ui/`
  Tkinter ベース UI と入力イベント制御の分離先
- `utils/`
  CSV、ファイル入出力、整形処理、共通補助関数の分離先
- `config/`
  `config.ini`、`config.local.ini`、`.env`、環境変数を統合する設定読込の責務を持つ

## Service Split Policy

- `Court_Reserv` の責務分離は後続 Issue で段階的に行う
- 先に薄いインターフェースと責務境界を整え、その後で実装を移す
- 振る舞い変更を伴う整理は、必ず個別 Issue で扱う
- Selenium の操作順序、Tkinter の見た目、予約ロジックは明示的な Issue が出るまで維持する

## Automation Policy

東京都スポーツ施設予約システムの自動化では、以下を共通ルールとします。

- CAPTCHA / reCAPTCHA の回避、突破、自動認証は実装しない
- 外部 CAPTCHA サービスは利用しない
- reCAPTCHA が表示された場合は、ユーザーが手動で認証するまで待機する
- 手動認証完了後は、そのまま既存処理を継続できる設計を維持する
- 今後追加する機能も、この方針を前提に設計する
