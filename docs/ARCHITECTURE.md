# Architecture

## Current State

Phase 1 完了時点では、初期の巨大な `Court_Reserv` クラスから責務分離の土台が整い、主要な Selenium 共通処理と業務フローは専用モジュールへ段階的に切り出されています。

- `court_reserv/court_reserv.py`
  既存 GUI 互換起動口と UI 主体の調停
- `court_reserv/browser/`
  Selenium セッション、ログイン、共通画面遷移
- `court_reserv/services/`
  抽選、予約、空き確認、ID 管理の業務フロー
- `court_reserv/models/`
  Phase 2 に向けた基本モデル
- `court_reserv/ui/app.py`
  module entrypoint

補助的に `court_reserv/manage_id.py` は既存互換ラッパーとして残っています。検証用補助スクリプトは削除または legacy 退避済みで、正式な起動パスは GUI のみです。

主なエントリーポイント:

- `court_reserv/court_reserv.py`
- `court_reserv/ui/app.py`

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

Phase 1 では構成方針の明文化に加えて、Browser / Service / Model の土台構築と Selenium 4 移行までを完了しました。Phase 2 では、この構成を前提に抽選申込み自動化を優先して進めます。

## Phase 1 Outcome

- Browser Layer 分離完了
- Service Layer 分離完了
- Model Foundation 追加完了
- Selenium 4 移行完了
- 旧検証スクリプトと legacy ルート整理完了
- Issue Driven Development と Completion Report 運用定着

## Phase 2 Direction

- `models/` を利用した抽選申込み候補、優先度、利用者設定の表現強化
- サービス層の既存挙動を維持したまま、抽選申込み自動化を追加
- まずは dry-run と候補提示を整え、その後に抽選申込みワークフローへ進む
- 認証情報は Preference に持たせず、ID CSV / `config.local.ini` / `.env` の順に既存設定から解決する
- 空き施設予約は低優先度とし、Phase 3 以降で扱う
- 通知、監視、運用補助は Phase 3 前提で境界を先に整理

## Lottery Automation Minimal Architecture

Phase 2 の主対象は抽選申込み自動化であり、既存の Browser / Service 層をできるだけ再利用する軽量構成を前提とする。

```text
Preference Config
  -> Lottery Candidate Collection
  -> Lottery Candidate Ranking
  -> Lottery Automation Dry-run
  -> Lottery Entry Workflow
  -> Retry / Recovery
```

最小フロー:

1. 設定ファイルから希望条件を読み込む
2. 既存 `AvailabilityService` 由来の候補収集ロジックを流用して候補を収集する
3. `Slot` モデルへ変換する
4. 候補を順位付けする
5. dry-run で候補表示と保存を行う
6. `Lottery Entry Workflow` で既存 `LoginService`、`NavigationService`、`LotteryService` を接続する
7. ランキング結果から最大 2 件、同一日時を除外して候補を選ぶ
8. 抽選申込み画面上で候補だけを自動選択する
9. CAPTCHA / reCAPTCHA が表示された場合は手動認証待機へ入る
10. 認証完了後に既存フローを継続する

Phase 2 では、完璧な抽象化や新しい大型レイヤ追加よりも、既存サービスの組み合わせで早く動かすことを優先する。

Issue 0021 では、この最小構成のうち `Preference Config`、`Lottery Candidate Collection`、`Lottery Candidate Ranking`、`Dry-run Runner` を先に実装し、実際の抽選申込み送信にはまだ進まない。

Issue 0022 では、上記 dry-run 結果を使って抽選申込み画面で候補を自動選択する軽量 workflow を追加する。最終送信は行わず、既存 `LotteryService.auto_select_and_submit_slots(..., submit=False)` を利用して pre-submit までに留める。

## Lottery Guide

- 抽選申込みは利用前月 1 日 0 時から 10 日 23 時 59 分まで
- 抽選結果確認は利用前月 14 日 0 時以降
- 当選した場合は利用前月 20 日 23 時 59 分までに確認および当選施設の利用申込みが必要
- 1 回の抽選につき、種目ごとに 2 件まで
- 空き施設予約は利用前月 22 日から利用開始時刻までだが、Phase 2 では低優先度

## Lottery Automation Exclusions

- 通知機能は Phase 2 では扱わない
- スケジューラ機能は Phase 2 では扱わない
- CAPTCHA / reCAPTCHA の回避、突破、自動認証は実装しない
- GUI 起点の操作フローはこの段階では変更しない
- この段階では実際の抽選申込み送信を行わない
- 将来的に当選結果確認後の確定候補表示は行ってよいが、最終確定はユーザー判断を挟み、完全自動確定は行わない

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
- Browser Session のような共通基盤は先に切り出してよいが、ログイン処理や予約処理の中身は同じ Issue で変更しない
- Login Service のような再利用処理は `browser/` に段階的に切り出してよいが、画面遷移や予約処理の責務までは同時に広げない
- Navigation Service は JavaScript 呼び出しと共通画面遷移の集約先とし、抽選や予約の分岐ロジックは含めない
- Lottery Service は抽選申込み、抽選申込み確認、抽選結果確認の業務フローを集約し、予約確定や空き確認は含めない
- Reservation Service は予約確定、予約確認の業務フローを集約し、抽選処理や空き確認は含めない
- Availability Service は空き確認、空き枠収集の業務フローを集約し、抽選処理や予約確定は含めない
- IdManager Service は ID CSV 読み込み、CSV 書き出し、ID 有効確認を集約し、既存 `Manage_Id` 互換を維持する
- Model 層は `Account`、`Facility`、`Slot`、`ReservationPreference` などの基本概念を保持し、既存サービスへは段階的に適用する

## UI Entrypoint Policy

- 既存の互換起動パス `python court_reserv/court_reserv.py` は維持する
- 新しい UI 層の起動口として `python -m court_reserv.ui.app` を追加する
- `Court_Reserv` クラス本体は当面 `court_reserv/court_reserv.py` に残し、段階的分割を後続 Issue で進める
- 起動口整理の段階では Selenium、Tkinter、予約ロジックの振る舞いを変えない

## Browser Session Policy

- WebDriver の生成、ChromeOptions の設定、`WebDriverWait` の生成、終了処理は `browser/session.py` に集約する
- ChromeDriver の固定パス管理は行わず、Selenium Manager にドライバー解決を委譲する
- `Court_Reserv` では Browser Session を呼び出すだけに留め、ログインや予約の本体ロジックはそのまま維持する
- `find_element_by_*` の置換や Selenium 4 対応は専用 Issue で段階的に行い、業務フローは変えない

## Selenium 4 Migration Policy

- Selenium 4 移行では `find_element(By.*)` 形式への置換を優先し、画面遷移や業務仕様は変更しない
- `WebDriverWait` へ安全に置換できる待機だけを対象とし、意図が不明な `time.sleep()` は無理に変更しない
- `except A or B:` のような不正確な例外処理は tuple 形式へ修正する
- CAPTCHA / reCAPTCHA の手動認証待機方針は維持する

## Login Service Policy

- ログイン処理と CAPTCHA / reCAPTCHA 手動待機方針は `browser/login.py` に集約する
- `Court_Reserv` では Login Service を呼び出すだけに留め、抽選処理や予約処理の本体ロジックは維持する
- CAPTCHA / reCAPTCHA の回避、突破、自動認証は実装しない

## Navigation Service Policy

- JavaScript 実行と共通画面遷移は `browser/navigation.py` に集約する
- `Court_Reserv` では Navigation Service を呼び出すだけに留め、抽選処理・予約処理・空き確認処理の本体ロジックは維持する
- `find_element_by_*` の置換や詳細なページオブジェクト化は別 Issue で扱う

## Lottery Service Policy

- 抽選申込み、抽選申込み状況確認、抽選当選結果確認は `services/lottery.py` に集約する
- `Court_Reserv` では Lottery Service を呼び出すだけに留め、Tkinter UI と他業務フローは維持する
- 予約確定処理、空き確認処理、ID 管理処理は別 Issue で扱う

## Reservation Service Policy

- 予約確定、予約確認は `services/reservation.py` に集約する
- `Court_Reserv` では Reservation Service を呼び出すだけに留め、抽選処理、空き確認処理、ID 管理処理は維持する
- 空き確認処理の分離は別 Issue で扱う

## Availability Service Policy

- 空き確認、空き枠収集は `services/availability.py` に集約する
- `Court_Reserv` では Availability Service を呼び出すだけに留め、抽選処理、予約確定処理、ID 管理処理は維持する
- 自動予約機能や予約戦略エンジンの追加は別 Issue で扱う

## IdManager Service Policy

- ID CSV 読み込み、CSV 書き出し、ID 有効確認は `services/id_manager.py` に整理する
- `manage_id.py` は既存互換ラッパーとして残し、必要に応じて `IdManagerService` へ委譲する
- CSV フォーマットは変更しない

## Model Policy

- `models/` には `Account`、`Facility`、`Slot`、`ReservationPreference` などの軽量 dataclass を追加する
- モデル追加 Issue では既存サービスへの大規模適用は行わず、今後の自動予約や戦略エンジンの土台に留める
- 既存 CSV フォーマットや既存画面値の扱いは変更しない

## Automation Policy

東京都スポーツ施設予約システムの自動化では、以下を共通ルールとします。

- CAPTCHA / reCAPTCHA の回避、突破、自動認証は実装しない
- 外部 CAPTCHA サービスは利用しない
- reCAPTCHA が表示された場合は、ユーザーが手動で認証するまで待機する
- 手動認証完了後は、そのまま既存処理を継続できる設計を維持する
- 今後追加する機能も、この方針を前提に設計する
