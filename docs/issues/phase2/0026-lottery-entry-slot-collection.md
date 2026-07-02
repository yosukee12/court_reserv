# Issue 0026: Lottery Entry Slot Collection

## Status

Done

---

## Summary

抽選申込み画面の日時一覧から、指定曜日の枠だけを取得し、各枠の現在申込数を表示できるようにする。

また、`preferences.yaml` に `default_entries` と `account_overrides` を追加し、IDごとにどの枠へ申し込むかを決められるようにする。

このIssueでは申込み送信のreCAPTCHA復旧は扱わない。送信後reCAPTCHA復旧は次Issueで扱う。

---

## Background

現状の Lottery Automation は、Preference / Ranking / dry-run により候補選定の土台はある。

ただし、実運用では以下が必要。

- 基本的には事前に決めた日時枠へ抽選申込みする
- 1IDあたり最大2枠まで申し込む
- 基本は全ID同じ枠に申し込む
- 必要に応じてIDごとに申込み枠を変えたい
- 抽選申込み画面では、各日時枠の現在申込数を確認できる
- 日時一覧は指定曜日、基本は土曜だけ見たい

---

## Goal

- 抽選申込み画面の日時一覧を取得する
- 指定曜日のみ取得対象にする
- デフォルト指定曜日は土曜とする
- 各枠の現在申込数を取得する
- `default_entries` で全ID共通の申込み枠を指定できる
- `account_overrides` でID別に申込み枠を上書きできる
- 1IDあたり最大2枠までに制限する
- 申込み候補と現在申込数を表示する
- このIssueでは送信後reCAPTCHA復旧は扱わない

---

## Scope

### In Scope

- 抽選申込み日時一覧の取得
- 指定曜日フィルタ
- 現在申込数の取得
- `preferences.yaml` の lottery 設定拡張
- `default_entries`
- `account_overrides`
- 1ID最大2枠チェック
- 申込み候補表示
- JSON保存
- docs更新
- Completion Report更新

### Out of Scope

- 最終送信後のreCAPTCHA復旧
- 送信後の申込み枠プルダウン再選択
- 当選結果確認
- 予約確定補助
- GUI変更
- 通知
- スケジューラ
- CAPTCHA / reCAPTCHA 回避・突破・自動認証

---

## Preference Format

`preferences.yaml` は以下の形式をサポートする。

```yaml
lottery:
  target_weekdays:
    - 土

  max_entries_per_account: 2

  default_entries:
    - facility: 府中の森公園
      date: 2026-08-15
      time_range: "09:00-11:00"

    - facility: 府中の森公園
      date: 2026-08-15
      time_range: "11:00-13:00"

  account_overrides:
    "12345678":
      entries:
        - facility: 府中の森公園
          date: 2026-08-16
          time_range: "09:00-11:00"

        - facility: 府中の森公園
          date: 2026-08-16
          time_range: "11:00-13:00"
```

認証情報は `preferences.yaml` に入れない。

---

## Slot Collection Policy

- 抽選申込み画面から日時一覧を取得する
- `target_weekdays` に含まれる曜日のみ対象にする
- `target_weekdays` 未指定時は `土` をデフォルトにする
- 指定曜日以外は候補にしない
- 各枠の現在申込数を取得する
- 現在申込数は今回は自動判断には使わず、表示情報として扱う
- 将来的に申込数の少ない枠を優先する戦略は別Issueで扱う

---

## Entry Selection Policy

- 基本は `default_entries` を全IDへ適用する
- `account_overrides` があるIDは override 側を優先する
- 1IDあたり最大2件まで
- 同一日時の重複は除外する
- 画面上に存在しない枠は選択対象にしない
- 存在しない枠は警告として結果に保存する

---

## Authentication Policy

認証情報は既存方式を利用する。

優先順位:

1. ID CSV / `IdManagerService`
2. `config.local.ini`
3. `.env`

禁止事項:

- コードへのID/password直書き
- `preferences.example.yaml` へのID/password記載
- README / docsへの実ID・実パスワード記載
- ZIP成果物への `.env` / `config.local.ini` / 実ID CSV 混入

---

## Automation Policy

- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない
- CAPTCHA / reCAPTCHA が表示された場合は手動認証を待つ
- 手動認証後は既存フローを継続する
- 外部CAPTCHAサービスは利用しない

---

## Proposed Files

```text
court_reserv/models/lottery_entry_slot.py
court_reserv/services/lottery_entry_slot_collector.py
```

必要に応じて既存ファイルを最小限更新してよい。

---

## Minimum Workflow

```text
Load Preferences
      ↓
Load Account
      ↓
Login
      ↓
Manual reCAPTCHA if needed
      ↓
Navigate to Lottery Entry
      ↓
Fetch Date/Time Slot List
      ↓
Filter by target_weekdays
      ↓
Read current entry count
      ↓
Apply default_entries / account_overrides
      ↓
Limit to max 2 entries
      ↓
Display candidate slots with current entry count
      ↓
Save JSON
```

---

## Tasks

- [x] LotteryEntrySlot model を追加する
- [x] LotteryEntrySlotCollector を追加する
- [x] 抽選申込み日時一覧を取得する
- [x] 指定曜日フィルタを追加する
- [x] 現在申込数を取得する
- [x] preferences.yaml に lottery.default_entries を追加する
- [x] preferences.yaml に lottery.account_overrides を追加する
- [x] IDごとの申込み候補生成を追加する
- [x] 1ID最大2枠制限を追加する
- [x] 存在しない枠の警告を保存する
- [x] JSON保存を追加する
- [x] README / docs を更新する
- [x] Verification を実施する
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] 指定曜日の枠だけを取得できる
- [x] 未指定時は土曜のみ対象になる
- [x] 各枠の現在申込数を取得できる
- [x] default_entries が全IDに適用される
- [x] account_overrides がID別に適用される
- [x] 1ID最大2件を超えない
- [x] 認証情報を preferences.yaml に追加していない
- [x] 既存GUIを変更していない
- [x] reCAPTCHA方針を変更していない
- [x] python -m compileall . が成功する

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
python scripts/lottery_entry_workflow.py --help
```

可能なら:

```bash
python scripts/lottery_entry_workflow.py --preferences config/preferences.example.yaml
```

実サイト確認が必要な部分は Completion Report に未確認として記載する。

---

## Notes

- このIssueでは送信後reCAPTCHA復旧は扱わない。

---

# Completion Report

## Summary

- `LotteryEntrySlot` model と `LotteryEntrySlotCollector` を追加し、抽選申込み画面の現在表示中一覧から指定曜日の枠と現在申込数を取得できるようにした
- `preferences.example.yaml` を `lottery.default_entries` / `lottery.account_overrides` 形式へ更新し、全ID共通設定とID別上書きを扱えるようにした
- `LotteryEntryWorkflowService` を更新し、曜日フィルタ後の画面枠に対して ID ごとの申込み予定を作成し、1ID 最大2枠までに制限し、存在しない枠は警告として JSON に保存するようにした
- 最終送信の `yes` 確認仕様は維持し、送信後 reCAPTCHA 復旧は追加していない
- 追加修正として `NavigationService.select_lottery_tennis_park()` の施設選択待機を安定化し、`bname` / `iname` の状態、URL、タイトルを診断できるようにした
- 施設選択失敗時は `output/debug_pages/` に HTML を保存し、取得済み option 一覧を `TimeoutException` メッセージへ含めるようにした
- さらに `iname` 選択後の `selected value` / `selected text` / `onchange` / URL / title を確認し、`changeIname(document.form1, gLotWTransLotInstSrchVacantAjaxAction)` を明示実行したうえで日時一覧ロード完了まで待つようにした
- 施設選択完了後の HTML を `output/debug_pages/lottery_entry_after_iname_selection.html`、DOM summary を `output/debug_pages/lottery_entry_after_iname_selection_dom_summary.json` へ保存するようにした
- `legacy/bk_court_reserv.py` の週送り処理を参考に、抽選申込み日時一覧を最大 8 週まで巡回しながら `default_entries` / `account_overrides` の対象日付を探索できるようにした
- 週ごとの取得件数と対象日付のマッチ状況を `week_search` として result JSON に保存し、週送りごとの HTML / DOM summary も `output/debug_pages/lottery_entry_week_*.html` / `*_dom_summary.json` として保存するようにした
- `900` のような 3 桁時刻を `09:00` へ正規化し、`preferences.example.yaml` の `09:00-11:00` 形式と正しくマッチするようにした

## Changed Files

- `court_reserv/models/lottery_entry_slot.py`
- `court_reserv/models/preference.py`
- `court_reserv/models/__init__.py`
- `court_reserv/config/preferences.py`
- `court_reserv/services/lottery_entry_slot_collector.py`
- `court_reserv/services/lottery_entry_workflow.py`
- `court_reserv/services/__init__.py`
- `scripts/lottery_entry_workflow.py`
- `scripts/lottery_result_workflow.py`
- `scripts/reservation_confirmation_workflow.py`
- `config/preferences.example.yaml`
- `README.md`
- `docs/ROADMAP.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase2/0026-lottery-entry-slot-collection.md`

## Verification Result

- `git status`
  変更ファイルを確認した
- `python -m compileall .`
  成功
- `python setup.py --name`
  `court_reserv`
- `python scripts/lottery_entry_workflow.py --help`
  成功
- `python scripts/lottery_entry_workflow.py --preferences config/preferences.example.yaml`
  再実行した。`.env` 解決の認証情報でログインし、`NavigationService.select_lottery_tennis_park()` の `TimeoutException` は解消した。workflow は `status=completed` で `output/lottery_automation/lottery_entry_workflow_result.json` を保存して終了した
  追加修正後の再実行では `iname=12700020` / `selected_text=テニス（人工芝）` と `bname=1301270` / `selected_text=府中の森公園` が DOM summary 上で確認でき、`LotteryEntrySlotCollector` は抽選申込み日時一覧から 6 件の土曜枠を取得できた
  ただし `preferences.example.yaml` の `default_entries` は `2026-08-15` を指しており、今回実取得できた一覧は `2026-08-01` の6枠だったため、指定2枠は `slot not found on the current lottery entry page` として警告保存された
- `python scripts/lottery_entry_workflow.py --preferences config/preferences.example.yaml`
  週送り対応後の再実行では、5週分を探索して `2026-08-01 / 08 / 15 / 22 / 29` の各週から土曜 6 枠ずつ、合計 30 枠を取得できた
  `preferences.example.yaml` の `default_entries` に指定した `2026-08-15 09:00-11:00` と `2026-08-15 11:00-13:00` を発見し、`planned_slots` として 2 件マッチできた
  申込み送信は行わず、`selection_result` では対象 2 枠が選択済みとなった
  実行途中でサイト側の「データ通信を正しく行うことができませんでした」アラートに一度遭遇したが、再試行では正常終了した

## Unverified

- 実サイト上での抽選申込み日時一覧 DOM に対する `LotteryEntrySlotCollector` の精度
- `target_weekdays` フィルタ後の画面枠と `default_entries` / `account_overrides` の実マッチング
- 実ページ上で存在しない枠警告と最終送信前選択の組み合わせ確認
- `select_lottery_tennis_park()` と週送り探索は成功したが、サイト側の一時的な通信エラーアラートが発生する場合がある

## Follow-up Items

- 送信後 reCAPTCHA 復旧
- Retry / Recovery
- 現在申込数を使った戦略最適化
- 送信後reCAPTCHA復旧は次Issueで扱う。
- 現在申込数は今回は表示情報として扱う。
- 通知・スケジューラは実装しない。
- サイト側通信エラー発生時の再試行 / Recovery
- `account_overrides` を含む複数IDでの週送り探索実確認
