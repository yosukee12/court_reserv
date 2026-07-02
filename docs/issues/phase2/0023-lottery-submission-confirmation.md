# Issue 0023: Lottery Submission Confirmation

## Status

Done

---

## Summary

Issue 0022 で実装した抽選申込み候補の自動選択後に、ユーザー確認を挟んで、明示的に `yes` と入力された場合のみ最終送信できるようにする。

この Issue では、抽選申込みの最終送信を完全自動化しない。必ずユーザー判断を挟む。

---

## Background

Phase 2 の主対象は、空き施設予約ではなく抽選申込み自動化である。

Issue 0021 では Preference / Ranking / dry-run の土台を作成した。Issue 0022 では抽選申込み画面で候補を自動選択する Workflow を追加した。

次の段階として、候補選択後に内容を表示し、ユーザー確認後にのみ最終送信できるようにする。

---

## Goal

- 抽選申込み候補の選択内容を表示する
- 最終送信前に必ずユーザー確認を挟む
- ユーザーが `yes` と入力した場合のみ送信する
- `no` / 空入力 / その他入力の場合は送信しない
- 送信結果を保存する
- reCAPTCHA 手動待機方針を維持する

---

## Scope

### In Scope

- `LotteryEntryWorkflowService` への確認付き送信処理追加
- `scripts/lottery_entry_workflow.py` から確認付き送信まで実行可能にする
- 候補選択後の内容表示
- 対話式確認プロンプト
- `yes` 入力時のみ最終送信
- 送信結果の JSON 保存
- README / docs 更新
- Completion Report 更新

### Out of Scope

- 完全自動送信フラグ追加
- `--yes` / `--live` フラグ追加
- GUI ボタン追加
- 抽選結果確認
- 当選後の予約確定
- 空き施設予約
- 通知
- スケジューラ
- CAPTCHA / reCAPTCHA 回避・突破・自動認証

---

## Authentication Policy

認証情報は `preferences.yaml` に入れない。

認証情報は既存の仕組みを利用する。

1. ID CSV (`IdManagerService`)
2. `config.local.ini`
3. `.env`

禁止事項:

- コードへの ID / Password の直書き
- `preferences.example.yaml` への認証情報記載
- README / docs への実 ID・実パスワード記載
- ZIP 成果物への `.env` / `config.local.ini` / 実 ID CSV 混入

---

## Lottery Submission Policy

- デフォルトでは送信しない
- 最終送信前に必ず候補内容を表示する
- ユーザーが `yes` と入力した場合のみ送信する
- `no` / 空入力 / その他入力の場合は送信しない
- この Issue では `--yes` / `--live` のような完全自動送信フラグは追加しない
- 送信結果を `output/lottery_automation/` に保存する

---

## Tasks

- [x] 確認付き送信処理を追加する
- [x] 候補選択内容を表示する
- [x] 対話式確認プロンプトを追加する
- [x] `yes` 入力時のみ送信する
- [x] `no` / 空入力 / その他入力では送信しない
- [x] 送信結果を JSON 保存する
- [x] README / docs を更新する
- [x] Verification を実施する
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] 候補選択内容が表示される
- [x] 最終送信前に確認プロンプトが表示される
- [x] `yes` 入力時のみ送信される
- [x] `no` / 空入力 / その他入力では送信されない
- [x] `--yes` / `--live` フラグが追加されていない
- [x] 認証情報を `preferences.yaml` に追加していない
- [x] 既存 GUI を変更していない
- [x] 空き施設予約を対象にしていない
- [x] CAPTCHA / reCAPTCHA 方針を変更していない
- [x] `python -m compileall .` が成功する

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
python scripts/lottery_entry_workflow.py --help
python scripts/lottery_entry_workflow.py --preferences config/preferences.example.yaml
```

実サイト確認が必要な部分は Completion Report に未確認として記載する。

---

# Completion Report

## Summary

- `LotteryEntryWorkflowService` に対話式確認付き送信処理を追加し、候補選択後に内容を表示したうえで `yes` 入力時のみ送信するようにした
- `LotteryService` の既存送信部分を `submit_selected_slots()` として再利用できる形に整理し、Issue 0022 の選択 workflow から同一セッション内で呼び出せるようにした
- `scripts/lottery_entry_workflow.py` は `--yes` / `--live` を追加せず、対話式プロンプト経由でのみ送信する構成を維持した
- 送信の有無、確認応答、送信結果を `output/lottery_automation/lottery_entry_workflow_result.json` に保存するようにした

## Changed Files

- `court_reserv/services/lottery.py`
- `court_reserv/services/lottery_entry_workflow.py`
- `scripts/lottery_entry_workflow.py`
- `README.md`
- `docs/ROADMAP.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase2/0023-lottery-submission-confirmation.md`

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
  成功。候補 0 件のため確認プロンプトや送信には進まず、安全に結果 JSON を保存して終了した

## Unverified

- 実サイト上で候補が存在する状態での確認プロンプト後送信は未確認
- reCAPTCHA 表示時の手動待機を挟んだ送信完了は未確認

## Follow-up Items

- Lottery Result Workflow
- Reservation Confirmation Assist
- Retry / Recovery
