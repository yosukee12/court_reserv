# Issue 0024: Lottery Result Workflow

## Status

Done

---

## Summary

抽選結果確認フローを自動化し、当選・落選結果を一覧表示および JSON / CSV に保存できるようにする。

この Issue では当選後の予約確定は行わない。  
予約確定は次 Issue の Reservation Confirmation Assist で扱う。

---

## Background

Phase 2 の主対象は抽選申込み自動化である。

Issue 0021〜0023 では、抽選申込み候補の選定、自動選択、ユーザー確認後の送信までを整備した。

次に、抽選結果確認を自動化し、当選・落選の状況を把握できるようにする。

東京都スポーツ施設予約システムのガイドでは、抽選結果確認は利用前月14日0時以降、当選した場合は利用前月20日23時59分までに確認および当選施設の利用申込みが必要である。

---

## Goal

- 既存認証情報でログインする
- reCAPTCHA が表示された場合は手動待機する
- 抽選結果確認画面へ遷移する
- 抽選結果を取得する
- 当選・落選を一覧表示する
- 結果を JSON / CSV に保存する
- 当選後の予約確定は行わない

---

## Scope

### In Scope

- Lottery Result Workflow の追加
- CLIスクリプト追加
- `LotteryService` / `LoginService` / `NavigationService` の活用
- 抽選結果一覧の取得
- 当選・落選の分類
- 結果の標準出力
- 結果の JSON / CSV 保存
- README / docs 更新
- Completion Report 更新

### Out of Scope

- 当選後の予約確定
- 予約確定ボタンの押下
- GUI変更
- 空き施設予約
- 通知
- スケジューラ
- CAPTCHA / reCAPTCHA 回避・突破・自動認証

---

## Authentication Policy

認証情報は `preferences.yaml` に入れない。

認証情報は既存の仕組みを利用する。

優先順位:

1. ID CSV / `IdManagerService`
2. `config.local.ini`
3. `.env`

禁止事項:

- コードへの ID / password 直書き
- `preferences.example.yaml` への ID / password 記載
- README / docs への実 ID・実パスワード記載
- ZIP 成果物への `.env` / `config.local.ini` / 実 ID CSV 混入

---

## Result Policy

- 抽選結果は一覧表示する
- 当選・落選を分類する
- 当選分は次 Issue の Reservation Confirmation Assist で利用できる形式で保存する
- この Issue では予約確定しない
- 勝手に確定しない

---

## Automation Policy

- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない
- CAPTCHA / reCAPTCHA が表示された場合は手動認証を待つ
- 手動認証後は既存フローを継続する
- 外部 CAPTCHA サービスは利用しない

---

## Proposed Files

```text
court_reserv/services/lottery_result_workflow.py
scripts/lottery_result_workflow.py
```

必要に応じて既存ファイルを最小限更新してよい。

---

## Minimum Workflow

```text
Load Account
        ↓
Login
        ↓
Manual reCAPTCHA if needed
        ↓
Navigate to Lottery Result
        ↓
Fetch Result Rows
        ↓
Classify Won / Lost / Unknown
        ↓
Print Summary
        ↓
Save JSON / CSV
        ↓
Do not confirm reservation
```

---

## Output

保存先:

```text
output/lottery_automation/
```

例:

```text
lottery_result_workflow_result.json
lottery_result_workflow_result.csv
```

JSON には最低限以下を含める。

```json
{
  "status": "completed",
  "results": [
    {
      "account": "masked-or-user-id",
      "date": "YYYY-MM-DD",
      "time_range": "09:00-11:00",
      "facility": "府中の森公園 テニス",
      "result": "won"
    }
  ]
}
```

---

## Implementation Plan

1. `LotteryResultWorkflowService` を追加する
2. `scripts/lottery_result_workflow.py` を追加する
3. 既存認証情報からログインする
4. 抽選結果確認画面へ遷移する
5. 既存 `LotteryService` の結果確認処理を活用・拡張する
6. 抽選結果を構造化する
7. 当選・落選を分類する
8. 結果を表示する
9. JSON / CSV に保存する
10. README / docs を更新する
11. Verification を実施する
12. Completion Report を記入する

---

## Tasks

- [x] Lottery Result Workflow を追加する
- [x] CLIスクリプトを追加する
- [x] 既存認証情報でログインする
- [x] 抽選結果確認画面へ遷移する
- [x] 結果行を取得する
- [x] 当選・落選を分類する
- [x] 結果を表示する
- [x] JSON / CSV に保存する
- [x] README / docs を更新する
- [x] Verification を実施する
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] 抽選結果確認 workflow が追加されている
- [x] CLI から実行できる
- [x] 結果を一覧表示できる
- [x] 当選・落選を分類できる
- [x] JSON / CSV に保存できる
- [x] この Issue では予約確定しない
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
python scripts/lottery_result_workflow.py --help
```

可能なら:

```bash
python scripts/lottery_result_workflow.py
```

実サイト確認が必要な部分は Completion Report に未確認として記載する。

---

## Notes

- この Issue では予約確定しない。
- 当選結果の確認と保存までを対象とする。
- 当選後の予約確定補助は次 Issue で扱う。
- 通知・スケジューラは実装しない。

---

# Completion Report

※ Codex が記入

## Summary

- `LotteryResultWorkflowService` と `scripts/lottery_result_workflow.py` を追加し、既存認証情報でログインして抽選結果確認画面の結果行を取得・分類できるようにした
- 認証情報は Issue 方針どおり `IdManagerService` の ID CSV、`config.local.ini`、`.env` の順に解決し、`preferences.yaml` には追加していない
- 抽選結果行は当選 / 落選 / 不明に分類し、標準出力に一覧表示するとともに `output/lottery_automation/` に JSON / CSV 保存するようにした
- この Issue では予約確定ボタン押下や当選後の確定処理は追加していない

## Changed Files

- `court_reserv/services/__init__.py`
- `court_reserv/services/lottery_result_workflow.py`
- `scripts/lottery_result_workflow.py`
- `README.md`
- `docs/ROADMAP.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase2/0024-lottery-result-workflow.md`

## Verification Result

- `git status`
  変更ファイルを確認した
- `python -m compileall .`
  成功
- `python setup.py --name`
  `court_reserv`
- `python scripts/lottery_result_workflow.py --help`
  成功
- `python scripts/lottery_result_workflow.py`
  実行した。Selenium 起動後、`.env` 解決された認証情報でログインを試みたが `利用者番号、またはパスワードが誤っています。再度入力して下さい。` の alert が表示され、結果行 0 件として `output/lottery_automation/lottery_result_workflow_result.json` / `.csv` を保存して終了した

## Unverified

- 実サイト上での抽選結果画面の行構造に対する parser の実確認
- reCAPTCHA 表示時の手動待機を挟んだ結果確認継続
- 実アカウントの当選 / 落選 / 不明分類の精度確認

## Follow-up Items

- Reservation Confirmation Assist
- Retry / Recovery
- 結果画面 HTML に合わせた parser 精度調整

## Changed Files

---

## Verification Result

---

## Follow-up Items

- Reservation Confirmation Assist
- Retry / Recovery
