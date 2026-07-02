# Issue 0025: Reservation Confirmation Assist

## Status

Done

---

## Summary

抽選当選一覧から予約確定候補を表示し、ユーザーが確定対象を選択した後、最終確認を経て予約確定を実行する。

本Issueでは「人が判断する」ことを前提とし、当選した予約を勝手に確定しない。

---

## Background

Issue 0024 で抽選結果取得が実装された。

次に、当選一覧から予約確定対象を選択し、既存 ReservationService を利用して予約確定できるようにする。

---

## Goal

- 当選一覧を表示する
- 確定対象をユーザーが選択する
- 最終確認を表示する
- yes の場合のみ予約確定する
- no は何もしない
- 結果を JSON 保存する

---

## Scope

### In Scope

- Reservation Confirmation Workflow
- 当選一覧表示
- 確定対象選択
- ReservationService利用
- 最終確認
- JSON保存
- docs更新

### Out of Scope

- 抽選申込み
- 抽選結果取得
- GUI変更
- 通知
- スケジューラ
- CAPTCHA / reCAPTCHA 回避

---

## Authentication Policy

認証情報は既存方式を利用する。

優先順位

1. IdManagerService
2. config.local.ini
3. .env

---

## Reservation Policy

- 当選した予約のみ表示する
- ユーザーが確定対象を選択する
- yes の場合のみ予約確定する
- no の場合は終了する
- 勝手に全件確定しない
- 確定結果を JSON 保存する

---

## Automation Policy

- CAPTCHA / reCAPTCHA は手動認証
- 回避・突破・自動認証は実装しない

---

## Minimum Workflow

```text
Load Accounts
      ↓
Login
      ↓
Manual reCAPTCHA
      ↓
Lottery Result
      ↓
Show Won Entries
      ↓
Select Reservation
      ↓
Ask Confirmation
      ↓
yes
      ↓
ReservationService.confirm()
      ↓
Save Result
```

---

## Proposed Files

```text
court_reserv/services/reservation_confirmation_workflow.py
scripts/reservation_confirmation_workflow.py
```

---

## Tasks

- [x] Reservation Confirmation Workflow追加
- [x] 当選一覧表示
- [x] 対象選択
- [x] ReservationService利用
- [x] 最終確認
- [x] JSON保存
- [x] docs更新
- [x] Verification
- [x] Completion Report

---

## Acceptance Criteria

- [x] 当選一覧が表示される
- [x] 確定対象を選択できる
- [x] yes の場合のみ予約確定される
- [x] no の場合は何もしない
- [x] compileall成功

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
python scripts/reservation_confirmation_workflow.py --help
```

---

# Completion Report

※ Codex が記入

## Summary

- `ReservationConfirmationWorkflowService` と `scripts/reservation_confirmation_workflow.py` を追加し、当選一覧を表示してユーザーが対象を選び、`yes` 入力時のみ予約確定できるようにした
- 既存 `LotteryResultWorkflowService` の当選一覧を利用し、認証情報は Issue 方針どおり `IdManagerService` / `config.local.ini` / `.env` の優先順位で解決する構成を維持した
- 既存 `ReservationService` にアカウント辞書入力の `confirm_accounts()` を追加し、手動 reCAPTCHA 待機を含む `LoginService` と組み合わせて予約確定できるようにした
- 既存 `ReservationService` の制約に合わせ、同一アカウントの当選はアカウント単位で確定し、一部だけを選ぶ指定は安全のため中断するようにした

## Changed Files

- `court_reserv/services/reservation.py`
- `court_reserv/services/reservation_confirmation_workflow.py`
- `court_reserv/services/__init__.py`
- `scripts/reservation_confirmation_workflow.py`
- `README.md`
- `docs/ROADMAP.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase2/0025-reservation-confirmation-assist.md`

## Verification Result

- `git status`
  変更ファイルを確認した
- `python -m compileall .`
  成功
- `python setup.py --name`
  `court_reserv`
- `python scripts/reservation_confirmation_workflow.py --help`
  成功
- `python scripts/reservation_confirmation_workflow.py`
  実行した。`.env` 解決の認証情報で結果確認 workflow を呼び出し、当選 0 件のため `No won entries found.` と表示して予約確定には進まず、`output/lottery_automation/reservation_confirmation_workflow_result.json` を保存して終了した

## Unverified

- 当選一覧が存在する状態での対象選択
- `yes` 入力後の実予約確定完了
- reCAPTCHA 表示時の手動待機を挟んだ予約確定継続

## Follow-up Items

- 0026 Retry / Recovery
- 結果画面と予約確定画面の HTML に基づく粒度改善
- アカウント単位ではなく当選行単位で確定できるかの調査

## Changed Files

---

## Verification Result

---

## Follow-up Items

0026 Retry / Recovery
