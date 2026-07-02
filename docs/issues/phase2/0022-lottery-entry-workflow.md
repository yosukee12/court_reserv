# Issue 0022: Lottery Entry Workflow

## Status

Done

---

## Summary

Phase 2 の抽選申込み自動化として、dry-run の順位結果を使って抽選申込み画面で候補を自動選択する workflow を追加する。

この Issue では最終送信は行わず、既存 `LoginService`、`NavigationService`、`LotteryService` を利用して pre-submit までを自動化する。

---

## Background

Issue 0021 で希望条件、候補収集、順位付け、dry-run の土台は整った。次の段階として、順位結果をそのまま抽選申込み画面の選択へ接続し、既存サービスを再利用した軽量 workflow を用意する。

Phase 2 の主対象は空き施設予約ではなく抽選申込み自動化であり、空き施設予約は低優先度とする。

---

## Goal

- CLI から抽選申込み候補選択 workflow を実行できる
- 希望条件とランキング結果を使って最大 2 件の候補を選べる
- 同一日時の候補を除外できる
- 抽選申込み画面上で候補を自動選択できる
- 最終送信は行わない

---

## Scope

### In Scope

- `lottery_entry_workflow` service の追加
- CLI スクリプトの追加
- `LoginService` / `NavigationService` / `LotteryService` の再利用
- 既存 Preference / Ranking 結果の利用
- 認証情報の優先順位制御
- README / docs 更新
- Completion Report 更新

### Out of Scope

- 抽選申込みの最終送信
- GUI ボタン追加
- Preference への認証情報追加
- CAPTCHA / reCAPTCHA の回避、突破、自動認証
- 通知
- スケジューラ
- 空き施設予約機能の拡張

---

## Authentication Policy

認証情報は既存の仕組みを次の優先順位で利用する。

1. `IdManagerService` から読む ID CSV
2. `config.local.ini`
3. `.env`

Preference に ID / password を持たせない。

---

## Tasks

- [x] `court_reserv/services/lottery_entry_workflow.py` を追加
- [x] `scripts/lottery_entry_workflow.py` を追加
- [x] ランキング結果から最大 2 件を選ぶ処理を追加
- [x] 同一日時除外を追加
- [x] 抽選申込み画面の候補自動選択を追加
- [x] 最終送信を行わない構成を維持
- [x] README / docs 更新
- [x] Verification 実施
- [x] Completion Report 記入

---

## Acceptance Criteria

- [x] `python scripts/lottery_entry_workflow.py --help` が動く
- [x] CLI から workflow を起動できる
- [x] 既存 Ranking 結果を利用できる
- [x] 最大 2 件までしか選ばない
- [x] 同一日時の候補を除外する
- [x] 最終送信を行わない
- [x] 既存 GUI を壊していない
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

実サイトアクセスが必要な確認は Completion Report に未確認として記載する。

---

# Completion Report

## Summary

- `LotteryEntryWorkflowService` を追加し、既存 Preference / Ranking 結果から抽選申込み候補を最大 2 件まで選ぶ workflow を実装した
- 認証情報は Issue 指定どおり `IdManagerService` の ID CSV、`config.local.ini`、`.env` の優先順位で解決するようにした
- `LotteryService.auto_select_and_submit_slots(..., submit=False)` を利用し、抽選申込み画面で候補を自動選択するところまでを CLI から実行できるようにした
- 最終送信は行わず、候補表示と選択結果表示、および JSON サマリ保存までに留めた

## Changed Files

- `court_reserv/services/__init__.py`
- `court_reserv/services/lottery_entry_workflow.py`
- `scripts/lottery_entry_workflow.py`
- `config/preferences.example.yaml`
- `README.md`
- `docs/ROADMAP.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/README.md`
- `docs/issues/phase2/0022-lottery-entry-workflow.md`

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
  成功。`court_reserv/debug_pages/available_slots_2026-06-10.csv` を候補元として読み込み、候補 0 件のためブラウザ起動や実サイト遷移なしで安全に終了し、`output/lottery_automation/lottery_entry_workflow_result.json` を保存した

## Unverified

- 抽選申込み画面での実候補選択は、実サイト上で候補が存在し、かつ認証情報が必要なため未確認

## Follow-up Items

- 抽選申込み画面で選択した候補の確認 UI / 出力内容を次 Issue でさらに整理する
- Retry / Recovery の対象失敗パターンを定義する
- 最終送信前の確認フローを別 Issue で定義する
