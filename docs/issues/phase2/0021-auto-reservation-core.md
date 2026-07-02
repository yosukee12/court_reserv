# Issue 0021: Lottery Automation Core

## Status

Done

---

## Summary

抽選申込み自動化の中核として、希望条件の読み込み、候補収集、候補スコアリング、dry-run実行までをまとめて実装する。

このIssueでは実際の抽選申込み送信は行わず、まず「希望条件に合う抽選申込み候補を自動抽出・順位付けできる」状態を作る。

---

## Background

Phase 2 では、完璧な抽象化よりも既存サービスを活用して早く動くことを優先する。主対象は空き施設予約ではなく抽選申込み自動化である。

当初は Preference Config / Slot Collection / Slot Ranking を別Issueに分ける予定だったが、スピード重視のため本Issueでまとめて実装する。

---

## Goal

- 希望条件を設定ファイルから読み込める
- 既存 `AvailabilityService` を活用して候補元データを取得できる
- 候補元データを `Slot` モデルへマッピングできる
- 希望条件に基づき候補をスコアリングできる
- dry-runで抽選申込み候補順位を表示・保存できる
- 実際の抽選申込み送信はまだ行わない

---

## Scope

### In Scope

- `lottery_automation` 用 service / runner の追加
- Preference Config 読み込み
- `ReservationPreference` モデルの活用
- Slot Collection Adapter
- Slot Ranking
- Dry-run Runner
- 結果の標準出力
- 必要に応じた JSON / CSV 出力
- サンプル設定ファイル追加
- README / docs 更新
- Completion Report 更新

### Out of Scope

- 実際の抽選申込み送信
- GUIボタン追加
- live実行
- リトライ / リカバリ
- スケジューラ
- 通知
- CAPTCHA / reCAPTCHA 回避・突破・自動認証

---

## Proposed Files

```text
court_reserv/services/lottery_automation.py
court_reserv/services/slot_ranking.py
court_reserv/config/preferences.py
config/preferences.example.yaml
scripts/lottery_automation_dry_run.py
```

ファイル名や構成は既存コードとの相性を優先して調整してよい。

---

## Minimum Flow

```text
1. Load preferences
2. Collect lottery candidates
3. Convert raw slots to Slot models
4. Score slots
5. Sort candidates
6. Print result
7. Save result
```

---

## Preference Example

```yaml
preferred_facilities:
  - park_name: 府中の森公園
    facility_name: テニス（人工芝）

preferred_weekdays:
  - 土
  - 日

preferred_time_ranges:
  - "09:00-11:00"
  - "11:00-13:00"

max_candidates: 10
```

---

## Ranking Policy

シンプルな加点方式でよい。

例：

```text
facility match: +50
weekday match: +30
time range match: +20
earlier date: +small bonus
```

完璧なロジックより、まず候補順位が出ることを優先する。

---

## Safety Policy

- dry-run をデフォルトにする
- このIssueでは抽選申込み送信しない
- live実行フラグは追加しない
- CAPTCHA / reCAPTCHA 回避は実装しない

## Lottery Guide

- 抽選申込みは利用前月 1 日 0 時から 10 日 23 時 59 分まで
- 抽選結果確認は利用前月 14 日 0 時以降
- 当選した場合は利用前月 20 日 23 時 59 分までに確認および当選施設の利用申込みが必要
- 1 回の抽選につき、種目ごとに 2 件まで
- 空き施設予約は利用前月 22 日から利用開始時刻までだが、Phase 2 では低優先度

## Reservation Confirmation Policy

- 将来的には当選結果確認後、確定候補を表示する
- 最終確定はユーザー判断を挟む
- 勝手に完全自動で確定しない

---

## Tasks

- [ ] Preference Config 読み込みを追加
- [ ] サンプル設定ファイルを追加
- [ ] Slot Collection Adapter を追加
- [ ] Slot Ranking を追加
- [ ] Lottery Automation Core を追加
- [ ] dry-runスクリプトを追加
- [ ] 結果出力を追加
- [ ] docs更新
- [ ] Verification実施
- [ ] Completion Report記入

---

## Acceptance Criteria

- [ ] 希望条件ファイルを読み込める
- [ ] dry-runを実行できる
- [ ] 候補が順位付きで出力される
- [ ] 実際の抽選申込み送信は行わない
- [ ] 既存GUIを壊していない
- [ ] 既存の抽選・予約・空き確認機能を壊していない
- [ ] reCAPTCHA 方針を変更していない
- [ ] `python -m compileall .` が成功する

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
python scripts/lottery_automation_dry_run.py --help
```

可能なら：

```bash
python scripts/lottery_automation_dry_run.py --preferences config/preferences.example.yaml --dry-run
```

実サイトアクセスが必要な場合は、Completion Report に未実施理由を記載する。

---

## Notes

- このIssueは Phase 2 の中核実装である。
- 実際の抽選申込み送信は次Issue以降で行う。
- 動くことを優先し、細かい抽象化は後回しでよい。
- 通知・スケジューラは作らない。
- 空き施設予約は低優先度とする。

---

# Completion Report

※ Codex が記入

## Summary

- 希望条件を YAML / JSON から読み込む `Preference Config` を追加した
- `AvailabilityService` 互換の `Slot Collection Adapter` を追加し、候補元 CSV や既存 service 出力を `Slot` モデルへ変換できるようにした
- シンプルな加点方式の `Slot Ranking` を追加し、dry-run で抽選申込み候補順位を表示・保存できるようにした
- `scripts/lottery_automation_dry_run.py` を追加し、実際の抽選申込み送信を行わない dry-run をデフォルトで実行できるようにした
- 通知、スケジューラ、live 実行、実際の抽選申込み送信はこの Issue の対象外として維持した

## Changed Files

- `court_reserv/models/preference.py`
- `court_reserv/models/slot.py`
- `court_reserv/config/__init__.py`
- `court_reserv/config/preferences.py`
- `court_reserv/services/__init__.py`
- `court_reserv/services/lottery_automation.py`
- `court_reserv/services/auto_reservation.py`
- `court_reserv/services/slot_ranking.py`
- `config/preferences.example.yaml`
- `scripts/lottery_automation_dry_run.py`
- `scripts/auto_reservation_dry_run.py`
- `README.md`
- `docs/ROADMAP.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase2/0021-auto-reservation-core.md`

## Verification Result

- `git status`
  新規モジュール、サンプル設定、dry-run スクリプト、関連ドキュメントの変更を確認した
- `python -m compileall .`
  成功
- `python setup.py --name`
  `court_reserv`
- `python scripts/lottery_automation_dry_run.py --help`
  成功
- `python scripts/lottery_automation_dry_run.py --preferences config/preferences.example.yaml --dry-run`
  成功。既存の `court_reserv/debug_pages/available_slots_2026-06-10.csv` を候補元として読み込み、候補 0 件として安全に dry-run 完了

## Follow-up Items
- `Lottery Entry Workflow` を既存 `LotteryService` 上にどう接続するか整理する
- dry-run の順位結果を使った抽選申込み自動化の入口を定義する
- Retry / Recovery をどの失敗種別まで扱うか明確にする
