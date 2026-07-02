# Issue 0020: Lottery Automation Architecture

## Status

Done

---

## Summary

Phase 2 の抽選申込み自動化に向けて、最小限のアーキテクチャと実装方針を定義する。

このIssueでは設計・ドキュメント整理を中心に行い、実装は最小限に留める。  
今後は「完璧な設計」よりも「動くこと」を優先し、段階的に改善する。

---

## Background

Phase 1 では以下が完了した。

- Browser layer separation
- Service layer separation
- Model foundation
- Selenium 4 migration
- Legacy cleanup
- Documentation and issue workflow

Phase 2 では、既存の空き確認・抽選関連処理を活用し、抽選申込み自動化を追加する。空き施設予約は低優先度とする。

---

## Goal

- 抽選申込み自動化の最小構成を決める
- 既存 service 層を活用する方針を決める
- Phase 2 の実装順を決める
- reCAPTCHA 手動待機方針を維持する
- スピード重視で動く実装へ進む準備をする

---

## Scope

### In Scope

- Lottery Automation の全体フロー定義
- 最小構成の責務定義
- Phase 2 Issue 構成の整理
- docs/ROADMAP.md 更新
- docs/ARCHITECTURE.md 更新
- docs/DEVELOPMENT.md 更新
- Completion Report 更新

### Out of Scope

- 実際の予約実行機能の追加
- GUIボタン追加
- 予約戦略エンジンの実装
- Slot Ranking の実装
- Selenium操作の大幅変更
- CAPTCHA / reCAPTCHA 回避・突破・自動認証

---

## Implementation Direction

Phase 2 は以下の順番で進める。

```text
0020 Lottery Automation Architecture
0021 Preference Config
0022 Slot Collection Adapter
0023 Slot Ranking
0024 Lottery Entry Workflow
0025 Lottery Automation Dry-run Expansion
0026 Retry / Recovery
```

### Minimum Lottery Automation Flow

```text
1. Load preferences
2. Login
3. Collect lottery candidates
4. Rank candidates
5. Select best candidates
6. Prepare lottery entry targets
7. If CAPTCHA / reCAPTCHA appears:
     wait for manual verification
8. Continue the existing flow
9. Save result
```

---

## Speed Priority Policy

Phase 2 では以下を優先する。

- 完璧な抽象化より、まず動くこと
- 既存 service をできるだけ再利用する
- 既存GUIを壊さない
- dry-run を先に用意する
- 実際の抽選申込み送信は後続 Issue に分ける
- reCAPTCHA は手動対応でよい
- 通知・スケジューラは作らない

---

## Automation Policy

- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない
- CAPTCHA / reCAPTCHA が表示された場合は手動認証を待つ
- 手動認証後は既存フローを継続する
- 外部CAPTCHAサービスは利用しない

---

## Proposed Components

### Preference Config

希望条件を設定ファイルで管理する。

例：

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

max_reservations: 1
```

### Slot Collector

既存 `AvailabilityService` を活用して抽選候補元データを取得する。

### Slot Ranking

希望条件に合う抽選申込み候補をスコアリングする。

### Lottery Entry Workflow

既存 `LotteryService` / `NavigationService` / `LoginService` を組み合わせて抽選申込みを実行する。

### Runner

最終的に抽選申込み自動化のオーケストレーションを行う。

---

## Target Files

- docs/ROADMAP.md
- docs/ARCHITECTURE.md
- docs/DEVELOPMENT.md
- docs/issues/phase2/0020-auto-reservation-architecture.md

---

## Tasks

- [ ] Phase 2 の方針を整理する
- [ ] Lottery Automation 最小フローを記載する
- [ ] Speed Priority Policy を記載する
- [ ] Automation Policy を再確認する
- [ ] Proposed Components を記載する
- [ ] docs を更新する
- [ ] Completion Report を記入する

---

## Acceptance Criteria

- [ ] Phase 2 の実装方針が明確になっている
- [ ] 通知・スケジューラを作らない方針が明記されている
- [ ] reCAPTCHA 手動対応方針が維持されている
- [ ] 0021以降の実装Issueへ進める状態になっている
- [ ] ソースコードの仕様変更をしていない

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
```

---

## Notes

このIssueは Phase 2 の軽量設計Issueである。  
設計を重くしすぎず、次Issueから実装に入る。

---

# Completion Report

※ Codex が記入

## Summary

- `docs/ROADMAP.md` を更新し、Phase 2 を抽選申込み自動化中心の実装順に合わせて整理した
- `docs/ARCHITECTURE.md` に、既存 service を再利用する抽選申込み自動化の最小構成を追記した
- `docs/DEVELOPMENT.md` に Speed Priority Policy を追記し、完璧な抽象化より早く動かす方針を明記した
- 通知機能とスケジューラ機能を Phase 2 対象外として明記した
- CAPTCHA / reCAPTCHA の手動認証待機方針を維持することを再確認した

## Changed Files

- `docs/ROADMAP.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase2/0020-auto-reservation-architecture.md`

## Verification Result

- `git status`
  `docs/ROADMAP.md`, `docs/ARCHITECTURE.md`, `docs/DEVELOPMENT.md` が変更、`docs/issues/phase2/0020-auto-reservation-architecture.md` が未追跡として表示された
- `python -m compileall .`
  成功
- `python setup.py --name`
  `court_reserv`

## Follow-up Items

- 0021 Lottery Automation Core
- 0022 Lottery Candidate Collection
- 0023 Lottery Candidate Ranking
- 0024 Lottery Entry Workflow
- 0025 Lottery Automation Dry-run Expansion
- 0026 Retry / Recovery
