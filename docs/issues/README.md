# Issue Driven Development

## Purpose

このプロジェクトでは、大規模リファクタリングや仕様変更を無秩序に進めず、Issue 単位で安全に変更を積み上げます。

Phase 1 は `docs/issues/phase1/0016-phase1-wrap-up.md` で完了とし、以後は Phase 2 の自動予約基盤整備へ進みます。

## Basic Rules

- 1 Issue = 1 目的で進める
- Scope に書かれた内容だけを実装する
- Out of Scope の内容は実装しない
- 既存動作は Issue に明記がない限り変更しない
- 不明点は仮実装せず、Completion Report に残す
- Issue 外の実装は禁止し、必要なら提案に留める
- 小さな PR / 小さな差分を優先する

## Automation Policy

- CAPTCHA / reCAPTCHA の回避、突破、自動認証は実装しない
- 外部 CAPTCHA サービスは利用しない
- reCAPTCHA が表示された場合は、ユーザーが手動で認証するまで待機する
- 手動認証完了後は、そのまま処理を継続できる設計を維持する
- 今後の自動化機能も、この方針を前提に設計する

## Directory Policy

- `phase1/`
  基盤整備、初期リファクタリング、準備作業。Issue 0016 で完了
- `phase2/`
  予約戦略、自動予約基盤、通知連携前提の機能拡張
- `tech-debt/`
  技術的負債の整理や横断的改善

## Phase 2 Policy

- Phase 2 は予約戦略・自動予約基盤の整備を中心に進める
- 仕様追加は必ず個別 Issue へ分割し、設計判断と実装範囲を明示する
- 小さな PR / 小さな差分を優先し、複数テーマを同時に混ぜない

## Phase 3 Policy

- Phase 3 は運用改善・機能拡張を中心に進める
- Multi Facility、複数自治体対応、ログ改善、テスト強化、性能改善は Phase 3 で扱う
- 既存 GUI と既存サービスの挙動は、明示的な Issue がない限り維持する

## Required Sections

各 Issue には少なくとも以下を含めます。

- Summary
- Background
- Goal
- Scope
- Implementation Plan
- Target Files
- Tasks
- Acceptance Criteria
- Verification
- Completion Report

## Completion Report

実装担当は作業完了時に以下を記入します。

- Summary
- Changed Files
- Verification Result
- Follow-up Items
