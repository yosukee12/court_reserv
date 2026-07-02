# Issue Driven Development

## Purpose

このプロジェクトでは、大規模リファクタリングや仕様変更を無秩序に進めず、Issue 単位で安全に変更を積み上げます。

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
  基盤整備、初期リファクタリング、準備作業
- `phase2/`
  機能分離、改善、拡張
- `tech-debt/`
  技術的負債の整理や横断的改善

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
