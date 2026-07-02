# Development

## Issue Driven Development

このリポジトリでは Issue Driven Development を採用します。

1. `docs/issues/` に Issue を作成する
2. Scope / Out of Scope / Acceptance Criteria を明記する
3. 実装担当は Issue に書かれた内容だけを変更する
4. 実装後は Completion Report を記入する

## Roles

- ChatGPT
  設計判断、方針整理、Issue 化の支援を担当する
- Codex
  Issue に記載された範囲の実装、検証、Completion Report 記入を担当する

## Completion Report

- 実装後は対象 Issue の Completion Report を必ず更新する
- 実装できなかった点や未確認事項は仮実装せず、Completion Report に残す
- Follow-up Items には次 Issue で扱うべき事項だけを記載する

## Implementation Rules

- Issue 外の仕様変更は行わない
- 既存動作は Issue に明記がない限り変更しない
- 大規模リファクタリングは Issue 単位で行う
- 不明点は仮実装せず、Issue または Completion Report に残す
- 小さな PR / 小さな差分を優先する
- Selenium の実装変更、GUI 変更、予約ロジック変更は明示的な Issue がある場合のみ行う
- CAPTCHA / reCAPTCHA の回避や自動認証は実装しない
- reCAPTCHA が表示された場合は手動認証完了まで待機し、その後に処理を継続する前提で実装する

## Local Setup

- `config.example.ini` を参考に `court_reserv/config.ini` を作成する
- 個人情報や認証情報を含むファイルは Git に追加しない
- 出力物は `logs/`, `output/`, `debug_pages/` を利用する

## Useful Commands

```bash
git status
find . -maxdepth 3 -type d | sort
find docs/issues -maxdepth 3 -type f | sort
python -m compileall .
```
