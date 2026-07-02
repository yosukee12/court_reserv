# Development

## Workflow

このリポジトリでは Issue Driven Development を採用します。

1. `docs/issues/` に Issue を作成する
2. Scope / Out of Scope / Acceptance Criteria を明記する
3. 実装担当は Issue に書かれた内容だけを変更する
4. 実装後は Completion Report を記入する

## Implementation Rules

- Issue 外の仕様変更は行わない
- 既存動作は Issue に明記がない限り変更しない
- 大規模リファクタリングは Issue 単位で行う
- 不明点は仮実装せず、Issue または Completion Report に残す

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
