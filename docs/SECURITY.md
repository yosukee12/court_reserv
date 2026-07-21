# Security

## Sensitive Data

以下は個人情報または秘密情報に該当する可能性があります。

- 利用者 ID
- パスワード
- `config.ini`
- `config.local.ini`
- `.env`
- 生成された CSV
- `debug_pages/` の HTML / PNG
- ログファイル

## Rules

- 実際の認証情報をドキュメントやサンプル設定に書かない
- 認証情報はコードへ直書きせず、`config.local.ini` または `.env` で管理する
- `court_reserv/config.ini` は秘密情報を含まない安全なベース設定だけを Git 管理する
- `config.local.ini`, `.env`, ログ, CSV, debug HTML は Git 管理対象にしない
- 画面ダンプには個人情報が含まれる前提で扱う
- CAPTCHA / reCAPTCHA の回避実装は行わない

## Incident Handling

- 認証情報を誤ってコミットした場合は履歴を含めて対処方針を確認する
- 外部共有前に `output/`, `logs/`, `debug_pages/` を確認する
