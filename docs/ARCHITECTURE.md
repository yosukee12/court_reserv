# Architecture

## Current State

現状の実装は `court_reserv/court_reserv.py` に UI、Selenium 操作、業務処理が集約されており、`court_reserv/manage_id.py` が ID CSV の読込・書出しと有効確認を担っています。

主なエントリーポイント:

- `court_reserv/court_reserv.py`
- `run_collect_slots.py`

## Current Module Layout

- `court_reserv.ui`
  現在は分離先ディレクトリのみ存在します
- `court_reserv.browser`
  Selenium 操作の分離先ディレクトリです
- `court_reserv.services`
  予約・抽選処理の分離先ディレクトリです
- `court_reserv.models`
  スロットや申込結果のデータ構造の分離先ディレクトリです
- `court_reserv.utils`
  CSV、設定、共通関数の分離先ディレクトリです
- `court_reserv.config`
  設定読み込みの分離先ディレクトリです

## Design Notes

- この段階では既存の Selenium / Tkinter / 予約処理ロジックは維持します
- 大規模リファクタリングや責務分離は後続 Issue で段階的に行います
- CAPTCHA / reCAPTCHA は手動対応前提です
