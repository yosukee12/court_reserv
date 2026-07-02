# Issue 0006: Court_Reserv Split Phase 1

## Status

- [ ] Draft
- [ ] Ready
- [ ] In Progress
- [ ] Review
- [x] Done

---

## Summary

巨大化している `Court_Reserv` クラスの分割に向けて、まずはアプリ起動・UI初期化・互換エントリーポイントを整理する。

このIssueでは、予約処理・ログイン処理・Selenium操作の中身は変更せず、既存挙動を保ったまま分割準備を行う。

---

## Background

現在、`court_reserv/court_reserv.py` には以下が集中している。

- Tkinterアプリ起動
- `Court_Reserv` クラス定義
- UIウィジェット作成
- Selenium操作
- ログイン処理
- 抽選申込
- 予約確定
- 空き確認
- debug出力

Issue 0001〜0005 で、基盤・設定・docs・legacy整理は完了した。

次の段階として、`Court_Reserv` クラスの段階的分割を開始する。

---

## Goal

- GUIアプリの正式起動口を `court_reserv/ui/app.py` に用意する
- 既存起動コマンド `python court_reserv/court_reserv.py` を維持する
- `Court_Reserv` クラス本体の大規模分割前に、UI層の受け皿を作る
- 今後の分割に備えて import / entrypoint を整理する
- 予約・ログイン・Seleniumの挙動は変更しない

---

## Scope

### In Scope

- `court_reserv/ui/app.py` の追加
- `court_reserv/ui/__init__.py` の更新
- GUI起動用 `main()` の整理
- 既存 `court_reserv/court_reserv.py` の互換エントリーポイント維持
- `README.md` の Directory Overview 表示崩れ修正
- `README.md` の起動方法更新
- `docs/ARCHITECTURE.md` 更新
- `docs/DEVELOPMENT.md` 更新
- Completion Report 更新

### Out of Scope

- `Court_Reserv` クラスの中身の大規模分割
- Selenium 4 移行
- ログイン処理変更
- 予約処理変更
- CAPTCHA / reCAPTCHA 処理変更
- 施設・時間帯・予約戦略の変更
- 自動予約機能追加
- テスト基盤の大幅追加

---

## Implementation Plan

1. `court_reserv/ui/app.py` を追加する
2. `court_reserv/ui/app.py` に GUI 起動用の `main()` を用意する
3. 既存の `Court_Reserv` クラスは当面 `court_reserv/court_reserv.py` に残す
4. `court_reserv/court_reserv.py` の既存 `main()` は互換性維持のため残す
5. 必要に応じて `court_reserv/ui/__init__.py` から `main` を export する
6. README の起動方法を更新する
7. README の Directory Overview の `scripts/` インデント崩れを修正する
8. docs/ARCHITECTURE.md に UI entrypoint 方針を追記する
9. docs/DEVELOPMENT.md に互換エントリーポイント維持ルールを追記する
10. Verification を実行する
11. Completion Report を記入する

---

## Target Files

- court_reserv/ui/app.py
- court_reserv/ui/__init__.py
- court_reserv/court_reserv.py
- README.md
- docs/ARCHITECTURE.md
- docs/DEVELOPMENT.md
- docs/issues/phase1/0006-court-reserv-split-phase1.md

---

## Tasks

- [x] `court_reserv/ui/app.py` を追加する
- [x] UI起動用 `main()` を追加する
- [x] 既存起動コマンドを維持する
- [x] `court_reserv/ui/__init__.py` を更新する
- [x] README の起動方法を更新する
- [x] README の Directory Overview を修正する
- [x] docs/ARCHITECTURE.md を更新する
- [x] docs/DEVELOPMENT.md を更新する
- [x] 動作確認を行う
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] `python court_reserv/court_reserv.py` が引き続き正式起動パスとして動作する
- [x] `python -m court_reserv.ui.app` でもGUI起動できる
- [x] `Court_Reserv` クラスの予約処理・ログイン処理・Selenium操作の中身を変更していない
- [x] README の Directory Overview の `scripts/` 表示崩れが修正されている
- [x] README に新旧起動方法が記載されている
- [x] docs/ARCHITECTURE.md に UI entrypoint 方針が記載されている
- [x] docs/DEVELOPMENT.md に互換エントリーポイント維持ルールが記載されている
- [x] CAPTCHA / reCAPTCHA 方針を変更していない
- [x] `python -m compileall .` が成功する

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
```

可能なら以下も実行する。

```bash
python -m court_reserv.ui.app
python court_reserv/court_reserv.py
```

GUI起動確認が環境上難しい場合は、その理由を Completion Report に記載する。

---

## Notes

- このIssueは大規模分割ではなく、UI entrypoint の整理である。
- `Court_Reserv` クラスの中身は次Issue以降で段階的に分割する。
- 互換性のため、既存の `python court_reserv/court_reserv.py` は維持する。
- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない。

---

# Completion Report

※ Codex が記入

## Summary

UI エントリーポイント整理の Phase 1 として `court_reserv/ui/app.py` を追加し、`python -m court_reserv.ui.app` で GUI を起動できる入口を用意した。既存の正式起動パス `python court_reserv/court_reserv.py` はそのまま維持し、`Court_Reserv` クラス本体の大規模分割は行っていない。

あわせて `court_reserv/ui/__init__.py`、`README.md`、`docs/ARCHITECTURE.md`、`docs/DEVELOPMENT.md` を更新し、新旧起動方法と UI entrypoint 方針、互換エントリーポイント維持ルールを明記した。Selenium / Tkinter / 予約ロジック、ログイン処理、CAPTCHA / reCAPTCHA 処理には変更を加えていない。

---

## Changed Files

- `court_reserv/ui/app.py`
- `court_reserv/ui/__init__.py`
- `court_reserv/court_reserv.py`
- `README.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase1/0006-court-reserv-split-phase1.md`

---

## Verification Result

- `git status` を実行し、対象ファイルの変更を確認
- `python -m compileall .` を実行し、構文エラーがないことを確認
- `python setup.py --name` を実行し、`court_reserv` が返ることを確認
- `python -m court_reserv.ui.app` は GUI 環境依存のため、この環境では起動確認まで実施していない
- `python court_reserv/court_reserv.py` も GUI 環境依存のため、この環境では起動確認まで実施していない

---

## Follow-up Items

- 次 Issue で `Court_Reserv` 内の UI 初期化責務を段階的に `court_reserv/ui/` へ移す
- その後の Phase で browser / services への責務分離を進める
