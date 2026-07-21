# Issue 0027: Lottery Submission Recovery

## Status

Done

---

## Summary

抽選申込み送信後に reCAPTCHA が表示された場合でも、ユーザーの手動認証後に申込み枠選択を再設定し、最大1回だけ再送信できるようにする。

この Issue では、reCAPTCHA の回避・突破・自動認証は行わない。

---

## Background

Issue 0021〜0026 により、抽選申込み自動化は以下まで実装済み。

- Preference 読み込み
- `default_entries` / `account_overrides`
- 指定曜日フィルタ
- 週送り探索
- 現在申込数取得
- 候補選択
- ユーザー確認後の送信

一方、legacy の半自動申込みでは、送信後に reCAPTCHA が表示されると処理が止まる問題があった。

通常時は、送信後に popup / alert が表示され、その後申込み完了画面へ進む。  
しかし reCAPTCHA が表示された場合、ユーザーが手動認証した後に、申込み枠選択プルダウンを再度選ばされることがある。

そのため、保存済みの選択情報から申込み枠選択を再設定し、最大1回だけ再送信できる recovery を追加する。

---

## Goal

- 送信後の状態を判定する
- 通常 popup / alert を処理する
- 申込み完了画面を検知する
- reCAPTCHA 表示時に手動認証を待つ
- 手動認証後に申込み枠選択プルダウンを再設定する
- 最大1回だけ再送信する
- 失敗時に debug HTML / DOM summary を保存する
- 結果を JSON に保存する

---

## Scope

### In Scope

- 送信後状態判定
- 通常 popup / alert 処理
- 申込み完了画面判定
- reCAPTCHA 表示判定
- reCAPTCHA 手動認証待機
- 申込み枠選択プルダウン再選択
- 最大1回の再送信
- 結果 JSON 保存
- 失敗時 debug HTML / DOM summary 保存
- README / docs 更新
- Completion Report 更新

### Out of Scope

- reCAPTCHA 回避
- reCAPTCHA 自動認証
- 外部 CAPTCHA サービス利用
- 申込み候補選定ロジック変更
- `default_entries` / `account_overrides` 仕様変更
- 週送り探索仕様変更
- GUI変更
- 通知
- スケジューラ

---

## Submission Recovery Policy

### 通常時

```text
候補選択
    ↓
確認画面
    ↓
ユーザーが yes
    ↓
送信ボタン押下
    ↓
popup / alert
    ↓
OK
    ↓
申込み完了画面
```

### reCAPTCHA 発生時

```text
候補選択
    ↓
確認画面
    ↓
ユーザーが yes
    ↓
送信ボタン押下
    ↓
reCAPTCHA 表示
    ↓
ユーザーが手動認証
    ↓
申込み枠選択プルダウンを再設定
    ↓
再送信
    ↓
popup / alert
    ↓
申込み完了画面
```

---

## Retry Policy

- reCAPTCHA 後の再送信は最大1回まで
- 2回目以降の reCAPTCHA / エラーでは停止する
- 停止時は debug HTML / DOM summary を保存する
- 無限ループしない

---

## Automation Policy

- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない
- CAPTCHA / reCAPTCHA が表示された場合はユーザーの手動認証を待つ
- 手動認証後は既存フローを継続する
- 外部 CAPTCHA サービスは利用しない

---

## Authentication Policy

認証情報は `preferences.yaml` に入れない。

認証情報は既存方式を利用する。

優先順位:

1. ID CSV / `IdManagerService`
2. `config.local.ini`
3. `.env`

禁止事項:

- コードへの ID / password 直書き
- `preferences.example.yaml` への ID / password 記載
- README / docs への実 ID・実パスワード記載
- ZIP 成果物への `.env` / `config.local.ini` / 実 ID CSV 混入

---

## Proposed Files

主に既存ファイルを更新する。

```text
court_reserv/services/lottery.py
court_reserv/services/lottery_entry_workflow.py
```

必要に応じて補助モジュールを追加してよい。

---

## Implementation Plan

1. 送信後状態判定を追加する
2. 通常 popup / alert を処理する
3. 完了画面を判定する
4. reCAPTCHA 表示を検知する
5. reCAPTCHA 表示時はユーザーに手動認証を促す
6. 手動認証完了後、保存済み選択情報から申込み枠選択プルダウンを再設定する
7. 最大1回だけ再送信する
8. 成功 / 失敗 / recovery 発生有無を JSON に保存する
9. 失敗時に debug HTML / DOM summary を保存する
10. README / docs を更新する
11. Verification を実施する
12. Completion Report を記入する

---

## Tasks

- [x] 送信後状態判定を追加する
- [x] popup / alert 処理を追加する
- [x] 完了画面判定を追加する
- [x] reCAPTCHA 表示判定を追加する
- [x] reCAPTCHA 手動認証待機を追加する
- [x] 申込み枠選択プルダウン再設定を追加する
- [x] 最大1回の再送信を追加する
- [x] 成功 / 失敗結果を JSON 保存する
- [x] 失敗時 debug HTML を保存する
- [x] 失敗時 DOM summary を保存する
- [x] README / docs を更新する
- [x] Verification を実施する
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [ ] 通常 popup / alert を処理できる
- [ ] 申込み完了画面を判定できる
- [ ] reCAPTCHA 表示時に手動認証待機できる
- [ ] 手動認証後に申込み枠選択プルダウンを再設定できる
- [ ] reCAPTCHA 後の再送信は最大1回まで
- [ ] 無限ループしない
- [ ] 成功 / 失敗 / recovery 発生有無を JSON 保存できる
- [ ] 失敗時 debug HTML / DOM summary を保存できる
- [ ] reCAPTCHA 回避・突破・自動認証を実装していない
- [ ] 申込み候補選定ロジックを変更していない
- [ ] 既存 GUI を変更していない
- [ ] `python -m compileall .` が成功する

---

## Verification

```bash
git status
python -m compileall .
python setup.py --name
python scripts/lottery_entry_workflow.py --help
```

可能なら:

```bash
python scripts/lottery_entry_workflow.py --preferences config/preferences.example.yaml
```

実サイトで reCAPTCHA が出ない場合は、通常送信経路のみ確認し、reCAPTCHA recovery は未確認として Completion Report に記載する。

---

## Notes

- この Issue は送信後 reCAPTCHA recovery のみを対象とする。
- 申込み候補選定・週送り探索・曜日フィルタは変更しない。
- GUI統合は次 Issue 以降で扱う。
- 通知・スケジューラは実装しない。

---

# Completion Report

## Summary

- `LotteryService.submit_selected_slots()` に送信後状態判定を追加し、`alert` / `申込み完了画面` / `reCAPTCHA` / `error` を見分けながら進めるようにした
- reCAPTCHA 表示時は既存 `LoginService.wait_for_manual_captcha()` で手動認証を待ち、認証完了後に保存済みの `apply` 選択値を復元して最大 1 回だけ再送信する recovery を追加した
- 送信結果 JSON に `recovery_triggered` / `recovery_attempts` / `states` / `debug_files` を残すようにし、失敗時は `output/debug_pages/` に HTML / DOM summary を保存できるようにした
- `scripts/lottery_entry_workflow.py` は非対話実行でも EOF で落ちず未送信扱いで終了するようにし、workflow 結果の出力を安定化した

## Changed Files

- `court_reserv/services/lottery.py`
- `court_reserv/services/lottery_entry_workflow.py`
- `scripts/lottery_entry_workflow.py`
- `README.md`
- `docs/ROADMAP.md`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase2/0027-lottery-submission-recovery.md`

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
  実行した。通常経路の確認として、週送り探索後に `2026-08-15` の2枠を選択し、非対話実行のため送信確認は空入力扱いとなり、未送信で正常終了した
  送信 recovery 自体は実サイトで reCAPTCHA が発生しなかったため未確認
  途中で一度だけサイト側の「データ通信を正しく行うことができませんでした」アラートに遭遇したが、再試行では正常終了した

## Unverified

- 実サイトで送信後に reCAPTCHA が表示されたケースでの recovery 成功
- recovery 後の popup / alert から完了画面遷移までの一連の正常終了
- recovery 失敗時の `debug_files` が実運用で十分な情報量になっているか

## Follow-up Items

- 実サイトで reCAPTCHA が表示されたケースの再現確認
- サイト側通信エラー発生時の Retry / Recovery
- GUI からの recovery 状態表示

※ Codex が記入

## Summary

---

## Changed Files

---

## Verification Result

---

## Follow-up Items

- Workflow Validation / Retry
- GUI Integration
