# Issue 0030: 設定画面の申込日入力改善

## Status

In Progress

## Scope

- 共通・ID別申込み枠にカレンダー選択、日付クリア、保存時検証を追加する。
- 編集画面の親・フォーカス・入力制御を整理する。
- 標準 Tkinter のみを使用する。

## Out of Scope

- 申込み処理、認証、既存の未コミット変更の整理。

## Acceptance Criteria

- カレンダーの月移動と日付選択ができ、既存日付を表示できる。
- 手入力の途中では空欄を許し、不正な日付は保存時に案内する。
- 編集終了後に親設定画面の入力制御を復元する。
- 関連テストを実行し、実画面で未確認の項目を明記する。

## Completion Report

### 変更内容

- 共通・ID別枠へ前月・翌月移動付きカレンダー、全選択、クリアを追加。
- 手入力を妨げず、保存時に実在する YYYY-MM-DD の日付を検証。
- 設定画面を編集ダイアログの親にし、初期フォーカスと終了後の grab を復元。

### 変更ファイル

- `court_reserv/court_reserv.py`
- `court_reserv/ui/date_entry.py`
- `tests/test_date_entry.py`
- 本 Issue

### 確認結果

- `PYTHONDONTWRITEBYTECODE=1 python -m unittest discover -s tests -p test_date_entry.py -v`: 4件成功。
- AST 構文確認、`git diff --check`: 成功。
- Tk ウィジェットの非表示スモーク: 日付選択・クリア・テキスト挿入と削除に成功。macOS のサービス接続警告あり。
- pytest は環境に未導入のため未実行。
- 実画面の見た目、物理キーによる削除、設定画面との往復操作は未確認。削除不調の再現・解消は未確定のため In Progress を維持。
