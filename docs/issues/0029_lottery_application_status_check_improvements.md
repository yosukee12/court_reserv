# Issue 0029: 抽選申込み状況確認ワークフロー改善

## 背景

抽選申込み状況確認ボタンを追加したが、以下の要件が不足している。

- CSV出力先パスを指定できない
- 抽選申込状況画面までは遷移できているが、申込み状況を取得できていない
- legacy の `check_lottery` と同等の取得結果にしたい

## 要件

### 1. CSV出力先指定

GUIでCSV出力先フォルダを指定できるようにする。

- 未指定の場合は `ID_list.csv` と同じフォルダへ出力する
- 指定されている場合は指定フォルダへ出力する
- 出力ファイル名は既存仕様を維持する

### 2. 抽選申込み状況取得の修正

legacy 実装を必ず確認する。

対象:

- `court_reserv/legacy/bk_court_reserv.py`
- `check_lottery(...)`

特に以下を比較する。

- 遷移Action
  - `gLotWTransLotCancelListAction`
- 画面遷移後の待機条件
- BeautifulSoup の生成タイミング
- HTML解析ロジック
- テニス以外の除外条件
- 日付・時間帯の抽出正規表現

現行実装で取得件数が0件になる原因を特定し、legacy と同等の抽選申込み状況を取得できるようにする。

### 3. 調査ログ追加

取得直前・直後に以下を INFO または DEBUG ログへ出す。

- current_url
- title
- displayNo
- table存在有無
- 一覧行数
- BeautifulSoup抽出件数
- 抽出後件数
- 除外後件数

取得件数が0件の場合は `output/debug_pages/` にHTMLを保存する。

## 完了条件

- GUIからCSV出力先フォルダを指定できる
- 未指定時はID_list.csvと同じフォルダへ出力される
- 抽選申込み状況がlegacyと同等に取得できる
- CSVへ正常に出力される
- 既存の抽選申込みワークフローに影響しない
- 以下が成功する
  - `python -m compileall .`
  - `python setup.py --name`
  - `python scripts/lottery_entry_workflow.py --help`
