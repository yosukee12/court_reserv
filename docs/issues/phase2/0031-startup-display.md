# Issue 0031: GUI を現在のディスプレイで起動

## Status

In Progress

## Scope

- macOS で起動時のマウスポインターがある画面にメインウィンドウを配置する。
- 保存済みサイズを利用し、過去の座標は復元しない。
- 標準ライブラリと macOS 標準 API のみ使用する。

## Out of Scope

- 予約処理、設定ダイアログの変更、他 OS の複数画面対応。

## Acceptance Criteria

- 画面の左右・上下配置に応じた位置計算をテストする。
- 画面情報取得に失敗しても GUI は起動する。
- 実機の複数画面で未確認の項目は明記する。

## Completion Report

- `court_reserv/ui/window_position.py` に macOS CoreGraphics によるポインター画面取得と配置処理を追加。
- `court_reserv/court_reserv.py` の保存座標復元を置換。両起動経路に適用。
- 前回サイズを画面内に収め、取得失敗時は Tk の既定画面へ配置。
- `tests/test_window_position.py`: unittest 5件成功（保存座標無視、左右上下配置、小画面、不正設定、取得失敗）。
- AST 構文確認と Tk 非表示スモーク（負座標・フォールバック配置）成功。
- `git diff --check` 成功。
- この実行環境ではネイティブの現在画面取得が失敗。実機での複数画面起動は未確認のため In Progress を維持。
- 新規依存なし。既存未コミット変更は保持。
