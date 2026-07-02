# Issue 0014: Model Foundation

## Status

Done

---

## Summary

今後の自動予約・予約戦略エンジン実装に向けて、ID・施設・予約枠を表す基本モデルを `court_reserv/models/` に追加する。

このIssueではモデル定義のみを行い、既存の抽選・予約・空き確認・ID管理の挙動は変更しない。

---

## Background

Issue 0007〜0013 により、Browser / Service 層の責務分離が進んだ。

次の段階では、自動予約や予約戦略エンジンを実装する前に、以下の概念を明確にする必要がある。

- 利用者ID / アカウント
- 施設 / 公園 / 種目
- 空き枠 / 予約枠
- 希望条件

現在はこれらが辞書・CSV行・画面値・固定文字列として混在しているため、最小限の dataclass を追加して今後の土台を作る。

---

## Goal

- `Account` モデルを追加する
- `Facility` モデルを追加する
- `Slot` モデルを追加する
- 必要に応じて `ReservationPreference` など将来用の軽量モデルを追加する
- 既存挙動は変更しない
- 既存コードへの大規模適用は行わない

---

## Scope

### In Scope

- `court_reserv/models/account.py` の追加
- `court_reserv/models/facility.py` の追加
- `court_reserv/models/slot.py` の追加
- 必要に応じた `court_reserv/models/preference.py` の追加
- `court_reserv/models/__init__.py` の更新
- docs/ARCHITECTURE.md 更新
- docs/DEVELOPMENT.md 更新
- Completion Report 更新

### Out of Scope

- 既存CSVフォーマット変更
- `IdManagerService` への大規模適用
- `LotteryService` への大規模適用
- `ReservationService` への大規模適用
- `AvailabilityService` への大規模適用
- Selenium 4 移行
- 自動予約機能追加
- 予約戦略エンジン追加
- UI変更

---

## Proposed Models

### Account

```python
@dataclass
class Account:
    user_id: str
    password: str | None = None
    name: str | None = None
    is_active: bool = True
```

### Facility

```python
@dataclass
class Facility:
    park_id: str
    park_name: str
    facility_id: str
    facility_name: str
    sport_id: str = "130"
    sport_name: str = "テニス"
```

### Slot

```python
@dataclass
class Slot:
    date: str
    time_range: str
    facility: Facility | None = None
    court_name: str | None = None
    status: str | None = None
```

### ReservationPreference

```python
@dataclass
class ReservationPreference:
    preferred_weekdays: list[str] = field(default_factory=list)
    preferred_time_ranges: list[str] = field(default_factory=list)
    preferred_facilities: list[Facility] = field(default_factory=list)
```

型やフィールド名は既存コードとの相性を優先して調整してよい。

---

## Automation Policy

- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない
- CAPTCHA / reCAPTCHA が表示された場合は手動認証を待つ
- 手動認証後は既存フローを継続する
- 外部CAPTCHAサービスは利用しない

---

## Implementation Plan

1. `models/account.py` を追加する
2. `models/facility.py` を追加する
3. `models/slot.py` を追加する
4. 必要に応じて `models/preference.py` を追加する
5. `models/__init__.py` を更新する
6. 既存コードへの適用は最小限に留める
7. docs/ARCHITECTURE.md を更新する
8. docs/DEVELOPMENT.md を更新する
9. Verification を実行する
10. Completion Report を記入する

---

## Target Files

- court_reserv/models/account.py
- court_reserv/models/facility.py
- court_reserv/models/slot.py
- court_reserv/models/preference.py
- court_reserv/models/__init__.py
- docs/ARCHITECTURE.md
- docs/DEVELOPMENT.md
- docs/issues/phase1/0014-model-foundation.md

---

## Tasks

- [x] Account モデルを追加する
- [x] Facility モデルを追加する
- [x] Slot モデルを追加する
- [x] ReservationPreference モデルを追加する
- [x] models/__init__.py を更新する
- [x] 既存コードへの影響を最小化する
- [x] docs を更新する
- [x] compileall を実行する
- [x] Completion Report を記入する

---

## Acceptance Criteria

- [x] `Account` モデルが追加されている
- [x] `Facility` モデルが追加されている
- [x] `Slot` モデルが追加されている
- [x] `ReservationPreference` モデルが追加されている
- [x] 既存CSVフォーマットを変更していない
- [x] 既存サービスの挙動を変更していない
- [x] UI挙動を変更していない
- [x] Selenium 4 移行を行っていない
- [x] 自動予約機能を追加していない
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
python - <<'PY'
from court_reserv.models import Account, Facility, Slot, ReservationPreference

account = Account(user_id="sample")
facility = Facility(
    park_id="1301270",
    park_name="府中の森公園",
    facility_id="12700020",
    facility_name="テニス（人工芝）",
)
slot = Slot(date="2026-01-01", time_range="09:00-11:00", facility=facility)
preference = ReservationPreference(
    preferred_weekdays=["土", "日"],
    preferred_time_ranges=["09:00-11:00"],
    preferred_facilities=[facility],
)

print(account)
print(slot)
print(preference)
PY
```

---

## Notes

- このIssueはモデルの土台作りのみを対象とする。
- 既存コードへの全面適用は次Issue以降で検討する。
- 自動予約機能や予約戦略エンジンは別Issueで扱う。
- CAPTCHA / reCAPTCHA の回避・突破・自動認証は実装しない。

---

# Completion Report

※ Codex が記入

## Summary

- `court_reserv/models/` に `Account`、`Facility`、`Slot`、`ReservationPreference` の軽量 dataclass を追加した。
- `models/__init__.py` を更新し、モデル群を import 可能にした。
- 既存サービスや CSV フォーマットには適用せず、将来の自動予約・予約戦略エンジン向けの基盤追加に留めた。
- `docs/ARCHITECTURE.md` と `docs/DEVELOPMENT.md` に Model 層の方針を追記した。

---

## Changed Files

- `court_reserv/models/account.py`
- `court_reserv/models/facility.py`
- `court_reserv/models/slot.py`
- `court_reserv/models/preference.py`
- `court_reserv/models/__init__.py`
- `docs/ARCHITECTURE.md`
- `docs/DEVELOPMENT.md`
- `docs/issues/phase1/0014-model-foundation.md`

---

## Verification Result

- `git status` 実行
- `python -m compileall .` 成功
- `python setup.py --name` で `court_reserv` を確認
- Issue 内のモデル import 確認スクリプトを実行し、各モデルの import とインスタンス生成を確認
- GUI起動確認 (`python court_reserv/court_reserv.py`, `python -m court_reserv.ui.app`) は今回のスコープ外かつ実行環境依存のため未実施

---

## Follow-up Items

- 各 service 層でのモデル段階適用
- CSV 行や画面値からモデルへ変換する mapper 整備
- 自動予約・予約戦略エンジン向けの preference 拡張
