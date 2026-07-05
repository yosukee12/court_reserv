# -*- coding: utf-8 -*-
from __future__ import annotations

import configparser
import contextlib
import datetime
import logging
import threading
import time
from pathlib import Path

try:
    from .browser import BrowserSession, LoginService, NavigationService
    from .config import (
        get_debug_output_dir,
        get_output_base_path,
        load_config,
        load_preferences_data,
        load_reservation_preference,
        save_preferences_data,
    )
    from .manage_id import Manage_Id as mi
    from .services import (
        AvailabilityService,
        IdManagerService,
        LotteryEntrySlotCollector,
        LotteryEntryWorkflowService,
        LotteryResultWorkflowService,
        LotteryService,
        ReservationConfirmationWorkflowService,
        ReservationService,
    )
except Exception:
    from browser import BrowserSession, LoginService, NavigationService
    from config import (
        get_debug_output_dir,
        get_output_base_path,
        load_config,
        load_preferences_data,
        load_reservation_preference,
        save_preferences_data,
    )
    from manage_id import Manage_Id as mi
    from services import (
        AvailabilityService,
        IdManagerService,
        LotteryEntrySlotCollector,
        LotteryEntryWorkflowService,
        LotteryResultWorkflowService,
        LotteryService,
        ReservationConfirmationWorkflowService,
        ReservationService,
    )

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


config = load_config()

logfile = config["PATH"]["LOG_PATH"] + "/court_reserv.log"
log_fmt = "%(asctime)s - %(levelname)s - %(message)s"
log_level_name = config.get("LOG", "LEVEL", fallback="INFO").upper()
log_level = getattr(logging, log_level_name, logging.INFO)
logging.basicConfig(filename=logfile, format=log_fmt, level=log_level)

check_lottery_csv = (
    config["PATH"]["OUTPUT_CSV_PATH"]
    + "/check_lottery_{0}.csv".format(datetime.date.today())
)
check_result_csv = (
    config["PATH"]["OUTPUT_CSV_PATH"]
    + "/check_result_{0}.csv".format(datetime.date.today())
)
determined_csv = (
    config["PATH"]["OUTPUT_CSV_PATH"]
    + "/determined_result_{0}.csv".format(datetime.date.today())
)
check_reserv_csv = (
    config["PATH"]["OUTPUT_CSV_PATH"]
    + "/check_reserv_{0}.csv".format(datetime.date.today())
)
alive_id_list_csv = (
    config["PATH"]["OUTPUT_CSV_PATH"]
    + "/ID_list_alive_{0}.csv".format(datetime.date.today())
)
dead_id_list_csv = (
    config["PATH"]["OUTPUT_CSV_PATH"]
    + "/ID_list_dead_{0}.csv".format(datetime.date.today())
)


class GuiLogHandler(logging.Handler):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback
        self.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

    def emit(self, record):
        try:
            self.callback(self.format(record))
        except Exception:
            pass


class GuiStreamWriter:
    def __init__(self, callback, level="INFO"):
        self.callback = callback
        self.level = level
        self._buffer = ""

    def write(self, text):
        self._buffer += str(text)
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            if line.strip():
                self.callback(f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} [{self.level}] {line}")

    def flush(self):
        if self._buffer.strip():
            self.callback(
                f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} [{self.level}] {self._buffer.strip()}"
            )
        self._buffer = ""


class Court_Reserv(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill=tk.BOTH, expand=True)
        self.master.geometry("980x780")
        self.master.minsize(900, 700)
        self.master.title("東京都テニスコート予約")
        self.repo_root = Path(__file__).resolve().parents[1]
        self.local_config_path = self.repo_root / "config.local.ini"
        self.logger = logging.getLogger("court_reserv.gui")
        self.worker_thread = None
        self.window_geometry = "980x780"

        self.browser_session = BrowserSession(config)
        self.login_service = LoginService(
            top_url=config["URL"]["TOP_URL"],
            wait_factory=self.browser_session.get_wait,
            logger=self.logger,
            show_info=self._threadsafe_show_info,
            ask_yes_no=self._threadsafe_ask_yes_no,
            sleep_func=time.sleep,
        )
        self.navigation_service = NavigationService(
            wait_factory=self.browser_session.get_wait,
            sleep_func=time.sleep,
            logger=self.logger,
            get_debug_output_dir=get_debug_output_dir,
        )
        self.lottery_service = LotteryService(
            config=config,
            browser_session=self.browser_session,
            login_service=self.login_service,
            navigation_service=self.navigation_service,
            logger=self.logger,
            show_info=self._threadsafe_show_info,
            output_id_dict=mi.output_csv_from_id_dict,
            sleep_func=time.sleep,
        )
        self.reservation_service = ReservationService(
            config=config,
            browser_session=self.browser_session,
            navigation_service=self.navigation_service,
            logger=self.logger,
            get_id_dict_from_csv=mi.get_id_dict_from_csv,
            output_id_dict=mi.output_csv_from_id_dict,
            sleep_func=time.sleep,
        )
        self.availability_service = AvailabilityService(
            config=config,
            browser_session=self.browser_session,
            navigation_service=self.navigation_service,
            logger=self.logger,
            get_debug_output_dir=get_debug_output_dir,
            sleep_func=time.sleep,
        )
        self.id_manager_service = IdManagerService(config=config, sleep_func=time.sleep)
        self.lottery_entry_workflow_service = LotteryEntryWorkflowService(
            config=config,
            browser_session=self.browser_session,
            login_service=self.login_service,
            navigation_service=self.navigation_service,
            lottery_service=self.lottery_service,
            id_manager_service=self.id_manager_service,
            slot_collector=LotteryEntrySlotCollector(
                navigation_service=self.navigation_service,
                browser_session=self.browser_session,
                logger=self.logger,
            ),
            logger=self.logger,
        )
        self.lottery_result_workflow_service = LotteryResultWorkflowService(
            config=config,
            browser_session=self.browser_session,
            login_service=self.login_service,
            navigation_service=self.navigation_service,
            lottery_service=self.lottery_service,
            id_manager_service=self.id_manager_service,
            logger=self.logger,
        )
        self.reservation_confirmation_workflow_service = (
            ReservationConfirmationWorkflowService(
                lottery_result_workflow_service=self.lottery_result_workflow_service,
                reservation_service=self.reservation_service,
                login_service=self.login_service,
                logger=self.logger,
            )
        )

        self.id_csv_var = tk.StringVar()
        self.preferences_var = tk.StringVar()
        self.driver = None

        self._load_last_settings()
        self.master.geometry(self.window_geometry)
        self.create_widgets()
        self._install_log_handler()
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)
        self._log_message("GUI initialized.")

    def create_widgets(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        container = ttk.Frame(self, padding=16)
        container.grid(row=0, column=0, sticky=tk.NSEW)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(5, weight=1)

        title = ttk.Label(
            container,
            text="東京都テニスコート予約",
            font=("", 16, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 12))

        ttk.Label(container, text="ID CSVファイル").grid(
            row=1, column=0, sticky=tk.W, pady=4
        )
        ttk.Entry(container, textvariable=self.id_csv_var).grid(
            row=1, column=1, sticky=tk.EW, padx=(8, 8), pady=4
        )
        ttk.Button(
            container,
            text="参照",
            command=self.browse_id_csv,
        ).grid(row=1, column=2, sticky=tk.EW, pady=4)

        ttk.Label(container, text="設定YAML").grid(
            row=2, column=0, sticky=tk.W, pady=4
        )
        ttk.Entry(container, textvariable=self.preferences_var).grid(
            row=2, column=1, sticky=tk.EW, padx=(8, 8), pady=4
        )
        ttk.Button(
            container,
            text="参照",
            command=self.browse_preferences,
        ).grid(row=2, column=2, sticky=tk.EW, pady=4)

        action_frame = ttk.LabelFrame(container, text="抽選", padding=12)
        action_frame.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=(16, 16))
        for index in range(4):
            action_frame.columnconfigure(index, weight=1)

        self.button_auto_lottery = ttk.Button(
            action_frame,
            text="抽選申込み自動化",
            command=self.run_lottery_entry_workflow,
        )
        self.button_auto_lottery.grid(row=0, column=0, sticky=tk.EW, padx=(0, 8))

        self.button_lottery_result = ttk.Button(
            action_frame,
            text="抽選結果確認",
            command=self.run_lottery_result_workflow,
        )
        self.button_lottery_result.grid(row=0, column=1, sticky=tk.EW, padx=8)

        self.button_reservation_confirm = ttk.Button(
            action_frame,
            text="予約確定補助",
            command=self.run_reservation_confirmation_workflow,
        )
        self.button_reservation_confirm.grid(row=0, column=2, sticky=tk.EW, padx=8)

        self.button_settings = ttk.Button(
            action_frame,
            text="設定",
            command=self.open_settings_dialog,
        )
        self.button_settings.grid(row=0, column=3, sticky=tk.EW, padx=(8, 0))

        log_frame = ttk.LabelFrame(container, text="ログ", padding=8)
        log_frame.grid(row=4, column=0, columnspan=3, sticky=tk.NSEW)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        container.rowconfigure(4, weight=1)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            height=24,
            state=tk.DISABLED,
        )
        self.log_text.grid(row=0, column=0, columnspan=2, sticky=tk.NSEW)

        ttk.Button(log_frame, text="ログ保存", command=self.save_log).grid(
            row=1, column=0, sticky=tk.W, pady=(8, 0)
        )
        ttk.Button(log_frame, text="ログクリア", command=self.clear_log).grid(
            row=1, column=1, sticky=tk.E, pady=(8, 0)
        )

    def _install_log_handler(self):
        self.gui_log_handler = GuiLogHandler(self._append_log_line)
        logging.getLogger().addHandler(self.gui_log_handler)

    def _append_log_line(self, line):
        def write():
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, line + "\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)

        self.after(0, write)

    def _log_message(self, message, level="INFO"):
        self._append_log_line(
            f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} [{level}] {message}"
        )

    def clear_log(self):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def save_log(self):
        target = filedialog.asksaveasfilename(
            title="ログ保存",
            defaultextension=".log",
            filetypes=[("ログ", "*.log"), ("テキスト", "*.txt"), ("すべてのファイル", "*.*")],
            initialfile=f"court_reserv_gui_{datetime.datetime.now():%Y%m%d_%H%M%S}.log",
        )
        if not target:
            return
        Path(target).write_text(self.log_text.get("1.0", tk.END), encoding="utf-8")
        self._log_message(f"ログを保存しました: {target}")

    def browse_id_csv(self):
        path = filedialog.askopenfilename(
            title="ID CSV を選択",
            filetypes=[("CSV", "*.csv"), ("すべてのファイル", "*.*")],
        )
        if path:
            self.id_csv_var.set(path)

    def browse_preferences(self):
        path = filedialog.askopenfilename(
            title="設定YAML を選択",
            filetypes=[("YAML/JSON", "*.yaml *.yml *.json"), ("すべてのファイル", "*.*")],
        )
        if path:
            self.preferences_var.set(path)

    def _run_in_background(self, label, func):
        if self.worker_thread and self.worker_thread.is_alive():
            self._threadsafe_show_info("実行中", "他のWorkflowが実行中です。完了を待ってください。")
            return
        self._set_action_state(tk.DISABLED)

        def runner():
            stdout_writer = GuiStreamWriter(self._append_log_line, level="INFO")
            stderr_writer = GuiStreamWriter(self._append_log_line, level="ERROR")
            try:
                self._log_message(f"{label} を開始しました。")
                with contextlib.redirect_stdout(stdout_writer), contextlib.redirect_stderr(
                    stderr_writer
                ):
                    func()
                self._log_message(f"{label} が完了しました。")
            except Exception:
                self.logger.exception("%s failed", label)
            finally:
                stdout_writer.flush()
                stderr_writer.flush()
                self.after(0, lambda: self._set_action_state(tk.NORMAL))
                self._save_last_settings()

        self.worker_thread = threading.Thread(target=runner, daemon=True)
        self.worker_thread.start()

    def _set_action_state(self, state):
        for button in (
            self.button_auto_lottery,
            self.button_lottery_result,
            self.button_reservation_confirm,
            self.button_settings,
        ):
            button.configure(state=state)

    def _call_on_main_thread(self, func, *args, **kwargs):
        if threading.current_thread() is threading.main_thread():
            return func(*args, **kwargs)
        box = {}
        event = threading.Event()

        def callback():
            try:
                box["result"] = func(*args, **kwargs)
            except Exception as exc:
                box["error"] = exc
            finally:
                event.set()

        self.after(0, callback)
        event.wait()
        if "error" in box:
            raise box["error"]
        return box.get("result")

    def _threadsafe_show_info(self, title, message):
        return self._call_on_main_thread(messagebox.showinfo, title, message)

    def _threadsafe_ask_yes_no(self, title, message):
        return self._call_on_main_thread(messagebox.askyesno, title, message)

    def _ensure_preferences_path(self):
        path = self.preferences_var.get().strip()
        if path:
            return Path(path)
        return self.repo_root / "config" / "preferences.example.yaml"

    def _load_current_preference(self):
        pref_path = self._ensure_preferences_path()
        return load_reservation_preference(pref_path)

    def _resolve_id_csv_for_execution(self):
        path = self.id_csv_var.get().strip()
        if not path:
            return None
        if not Path(path).exists():
            raise FileNotFoundError(f"ID CSV が見つかりません: {path}")
        return path

    def _load_last_settings(self):
        local = configparser.ConfigParser()
        if self.local_config_path.exists():
            local.read(self.local_config_path, encoding="utf-8")
        self.id_csv_var.set(
            local.get(
                "GUI",
                "last_id_csv",
                fallback="",
            )
        )
        self.preferences_var.set(
            local.get(
                "GUI",
                "last_preferences",
                fallback=str(self.repo_root / "config" / "preferences.example.yaml"),
            )
        )
        self.window_geometry = local.get(
            "GUI",
            "last_window_geometry",
            fallback=self.window_geometry,
        )

    def _save_last_settings(self):
        local = configparser.ConfigParser()
        if self.local_config_path.exists():
            local.read(self.local_config_path, encoding="utf-8")
        if not local.has_section("GUI"):
            local.add_section("GUI")
        local.set("GUI", "last_id_csv", self.id_csv_var.get().strip())
        local.set("GUI", "last_preferences", self.preferences_var.get().strip())
        local.set("GUI", "last_window_geometry", self.master.geometry())
        with self.local_config_path.open("w", encoding="utf-8") as fh:
            local.write(fh)

    def on_close(self):
        self._save_last_settings()
        self.master.destroy()

    def run_lottery_entry_workflow(self):
        self._run_in_background("抽選申込みワークフロー", self._execute_lottery_entry_workflow)

    def _execute_lottery_entry_workflow(self):
        preference = self._load_current_preference()
        result = self.lottery_entry_workflow_service.run(
            preference=preference,
            id_csv=self._resolve_id_csv_for_execution(),
            max_select=preference.lottery_max_entries_per_account or 2,
            display_result_callback=self._log_lottery_entry_preview,
            confirm_submit_callback=self._confirm_submission_from_gui,
            output_dir=get_output_base_path() / "lottery_automation",
        )
        self.lottery_entry_workflow_service.print_result(result)
        result_path = self.lottery_entry_workflow_service.save_result(
            result,
            get_output_base_path() / "lottery_automation",
        )
        self._log_message(f"ワークフロー結果を保存しました: {result_path}")

    def _confirm_submission_from_gui(self, result, entry_result):
        preference = self._load_current_preference()
        if preference.lottery_dry_run:
            self._log_message("ドライランが有効なため、送信は行いません。")
            return "dry-run"
        slot = (entry_result or {}).get("slot", {})
        if not slot:
            return ""
        account_line = result.get("masked_user_id", "")
        if result.get("account_label"):
            account_line = f"{account_line} ({result.get('account_label')})"
        message_lines = [
            "今回申し込む枠:",
            "",
            f"ID / アカウント: {account_line}",
            f"申込み番号: {entry_result.get('apply_label', '')}",
            f"日付: {slot.get('date', '')}",
            f"曜日: {slot.get('weekday', '')}",
            f"時間帯: {slot.get('time_range', '')}",
            f"施設名: {slot.get('facility', '')}",
            f"現在申込数: {slot.get('current_entry_count', '')}",
            "",
        ]
        message_lines.append("送信しますか？")
        answer = self._threadsafe_ask_yes_no(
            "抽選申込み確認",
            "\n".join(message_lines),
        )
        return "yes" if answer else ""

    def _log_lottery_entry_preview(self, account_result):
        self._log_message(
            "アカウント={account} 状態={status} 取得枠数={count}".format(
                account=account_result.get("masked_user_id"),
                status=account_result.get("status"),
                count=len(account_result.get("collected_slots", [])),
            )
        )
        for slot in account_result.get("planned_slots", []):
            self._log_message(
                "申込み予定: {date} {time_range} 施設={facility} 現在申込数={applied}".format(
                    date=slot.get("date"),
                    time_range=slot.get("time_range"),
                    facility=" ".join(
                        part
                        for part in (
                            slot.get("park_name", ""),
                            slot.get("facility_name", ""),
                        )
                        if part
                    ).strip(),
                    applied=slot.get("applied_count"),
                )
            )
        for warning in account_result.get("missing_slots", []):
            self._log_message(
                "警告: {date} {time_range} 施設={facility} {warning_text}".format(
                    date=warning.get("date"),
                    time_range=warning.get("time_range"),
                    facility=warning.get("facility"),
                    warning_text=warning.get("warning"),
                ),
                level="WARNING",
            )

    def run_lottery_result_workflow(self):
        self._run_in_background("抽選結果確認ワークフロー", self._execute_lottery_result_workflow)

    def _execute_lottery_result_workflow(self):
        result = self.lottery_result_workflow_service.run(
            id_csv=self._resolve_id_csv_for_execution()
        )
        self.lottery_result_workflow_service.print_result(result)
        output_dir = get_output_base_path() / "lottery_automation"
        json_path, csv_path = self.lottery_result_workflow_service.save_result(
            result,
            output_dir,
        )
        self._log_message(f"結果JSONを保存しました: {json_path}")
        self._log_message(f"結果CSVを保存しました: {csv_path}")

    def run_reservation_confirmation_workflow(self):
        self._run_in_background(
            "予約確定補助ワークフロー",
            self._execute_reservation_confirmation_workflow,
        )

    def _execute_reservation_confirmation_workflow(self):
        result = self.reservation_confirmation_workflow_service.run(
            id_csv=self._resolve_id_csv_for_execution(),
            select_entries_callback=self._select_entries_from_gui,
            confirm_callback=self._confirm_reservation_from_gui,
        )
        self.reservation_confirmation_workflow_service.print_result(result)
        output_dir = get_output_base_path() / "lottery_automation"
        json_path = self.reservation_confirmation_workflow_service.save_result(
            result,
            output_dir,
        )
        self._log_message(f"予約確定補助の結果を保存しました: {json_path}")

    def _select_entries_from_gui(self, won_entries):
        return self._call_on_main_thread(self._open_entry_selection_dialog, won_entries)

    def _open_entry_selection_dialog(self, won_entries):
        dialog = tk.Toplevel(self.master)
        dialog.title("予約確定対象の選択")
        dialog.geometry("720x360")
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(0, weight=1)

        tree = ttk.Treeview(
            dialog,
            columns=("account", "date", "time", "facility"),
            show="headings",
            selectmode="extended",
        )
        for key, title, width in (
            ("account", "アカウント", 120),
            ("date", "日付", 120),
            ("time", "時間帯", 120),
            ("facility", "施設", 280),
        ):
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor=tk.W)
        for index, row in enumerate(won_entries):
            account_part = row.get("account")
            if row.get("account_label"):
                account_part = f"{account_part} ({row.get('account_label')})"
            tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(
                    account_part,
                    row.get("date"),
                    row.get("time_range"),
                    row.get("facility"),
                ),
            )
        tree.grid(row=0, column=0, columnspan=2, sticky=tk.NSEW, padx=12, pady=12)

        selected_indices = []

        def on_ok():
            selected_indices.extend(sorted(int(item) for item in tree.selection()))
            dialog.destroy()

        ttk.Button(dialog, text="選択", command=on_ok).grid(
            row=1, column=0, sticky=tk.W, padx=12, pady=(0, 12)
        )
        ttk.Button(dialog, text="キャンセル", command=dialog.destroy).grid(
            row=1, column=1, sticky=tk.E, padx=12, pady=(0, 12)
        )

        self.master.wait_window(dialog)
        return selected_indices

    def _confirm_reservation_from_gui(self, result):
        count = len(result.get("selected_accounts", []))
        if count <= 0:
            return ""
        answer = self._threadsafe_ask_yes_no(
            "予約確定確認",
            f"{count} アカウントの予約確定を実行しますか？",
        )
        return "yes" if answer else ""

    def open_settings_dialog(self):
        pref_path = self._ensure_preferences_path()
        existing_data = {}
        if pref_path.exists():
            try:
                existing_data = load_preferences_data(pref_path)
            except Exception:
                existing_data = {}
        lottery_data = existing_data.get("lottery", {})
        if not isinstance(lottery_data, dict):
            lottery_data = {}

        dialog = tk.Toplevel(self.master)
        dialog.title("設定")
        dialog.geometry("920x760")
        dialog.minsize(900, 720)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(3, weight=1)
        dialog.rowconfigure(4, weight=1)

        id_csv_var = tk.StringVar(value=self.id_csv_var.get())
        preferences_var = tk.StringVar(value=str(pref_path))
        weekdays_var = tk.StringVar(
            value=", ".join(lottery_data.get("target_weekdays", ["土"]))
        )
        search_weeks_var = tk.IntVar(value=int(lottery_data.get("search_weeks", 8)))
        max_entries_var = tk.IntVar(
            value=int(lottery_data.get("max_entries_per_account", 2))
        )
        dry_run_var = tk.BooleanVar(value=bool(lottery_data.get("dry_run", True)))
        manual_final_submit_var = tk.BooleanVar(
            value=bool(lottery_data.get("manual_final_submit", False))
        )
        manual_preconfirm_submit_var = tk.BooleanVar(
            value=bool(lottery_data.get("manual_preconfirm_submit", False))
        )
        reuse_browser_session_var = tk.BooleanVar(
            value=bool(lottery_data.get("reuse_browser_session", False))
        )

        file_frame = ttk.LabelFrame(dialog, text="ファイル", padding=12)
        file_frame.grid(row=0, column=0, sticky=tk.EW, padx=12, pady=(12, 8))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="ID CSVファイル").grid(
            row=0, column=0, sticky=tk.W, pady=4
        )
        ttk.Entry(file_frame, textvariable=id_csv_var).grid(
            row=0, column=1, sticky=tk.EW, padx=8, pady=4
        )
        ttk.Button(
            file_frame,
            text="参照",
            command=lambda: self._browse_into_var(
                id_csv_var,
                [("CSV", "*.csv"), ("すべてのファイル", "*.*")],
            ),
        ).grid(row=0, column=2, pady=4)

        ttk.Label(file_frame, text="設定YAML").grid(
            row=1, column=0, sticky=tk.W, pady=4
        )
        ttk.Entry(file_frame, textvariable=preferences_var).grid(
            row=1, column=1, sticky=tk.EW, padx=8, pady=4
        )
        ttk.Button(
            file_frame,
            text="参照",
            command=lambda: self._browse_into_var(
                preferences_var,
                [("YAML/JSON", "*.yaml *.yml *.json"), ("すべてのファイル", "*.*")],
            ),
        ).grid(row=1, column=2, pady=4)

        lottery_frame = ttk.LabelFrame(dialog, text="抽選設定", padding=12)
        lottery_frame.grid(row=1, column=0, sticky=tk.EW, padx=12, pady=8)
        for index in range(4):
            lottery_frame.columnconfigure(index, weight=1)
        ttk.Label(lottery_frame, text="対象曜日").grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Entry(lottery_frame, textvariable=weekdays_var).grid(
            row=0, column=1, sticky=tk.EW, padx=8
        )
        ttk.Label(lottery_frame, text="探索週数").grid(row=0, column=2, sticky=tk.W)
        ttk.Spinbox(
            lottery_frame,
            from_=1,
            to=12,
            textvariable=search_weeks_var,
            width=6,
        ).grid(row=0, column=3, sticky=tk.W)
        ttk.Label(lottery_frame, text="1IDあたり最大申込み数").grid(
            row=1, column=0, sticky=tk.W, pady=(8, 0)
        )
        ttk.Spinbox(
            lottery_frame,
            from_=1,
            to=2,
            textvariable=max_entries_var,
            width=6,
        ).grid(row=1, column=1, sticky=tk.W, padx=8, pady=(8, 0))
        ttk.Checkbutton(lottery_frame, text="ドライラン", variable=dry_run_var).grid(
            row=1, column=2, columnspan=2, sticky=tk.W, pady=(8, 0)
        )
        ttk.Checkbutton(
            lottery_frame,
            text="最終申込みを手動にする（切り分け用）",
            variable=manual_final_submit_var,
        ).grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(8, 0))
        ttk.Checkbutton(
            lottery_frame,
            text="申込みボタン前を手動にする（切り分け用）",
            variable=manual_preconfirm_submit_var,
        ).grid(row=3, column=0, columnspan=4, sticky=tk.W, pady=(8, 0))
        ttk.Checkbutton(
            lottery_frame,
            text="ブラウザ再利用（アカウント間）",
            variable=reuse_browser_session_var,
        ).grid(row=4, column=0, columnspan=4, sticky=tk.W, pady=(8, 0))

        default_tree = self._build_entry_tree(dialog, "共通申込み枠", row=2)
        overrides_tree = self._build_override_tree(dialog, row=3)

        for entry in lottery_data.get("default_entries", []):
            default_tree.insert(
                "",
                tk.END,
                values=(
                    entry.get("facility", ""),
                    entry.get("date", ""),
                    entry.get("time_range", ""),
                ),
            )

        raw_overrides = lottery_data.get("account_overrides", {})
        if isinstance(raw_overrides, dict):
            for account_id, override_data in raw_overrides.items():
                entries = override_data.get("entries", []) if isinstance(override_data, dict) else []
                for entry in entries:
                    overrides_tree.insert(
                        "",
                        tk.END,
                        values=(
                            account_id,
                            entry.get("facility", ""),
                            entry.get("date", ""),
                            entry.get("time_range", ""),
                        ),
                    )

        self._attach_entry_buttons(
            dialog,
            default_tree,
            row=2,
            add_command=lambda: self._add_default_entry(default_tree),
            edit_command=lambda: self._edit_default_entry(default_tree),
            remove_command=lambda: self._remove_selected_tree_item(default_tree),
        )
        self._attach_entry_buttons(
            dialog,
            overrides_tree,
            row=3,
            add_command=lambda: self._add_override_entry(overrides_tree),
            edit_command=lambda: self._edit_override_entry(overrides_tree),
            remove_command=lambda: self._remove_selected_tree_item(overrides_tree),
        )

        button_frame = ttk.Frame(dialog, padding=12)
        button_frame.grid(row=4, column=0, sticky=tk.EW)
        button_frame.columnconfigure(0, weight=1)

        def on_save():
            updated_data = existing_data if isinstance(existing_data, dict) else {}
            lottery = updated_data.setdefault("lottery", {})
            weekdays = [
                item.strip()
                for item in weekdays_var.get().replace("、", ",").split(",")
                if item.strip()
            ]
            lottery["target_weekdays"] = weekdays or ["土"]
            lottery["search_weeks"] = int(search_weeks_var.get() or 8)
            lottery["max_entries_per_account"] = int(max_entries_var.get() or 2)
            lottery["dry_run"] = bool(dry_run_var.get())
            lottery["manual_final_submit"] = bool(manual_final_submit_var.get())
            lottery["manual_preconfirm_submit"] = bool(manual_preconfirm_submit_var.get())
            lottery["reuse_browser_session"] = bool(reuse_browser_session_var.get())
            lottery["default_entries"] = [
                {
                    "facility": values[0],
                    "date": values[1],
                    "time_range": values[2],
                }
                for values in self._tree_values(default_tree)
            ]

            overrides = {}
            for values in self._tree_values(overrides_tree):
                account_id, facility, date, time_range = values
                overrides.setdefault(account_id, {"entries": []})
                overrides[account_id]["entries"].append(
                    {
                        "facility": facility,
                        "date": date,
                        "time_range": time_range,
                    }
                )
            lottery["account_overrides"] = overrides

            save_preferences_data(preferences_var.get().strip(), updated_data)
            self.id_csv_var.set(id_csv_var.get().strip())
            self.preferences_var.set(preferences_var.get().strip())
            self._save_last_settings()
            self._log_message(f"設定YAMLを保存しました: {preferences_var.get().strip()}")
            dialog.destroy()

        ttk.Button(button_frame, text="保存", command=on_save).grid(
            row=0, column=1, sticky=tk.E, padx=(0, 8)
        )
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).grid(
            row=0, column=2, sticky=tk.E
        )

        self.master.wait_window(dialog)

    def _browse_into_var(self, variable, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            variable.set(path)

    def _build_entry_tree(self, master, title, row):
        frame = ttk.LabelFrame(master, text=title, padding=12)
        frame.grid(row=row, column=0, sticky=tk.NSEW, padx=12, pady=8)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        master.rowconfigure(row, weight=1)
        tree = ttk.Treeview(
            frame,
            columns=("facility", "date", "time_range"),
            show="headings",
            height=5,
        )
        for key, label, width in (
            ("facility", "施設名", 240),
            ("date", "日付", 120),
            ("time_range", "時間帯", 120),
        ):
            tree.heading(key, text=label)
            tree.column(key, width=width, anchor=tk.W)
        tree.grid(row=0, column=0, sticky=tk.NSEW)
        tree._button_parent = frame
        return tree

    def _build_override_tree(self, master, row):
        frame = ttk.LabelFrame(master, text="ID別上書き設定", padding=12)
        frame.grid(row=row, column=0, sticky=tk.NSEW, padx=12, pady=8)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        master.rowconfigure(row, weight=1)
        tree = ttk.Treeview(
            frame,
            columns=("account_id", "facility", "date", "time_range"),
            show="headings",
            height=6,
        )
        for key, label, width in (
            ("account_id", "ID", 120),
            ("facility", "施設名", 220),
            ("date", "日付", 120),
            ("time_range", "時間帯", 120),
        ):
            tree.heading(key, text=label)
            tree.column(key, width=width, anchor=tk.W)
        tree.grid(row=0, column=0, sticky=tk.NSEW)
        tree._button_parent = frame
        return tree

    def _attach_entry_buttons(self, master, tree, row, add_command, edit_command, remove_command):
        button_parent = getattr(tree, "_button_parent", master)
        button_frame = ttk.Frame(button_parent)
        button_frame.grid(row=1, column=0, sticky=tk.E, pady=(8, 0))
        ttk.Button(button_frame, text="追加", command=add_command).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(button_frame, text="編集", command=edit_command).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Button(button_frame, text="削除", command=remove_command).pack(
            side=tk.LEFT, padx=4
        )

    def _add_default_entry(self, tree):
        values = self._prompt_entry_values()
        if values:
            tree.insert("", tk.END, values=values)

    def _edit_default_entry(self, tree):
        selected = tree.selection()
        if not selected:
            return
        current = tree.item(selected[0], "values")
        values = self._prompt_entry_values(current=current)
        if values:
            tree.item(selected[0], values=values)

    def _add_override_entry(self, tree):
        values = self._prompt_override_values()
        if values:
            tree.insert("", tk.END, values=values)

    def _edit_override_entry(self, tree):
        selected = tree.selection()
        if not selected:
            return
        current = tree.item(selected[0], "values")
        values = self._prompt_override_values(current=current)
        if values:
            tree.item(selected[0], values=values)

    def _remove_selected_tree_item(self, tree):
        for item in tree.selection():
            tree.delete(item)

    def _prompt_entry_values(self, current=None):
        return self._show_entry_editor_dialog(current=current)

    def _prompt_override_values(self, current=None):
        return self._show_entry_editor_dialog(current=current, include_account=True)

    def _show_entry_editor_dialog(self, current=None, include_account=False):
        dialog = tk.Toplevel(self.master)
        dialog.title("申込み枠の編集" if not include_account else "ID別申込み枠の編集")
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.columnconfigure(1, weight=1)

        current = current or ()
        row = 0
        account_var = tk.StringVar(value=current[0] if include_account and current else "")
        if include_account:
            ttk.Label(dialog, text="ID").grid(
                row=row, column=0, sticky=tk.W, padx=12, pady=(12, 6)
            )
            ttk.Entry(dialog, textvariable=account_var, width=32).grid(
                row=row, column=1, sticky=tk.EW, padx=(0, 12), pady=(12, 6)
            )
            row += 1

        offset = 1 if include_account else 0
        facility_var = tk.StringVar(value=current[offset] if len(current) > offset else "")
        date_var = tk.StringVar(value=current[offset + 1] if len(current) > offset + 1 else "")
        time_range_var = tk.StringVar(
            value=current[offset + 2] if len(current) > offset + 2 else ""
        )

        ttk.Label(dialog, text="施設名").grid(
            row=row, column=0, sticky=tk.W, padx=12, pady=6
        )
        ttk.Entry(dialog, textvariable=facility_var, width=40).grid(
            row=row, column=1, sticky=tk.EW, padx=(0, 12), pady=6
        )
        row += 1

        ttk.Label(dialog, text="日付").grid(
            row=row, column=0, sticky=tk.W, padx=12, pady=6
        )
        ttk.Entry(dialog, textvariable=date_var, width=24).grid(
            row=row, column=1, sticky=tk.EW, padx=(0, 12), pady=6
        )
        row += 1

        ttk.Label(dialog, text="時間帯").grid(
            row=row, column=0, sticky=tk.W, padx=12, pady=6
        )
        ttk.Entry(dialog, textvariable=time_range_var, width=24).grid(
            row=row, column=1, sticky=tk.EW, padx=(0, 12), pady=6
        )
        row += 1

        ttk.Label(
            dialog,
            text="日付は YYYY-MM-DD、時間帯は HH:MM-HH:MM 形式で入力してください。",
        ).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=12, pady=(0, 8))
        row += 1

        result = {"values": None}

        def on_save():
            values = (
                facility_var.get().strip(),
                date_var.get().strip(),
                time_range_var.get().strip(),
            )
            if include_account:
                result["values"] = (account_var.get().strip(),) + values
            else:
                result["values"] = values
            dialog.destroy()

        button_frame = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        button_frame.grid(row=row, column=0, columnspan=2, sticky=tk.EW)
        button_frame.columnconfigure(0, weight=1)
        ttk.Button(button_frame, text="保存", command=on_save).grid(
            row=0, column=1, padx=(0, 8)
        )
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).grid(
            row=0, column=2
        )

        self.master.wait_window(dialog)
        return result["values"]

    def _tree_values(self, tree):
        values = []
        for item in tree.get_children():
            values.append(tuple(tree.item(item, "values")))
        return values

    # Compatibility helpers / legacy wrappers
    def _start_driver(self):
        if getattr(self, "driver", None) is None:
            self.driver = self.browser_session.create_driver()

    def _get_wait(self, timeout=10):
        return self.browser_session.get_wait(self.driver, timeout)

    def _login(self, user_id, password):
        self._start_driver()
        return self.login_service.login(self.driver, user_id, password)

    def _logout(self):
        try:
            self.navigation_service.logout(self.driver)
        except Exception:
            pass

    def _detect_captcha(self):
        return self.login_service.detect_captcha(self.driver)

    def _wait_for_captcha_solve(self, timeout_minutes=5):
        return self.login_service.wait_for_manual_captcha(
            self.driver, timeout_minutes=timeout_minutes
        )

    def _navigate_to_lottery_entry(self):
        self.navigation_service.go_to_lottery_entry(self.driver)
        self.navigation_service.select_lottery_tennis_park(self.driver)

    def collect_all_available_slots(self, weeks_limit=8, only_weekday=None):
        return self.availability_service.collect_all_available_slots(
            self.driver,
            weeks_limit=weeks_limit,
            only_weekday=only_weekday,
        )

    def semiauto_reserv(self, id_dict={}):
        if not id_dict:
            return
        self.lottery_service.semiauto_reserv(id_dict)

    def full_auto_reserv(self, id_dict={}, max_attempts=2):
        if not id_dict:
            return
        self.lottery_service.full_auto_reserv(id_dict, max_attempts=max_attempts)

    def auto_select_and_submit_slots(self, selected_slots, submit=True, wait_alert_seconds=10):
        return self.lottery_service.auto_select_and_submit_slots(
            self.driver,
            selected_slots,
            submit=submit,
            wait_alert_seconds=wait_alert_seconds,
        )

    def check_lottery(self, id_dict={}, output_csv_path=""):
        if not id_dict:
            return {}
        return self.lottery_service.check_lottery(id_dict, output_csv_path)

    def check_result(self, id_dict={}, output_csv_path=""):
        if not id_dict:
            return {}
        return self.lottery_service.check_result(id_dict, output_csv_path)

    def determine_reserv(self, input_csv_path="", output_csv_path=""):
        return self.reservation_service.determine_reserv(
            input_csv_path,
            output_csv_path,
        )

    def check_reserv(self, id_dict={}, output_csv_path=""):
        if not id_dict:
            return {}
        return self.reservation_service.check_reserv(id_dict, output_csv_path)

    def check_court(self, month):
        return self.availability_service.check_court(month)


def main():
    root = tk.Tk()
    cr = Court_Reserv(master=root)
    cr.mainloop()


if __name__ == "__main__":
    main()
