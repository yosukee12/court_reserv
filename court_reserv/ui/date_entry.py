"""Dependency-free date entry with an inline calendar."""

import calendar
from datetime import date
import tkinter as tk
from tkinter import ttk


def parse_entry_date(value):
    """Validate at save time, without interrupting text editing."""
    value = value.strip()
    normalized = date.fromisoformat(value).isoformat()
    if value != normalized:
        raise ValueError("Expected YYYY-MM-DD")
    return normalized


def adjacent_month(month, step):
    index = month.year * 12 + month.month - 1 + step
    year, month_index = divmod(index, 12)
    return date(year, month_index + 1, 1)


class DateEntry(ttk.Frame):
    def __init__(self, master, variable):
        super().__init__(master)
        self.variable = variable
        try:
            self.month = date.fromisoformat(parse_entry_date(variable.get())).replace(day=1)
        except ValueError:
            self.month = date.today().replace(day=1)
        self.columnconfigure(0, weight=1)
        self.entry = ttk.Entry(self, textvariable=variable, width=16)
        self.entry.grid(row=0, column=0, sticky=tk.EW)
        self.entry.bind("<Control-a>", self._select_all)
        if self.tk.call("tk", "windowingsystem") == "aqua":
            self.entry.bind("<Command-a>", self._select_all)
        ttk.Button(self, text="クリア", command=self._clear).grid(row=0, column=1)
        self.calendar = ttk.Frame(self)
        self.calendar.grid(row=1, column=0, columnspan=2, pady=(6, 0))
        self._render()

    def _select_all(self, event=None):
        self.entry.selection_range(0, tk.END)
        self.entry.icursor(tk.END)
        return "break"

    def _clear(self):
        self.variable.set("")
        self.entry.focus_set()

    def _move(self, step):
        try:
            self.month = adjacent_month(self.month, step)
        except ValueError:
            return
        self._render()

    def _choose(self, day):
        self.variable.set(self.month.replace(day=day).isoformat())
        self._render()
        self.entry.focus_set()

    def _render(self):
        for child in self.calendar.winfo_children():
            child.destroy()
        ttk.Button(self.calendar, text="前月", command=lambda: self._move(-1)).grid(
            row=0, column=0, columnspan=2
        )
        ttk.Label(self.calendar, text=f"{self.month.year}年{self.month.month}月").grid(
            row=0, column=2, columnspan=3
        )
        ttk.Button(self.calendar, text="翌月", command=lambda: self._move(1)).grid(
            row=0, column=5, columnspan=2
        )
        for column, label in enumerate("月火水木金土日"):
            ttk.Label(self.calendar, text=label, anchor=tk.CENTER).grid(
                row=1, column=column, sticky=tk.EW
            )
        for row, week in enumerate(calendar.Calendar().monthdayscalendar(
            self.month.year, self.month.month
        ), start=2):
            for column, day in enumerate(week):
                if day:
                    button = ttk.Button(
                        self.calendar, text=str(day), width=3,
                        command=lambda day=day: self._choose(day),
                    )
                    button.grid(row=row, column=column)
                    if self.variable.get() == self.month.replace(day=day).isoformat():
                        button.state(["pressed"])
