# -*- coding: utf-8 -*-
import tkinter as tk

from court_reserv.court_reserv import Court_Reserv


def main():
    """GUI entrypoint for `python -m court_reserv.ui.app`."""
    root = tk.Tk()
    app = Court_Reserv(master=root)
    app.mainloop()


if __name__ == "__main__":
    main()
