#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import getpass
import sys
import tkinter as tk
import traceback
from pathlib import Path
import datetime

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from court_reserv.court_reserv import Court_Reserv
from court_reserv.config import get_debug_output_dir, get_default_credentials


def save_debug(app, label='error'):
    out_dir = get_debug_output_dir()
    out_dir.mkdir(parents=True, exist_ok=True)

    # repository-level debug dir (alias, same as out_dir)
    repo_dir = out_dir

    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    saved = []

    if getattr(app, 'driver', None):
        # save main html
        try:
            html = app.driver.page_source
            repo_path = repo_dir / f'{label}_page_{ts}.html'
            with open(repo_path, 'w', encoding='utf-8') as f:
                f.write(html)
            saved.append(str(repo_path))

            print('Saved debug HTML (repo):', repo_path)
            if out_dir != repo_dir:
                # keep legacy/configured location as secondary
                conf_path = out_dir / f'{label}_page_{ts}.html'
                with open(conf_path, 'w', encoding='utf-8') as f:
                    f.write(html)
                saved.append(str(conf_path))
                print('Also saved debug HTML (configured):', conf_path)
        except Exception:
            print('Failed to save page_source')

        # save screenshots
        try:
            # save screenshots into the unified output/debug_pages directory
            repo_png = repo_dir / f'{label}_screenshot_{ts}.png'
            app.driver.save_screenshot(str(repo_png))
            saved.append(str(repo_png))

            print('Saved screenshot (repo):', repo_png)
            if out_dir != repo_dir:
                conf_png = out_dir / f'{label}_screenshot_{ts}.png'
                app.driver.save_screenshot(str(conf_png))
                saved.append(str(conf_png))
                print('Also saved screenshot (configured):', conf_png)
        except Exception:
            print('Failed to save screenshot')

    return saved


def main():
    root = tk.Tk()
    root.withdraw()
    app = Court_Reserv(master=root)

    default_uid, default_pwd = get_default_credentials()

    while True:
        uid_prompt = 'Login userId'
        if default_uid:
            uid_prompt += ' [press Enter to use configured default]'
        uid = input(f'{uid_prompt}: ').strip()
        if not uid and default_uid:
            uid = default_uid
            print('Using configured default user id')

        pwd_prompt = 'Password'
        if default_pwd:
            pwd_prompt += ' (press Enter to use configured default)'
        pwd = getpass.getpass(f'{pwd_prompt}: ')
        if not pwd and default_pwd:
            pwd = default_pwd
            print('Using configured default password')

        try:
            ok = app._login(uid, pwd)
            if not ok:
                print('Login failed or expired')
                retry = input('Retry login? (y/N): ').strip().lower()
                if retry == 'y':
                    continue
                else:
                    return
            # navigate and collect (only Saturdays)
            app._navigate_to_lottery_entry()
            slots = app.collect_all_available_slots(weeks_limit=8, only_weekday=5)
            selected = app.prompt_user_to_select_slots(slots, max_select=2)
            print('Final selected:', selected)
            if selected:
                do_apply = input('Attempt auto-apply these slots now? (y/N): ').strip().lower()
                if do_apply == 'y':
                    print('Attempting auto-apply...')
                    res = app.auto_select_and_submit_slots(selected, submit=True)
                    print('Apply results:')
                    for k, v in res.items():
                        print(k, '=>', 'OK' if v else 'NOT_FOUND')
            break

        except Exception as e:
            print('Exception during run:', e)
            traceback.print_exc()
            try:
                save_debug(app, label='exception')
            except Exception:
                pass
            retry = input('An error occurred. Retry from login? (y/N): ').strip().lower()
            if retry == 'y':
                # attempt logout/quit before retry
                try:
                    app._logout()
                    if getattr(app, 'driver', None):
                        app.driver.quit()
                        delattr(app, 'driver')
                except Exception:
                    pass
                continue
            else:
                break
    try:
        app._logout()
        if getattr(app, 'driver', None):
            app.driver.quit()
    except Exception:
        pass


if __name__ == '__main__':
    main()
