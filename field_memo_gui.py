import os
import re
import glob
import shutil
import tkinter as tk
from tkinter import filedialog

# ─────────────────────────────────────────────
#  DEFAULT PATHS
# ─────────────────────────────────────────────
DEFAULT_FIX_IMAGE    = 'C:/Users/Public/Documents/4D Nav/NavView/OBN013626_Eng10/Local/Online/Data/Node Dashboard/Fix Images/*.png'
DEFAULT_UHD333_IMAGE = 'C:/Users/Public/Documents/4D Nav/NavView/OBN013626_Eng10/Shared/Data/Clips/UHD333 Color/*.png'
DEFAULT_UHD334_IMAGE = 'C:/Users/Public/Documents/4D Nav/NavView/OBN013626_Eng10/Shared/Data/Clips/UHD334 Color/*.png'
DEFAULT_DEST_FOLDER  = 'Z:/Projects/OBN013626_SLB_USA_Engagement10/04_SURVEY/01.Data/01.Node_Pictures'

EVENT_LABEL = 'Position deviation'   # folder name middle segment: FM-055-Position deviation[-reason]

BG      = '#0d0f14'
PANEL   = '#13161e'
BORDER  = '#1e2330'
FG      = '#c8cfe0'
FG_DIM  = '#4a5270'
GREEN   = '#00e676'
AMBER   = '#ffb300'
RED     = '#ff1744'
FM      = ('Courier New', 9)
FB      = ('Courier New', 11, 'bold')
FT      = ('Courier New', 12, 'bold')

FLASH_INTERVAL_MS = 500  # blink period while a button is armed/waiting for its second click


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Field Memo — Position Deviation Photo Logger')
        self.configure(bg=BG)
        self.minsize(560, 380)
        self.geometry('720x480')
        # Scalable window — unlike the old fixed-size tool, this one resizes freely.
        self.resizable(True, True)

        self.fix_image_var    = tk.StringVar(value=DEFAULT_FIX_IMAGE)
        self.uhd333_image_var = tk.StringVar(value=DEFAULT_UHD333_IMAGE)
        self.uhd334_image_var = tk.StringVar(value=DEFAULT_UHD334_IMAGE)
        self.dest_folder_var  = tk.StringVar(value=DEFAULT_DEST_FOLDER)

        # ── event state machine ──
        # active_uhd: None | "333" | "334" — which button (if any) is armed,
        # waiting for its second click. Only one event can be in progress at
        # a time; the other UHD button is disabled while one is active.
        self.active_uhd     = None
        self.current_folder = None
        self.reason_win     = None
        self._flash_job      = None
        self._flash_on       = False

        self._build_ui()

    # ══════════════════════════════════════════
    #  UI
    # ══════════════════════════════════════════
    def _build_ui(self):
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ── title ──
        top = tk.Frame(self, bg=BG)
        top.grid(row=0, column=0, sticky='ew', padx=10, pady=(10, 4))
        tk.Label(top, text='⬡ FIELD MEMO', font=FT, bg=BG, fg=GREEN).pack(side='left')

        # ── paths ──
        pf = tk.LabelFrame(self, text=' PATHS ', font=FM, bg=BG, fg=FG_DIM,
                            bd=1, relief='flat', highlightthickness=1,
                            highlightbackground=BORDER)
        pf.grid(row=1, column=0, sticky='ew', padx=10, pady=(2, 6))
        pf.columnconfigure(1, weight=1)

        entries = [
            ('Fix Images', self.fix_image_var, False),
            ('UHD333',     self.uhd333_image_var, False),
            ('UHD334',     self.uhd334_image_var, False),
            ('Destination', self.dest_folder_var, True),
        ]
        for i, (lbl, var, is_dir) in enumerate(entries):
            tk.Label(pf, text=lbl, font=FM, bg=BG, fg=FG_DIM, anchor='w', width=11
                      ).grid(row=i, column=0, sticky='w', padx=(8, 2), pady=3)
            tk.Entry(pf, textvariable=var, font=FM, bg=PANEL, fg=FG,
                      insertbackground=FG, relief='flat', bd=0,
                      highlightthickness=1, highlightcolor=GREEN,
                      highlightbackground=BORDER
                      ).grid(row=i, column=1, sticky='ew', pady=3)
            tk.Button(pf, text='…', font=FM, bg=BORDER, fg=FG, relief='flat',
                      bd=0, cursor='hand2', padx=6,
                      command=lambda v=var, d=is_dir: self._browse(v, d)
                      ).grid(row=i, column=2, padx=(4, 8), pady=3)

        # ── log ──
        lf = tk.Frame(self, bg=BG)
        lf.grid(row=2, column=0, sticky='nsew', padx=10, pady=(2, 6))
        lf.grid_rowconfigure(0, weight=1)
        lf.grid_columnconfigure(0, weight=1)

        self.log_box = tk.Text(lf, font=FM, bg=PANEL, fg=FG, insertbackground=FG,
                                relief='flat', bd=0, highlightthickness=1,
                                highlightbackground=BORDER, state='disabled', wrap='word')
        self.log_box.grid(row=0, column=0, sticky='nsew')
        sb = tk.Scrollbar(lf, command=self.log_box.yview, bg=BORDER, troughcolor=BG, relief='flat')
        sb.grid(row=0, column=1, sticky='ns')
        self.log_box.configure(yscrollcommand=sb.set)
        self.log_box.tag_config('ok',  foreground=GREEN)
        self.log_box.tag_config('err', foreground=RED)
        self.log_box.tag_config('w',   foreground=AMBER)

        # ── UHD action buttons ──
        bot = tk.Frame(self, bg=BG)
        bot.grid(row=3, column=0, sticky='ew', padx=10, pady=(2, 10))
        bot.columnconfigure(0, weight=1)
        bot.columnconfigure(1, weight=1)

        self.btn_uhd333 = tk.Button(
            bot, text='UHD333', font=FB, bg=PANEL, fg=FG,
            activebackground=BORDER, relief='flat', bd=0, cursor='hand2',
            pady=10, command=lambda: self._on_uhd_click('333'))
        self.btn_uhd333.grid(row=0, column=0, sticky='ew', padx=(0, 4))

        self.btn_uhd334 = tk.Button(
            bot, text='UHD334', font=FB, bg=PANEL, fg=FG,
            activebackground=BORDER, relief='flat', bd=0, cursor='hand2',
            pady=10, command=lambda: self._on_uhd_click('334'))
        self.btn_uhd334.grid(row=0, column=1, sticky='ew', padx=(4, 0))

        self._buttons = {'333': self.btn_uhd333, '334': self.btn_uhd334}

        tk.Button(self, text='CLR LOG', font=FM, bg=BORDER, fg=FG_DIM,
                  relief='flat', bd=0, cursor='hand2', padx=8, pady=3,
                  command=self._clear_log).grid(row=4, column=0, sticky='e', padx=10, pady=(0, 8))

    # ══════════════════════════════════════════
    #  LOG
    # ══════════════════════════════════════════
    def _log(self, msg, tag='ok'):
        self.log_box.configure(state='normal')
        self.log_box.insert('end', f'{msg}\n', tag)
        self.log_box.see('end')
        self.log_box.configure(state='disabled')

    def _clear_log(self):
        self.log_box.configure(state='normal')
        self.log_box.delete('1.0', 'end')
        self.log_box.configure(state='disabled')

    # ══════════════════════════════════════════
    #  BROWSE
    # ══════════════════════════════════════════
    def _browse(self, var, is_dir):
        if is_dir:
            p = filedialog.askdirectory(initialdir=var.get())
            if p:
                var.set(p.replace('/', '\\'))
        else:
            current = var.get()
            p = filedialog.askdirectory(initialdir=os.path.dirname(current))
            if p:
                var.set(p.replace('/', '\\') + '\\' + os.path.basename(current))

    # ══════════════════════════════════════════
    #  HELPERS
    # ══════════════════════════════════════════
    def _latest_photo(self, pattern):
        files = glob.glob(pattern)
        if not files:
            return None
        return max(files, key=os.path.getctime)

    def _next_event_number(self, dest_folder):
        """Scan dest_folder for existing 'FM-###-...' folders and return the
        next sequential number, zero-padded to at least 3 digits."""
        highest = 0
        if os.path.isdir(dest_folder):
            for name in os.listdir(dest_folder):
                m = re.match(r'FM-(\d+)-', name)
                if m:
                    highest = max(highest, int(m.group(1)))
        return highest + 1

    def _copy_latest(self, pattern, dest_folder, label):
        src = self._latest_photo(pattern)
        if not src:
            self._log(f'✗ No photos found for {label} ({pattern})', 'err')
            return False
        try:
            shutil.copy2(src, os.path.join(dest_folder, os.path.basename(src)))
            self._log(f'✓ Copied {label}: {os.path.basename(src)}', 'ok')
            return True
        except Exception as e:
            self._log(f'✗ Copy failed for {label}: {e}', 'err')
            return False

    # ══════════════════════════════════════════
    #  UHD BUTTON CLICK — 2-click event state machine
    # ══════════════════════════════════════════
    def _on_uhd_click(self, which):
        if self.active_uhd is None:
            self._start_event(which)
        elif self.active_uhd == which:
            self._complete_event(which)
        # else: the other button is disabled while one is active, so this
        # branch shouldn't normally be reachable — ignored defensively.

    def _uhd_pattern(self, which):
        return self.uhd333_image_var.get() if which == '333' else self.uhd334_image_var.get()

    def _start_event(self, which):
        dest_root = self.dest_folder_var.get()
        try:
            os.makedirs(dest_root, exist_ok=True)
        except Exception as e:
            self._log(f'✗ Cannot access destination folder: {e}', 'err')
            return

        num = self._next_event_number(dest_root)
        folder_name = f'FM-{num:03d}-{EVENT_LABEL}'
        folder_path = os.path.join(dest_root, folder_name)
        try:
            os.makedirs(folder_path, exist_ok=True)
        except Exception as e:
            self._log(f'✗ Could not create {folder_name}: {e}', 'err')
            return

        self._log(f'▶ Started event {folder_name} (UHD{which})', 'ok')
        self.active_uhd = which
        self.current_folder = folder_path

        self._copy_latest(self._uhd_pattern(which), folder_path, f'UHD{which} #1')

        self._disable_other_button(which)
        self._start_flash(which)
        self._open_reason_prompt(which)

    def _complete_event(self, which):
        folder_path = self.current_folder
        self._copy_latest(self._uhd_pattern(which), folder_path, f'UHD{which} #2')
        self._copy_latest(self.fix_image_var.get(), folder_path, 'Fix Image')

        if self.reason_win is not None:
            try:
                self.reason_win.destroy()
            except Exception:
                pass
            self.reason_win = None

        self._log(f'■ Completed event in {os.path.basename(folder_path)}', 'ok')
        self._stop_flash(which)
        self._enable_both_buttons()
        self.active_uhd = None
        self.current_folder = None

    # ══════════════════════════════════════════
    #  BUTTON ENABLE / DISABLE / FLASH
    # ══════════════════════════════════════════
    def _disable_other_button(self, which):
        other = '334' if which == '333' else '333'
        self._buttons[other].configure(state='disabled')

    def _enable_both_buttons(self):
        for btn in self._buttons.values():
            btn.configure(state='normal', bg=PANEL, fg=FG)

    def _start_flash(self, which):
        self._flash_on = False
        self._do_flash(which)

    def _do_flash(self, which):
        btn = self._buttons[which]
        self._flash_on = not self._flash_on
        if self._flash_on:
            btn.configure(bg=GREEN, fg='#000')
        else:
            btn.configure(bg=PANEL, fg=FG)
        self._flash_job = self.after(FLASH_INTERVAL_MS, lambda: self._do_flash(which))

    def _stop_flash(self, which):
        if self._flash_job is not None:
            self.after_cancel(self._flash_job)
            self._flash_job = None
        self._buttons[which].configure(bg=PANEL, fg=FG)

    # ══════════════════════════════════════════
    #  REASON PROMPT — non-modal, doesn't block clicking the flashing button
    # ══════════════════════════════════════════
    def _open_reason_prompt(self, which):
        win = tk.Toplevel(self)
        win.title('Reason')
        win.configure(bg=BG)
        win.geometry('300x120')
        win.transient(self)
        # Deliberately no grab_set() — the flashing button must stay
        # clickable while this is open (per the 2-click workflow).

        tk.Label(win, text=f'Reason for FM event (UHD{which}):', font=FM,
                 bg=BG, fg=FG_DIM).pack(padx=12, pady=(14, 4), anchor='w')

        entry_var = tk.StringVar()
        entry = tk.Entry(win, textvariable=entry_var, font=FM, bg=PANEL, fg=FG,
                          insertbackground=FG, relief='flat', bd=0,
                          highlightthickness=1, highlightcolor=GREEN,
                          highlightbackground=BORDER)
        entry.pack(fill='x', padx=12, pady=4)
        entry.focus_set()

        def submit(event=None):
            reason = entry_var.get().strip().lower()
            if reason:
                self._apply_reason(which, reason)
            win.destroy()
            if self.reason_win is win:
                self.reason_win = None

        entry.bind('<Return>', submit)

        btn_row = tk.Frame(win, bg=BG)
        btn_row.pack(fill='x', padx=12, pady=(8, 10))
        tk.Button(btn_row, text='OK', font=FM, bg=GREEN, fg='#000', relief='flat',
                  bd=0, cursor='hand2', padx=10, command=submit).pack(side='right')

        def on_close():
            win.destroy()
            if self.reason_win is win:
                self.reason_win = None
        win.protocol('WM_DELETE_WINDOW', on_close)

        self.reason_win = win

    def _apply_reason(self, which, reason):
        # Only meaningful if this reason answer still belongs to the
        # currently active event (it may have already been completed).
        if self.active_uhd != which or self.current_folder is None:
            self._log(f'(Reason "{reason}" ignored — event already finished)', 'w')
            return
        old_path = self.current_folder
        new_name = f'{os.path.basename(old_path)}-{reason}'
        new_path = os.path.join(os.path.dirname(old_path), new_name)
        try:
            os.rename(old_path, new_path)
            self.current_folder = new_path
            self._log(f'✓ Renamed to {new_name}', 'ok')
        except Exception as e:
            self._log(f'✗ Could not rename folder: {e}', 'err')


if __name__ == '__main__':
    app = App()
    app.mainloop()
