import os
import glob
import re
import logging
import threading
import tkinter as tk
from tkinter import filedialog
from shutil import copy2
from datetime import datetime, timedelta
from time import sleep

# ─────────────────────────────────────────────
#  DEFAULT PATHS
# ─────────────────────────────────────────────
DEFAULT_FIX_IMAGE    = 'C:/Users/Public/Documents/4D Nav/NavView/OBN013626_Eng10/Local/Online/Data/Node Dashboard/Fix Images/*.png'
DEFAULT_UHD333_IMAGE = 'C:/Users/Public/Documents/4D Nav/NavView/OBN013626_Eng10/Shared/Data/Clips/UHD333 Color/*.png'
DEFAULT_UHD334_IMAGE = 'C:/Users/Public/Documents/4D Nav/NavView/OBN013626_Eng10/Shared/Data/Clips/UHD334 Color/*.png'
DEFAULT_RL_FOLDER    = 'Z:/Projects/OBN013626_SLB_USA_Engagement10/04_SURVEY/01.Data/01.Node_Pictures'

BG      = '#0d0f14'
PANEL   = '#13161e'
BORDER  = '#1e2330'
FG      = '#c8cfe0'
FG_DIM  = '#4a5270'
GREEN   = '#00e676'
AMBER   = '#ffb300'
RED     = '#ff1744'
FM      = ('Courier New', 8)
FB      = ('Courier New', 10, 'bold')
FT      = ('Courier New', 10, 'bold')


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('IFD · Node Pic Router')
        self.configure(bg=BG)
        self.resizable(False, False)

        self.deploying        = tk.BooleanVar(value=True)
        self.time_dif         = tk.IntVar(value=0)
        self.running          = False
        self._worker          = None
        self._flash_job       = None
        self._transfer_count  = 0

        self.fix_image_var    = tk.StringVar(value=DEFAULT_FIX_IMAGE)
        self.uhd333_image_var = tk.StringVar(value=DEFAULT_UHD333_IMAGE)
        self.uhd334_image_var = tk.StringVar(value=DEFAULT_UHD334_IMAGE)
        self.rl_folder_var    = tk.StringVar(value=DEFAULT_RL_FOLDER)

        self._build_ui()
        self._setup_logging()
        self.protocol('WM_DELETE_WINDOW', self._on_close)

    # ══════════════════════════════════════════
    #  UI
    # ══════════════════════════════════════════
    def _build_ui(self):

        # ── top: title + status dot ──
        top = tk.Frame(self, bg=BG)
        top.pack(fill='x', padx=10, pady=(8, 2))

        tk.Label(top, text='⬡ ISLAND FRONTIER', font=FT,
                 bg=BG, fg=GREEN).pack(side='left')

        # status dot (canvas circle)
        self._canvas = tk.Canvas(top, width=14, height=14, bg=BG,
                                 highlightthickness=0)
        self._canvas.pack(side='right', padx=(4, 0))
        self._dot = self._canvas.create_oval(1, 1, 13, 13,
                                             fill=FG_DIM, outline='')

        self.status_lbl = tk.Label(top, text='IDLE', font=FM,
                                   bg=BG, fg=FG_DIM)
        self.status_lbl.pack(side='right', padx=6)

        # ── mode row ──
        mrow = tk.Frame(self, bg=BG)
        mrow.pack(fill='x', padx=10, pady=(4, 2))

        tk.Label(mrow, text='MODE:', font=FM, bg=BG, fg=FG_DIM).pack(side='left')

        self.btn_deploy = tk.Button(
            mrow, text=' DEPLOY ', font=FB,
            bg=GREEN, fg='#000', activebackground='#00c853',
            relief='flat', bd=0, cursor='hand2',
            command=lambda: self._set_mode(True))
        self.btn_deploy.pack(side='left', padx=(6, 2))

        self.btn_recover = tk.Button(
            mrow, text=' RECOVERY ', font=FB,
            bg=BORDER, fg=FG_DIM, activebackground='#2a3050',
            relief='flat', bd=0, cursor='hand2',
            command=lambda: self._set_mode(False))
        self.btn_recover.pack(side='left', padx=2)

        # UTC offset
        tk.Label(mrow, text='UTC±:', font=FM, bg=BG, fg=FG_DIM).pack(side='right')
        tk.Spinbox(mrow, from_=-12, to=12, textvariable=self.time_dif,
                   width=3, font=FM, bg=PANEL, fg=FG,
                   buttonbackground=BORDER, relief='flat',
                   insertbackground=FG).pack(side='right', padx=(0, 4))

        # ── paths ──
        pf = tk.LabelFrame(self, text=' PATHS ', font=FM,
                           bg=BG, fg=FG_DIM, bd=1, relief='flat',
                           highlightthickness=1, highlightbackground=BORDER)
        pf.pack(fill='x', padx=10, pady=(6, 2))

        entries = [
            ('Fix Images',  self.fix_image_var,    False),
            ('UHD333',      self.uhd333_image_var, False),
            ('UHD334',      self.uhd334_image_var, False),
            ('Destination', self.rl_folder_var,    True),
        ]
        for i, (lbl, var, is_dir) in enumerate(entries):
            tk.Label(pf, text=lbl, font=FM, bg=BG, fg=FG_DIM,
                     anchor='w', width=11).grid(row=i, column=0,
                                                sticky='w', padx=(8, 2), pady=2)
            tk.Entry(pf, textvariable=var, font=FM,
                     bg=PANEL, fg=FG, insertbackground=FG,
                     relief='flat', bd=0,
                     highlightthickness=1, highlightcolor=GREEN,
                     highlightbackground=BORDER, width=55
                     ).grid(row=i, column=1, sticky='ew', pady=2)
            tk.Button(pf, text='…', font=FM, bg=BORDER, fg=FG,
                      relief='flat', bd=0, cursor='hand2', padx=4,
                      command=lambda v=var, d=is_dir: self._browse(v, d)
                      ).grid(row=i, column=2, padx=(4, 8), pady=2)
        pf.columnconfigure(1, weight=1)

        # ── log ──
        lf = tk.Frame(self, bg=BG)
        lf.pack(fill='both', expand=True, padx=10, pady=(6, 2))

        self.log_box = tk.Text(lf, height=7, font=FM,
                               bg=PANEL, fg=FG, insertbackground=FG,
                               relief='flat', bd=0,
                               highlightthickness=1,
                               highlightbackground=BORDER,
                               state='disabled', wrap='none')
        self.log_box.pack(side='left', fill='both', expand=True)
        sb = tk.Scrollbar(lf, command=self.log_box.yview,
                          bg=BORDER, troughcolor=BG, relief='flat')
        sb.pack(side='right', fill='y')
        self.log_box.configure(yscrollcommand=sb.set)
        self.log_box.tag_config('ok',  foreground=GREEN)
        self.log_box.tag_config('err', foreground=RED)
        self.log_box.tag_config('w',   foreground=AMBER)

        # ── bottom ──
        bot = tk.Frame(self, bg=BG)
        bot.pack(fill='x', padx=10, pady=(4, 10))

        self.btn_start = tk.Button(
            bot, text='▶ START', font=FB,
            bg=GREEN, fg='#000', activebackground='#00c853',
            relief='flat', bd=0, cursor='hand2', padx=12, pady=4,
            command=self._toggle)
        self.btn_start.pack(side='left')

        tk.Button(bot, text='CLR', font=FM,
                  bg=BORDER, fg=FG_DIM, relief='flat', bd=0,
                  cursor='hand2', padx=8, pady=4,
                  command=self._clear_log).pack(side='left', padx=6)

        self.counter_lbl = tk.Label(bot, text='Transfers: 0',
                                    font=FM, bg=BG, fg=FG_DIM)
        self.counter_lbl.pack(side='right')

    # ══════════════════════════════════════════
    #  LOGGING
    # ══════════════════════════════════════════
    def _setup_logging(self):
        self._logger = logging.getLogger('IFD')
        self._logger.setLevel(logging.INFO)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(script_dir,
                                'debug' + datetime.now().strftime('%d%m%Y') + '.log')
        fh = logging.FileHandler(log_path, mode='a')
        fh.setFormatter(logging.Formatter('%(asctime)s %(message)s'))
        self._logger.addHandler(fh)

    def _log(self, msg, tag='ok'):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_box.configure(state='normal')
        self.log_box.insert('end', f'[{ts}] {msg}\n', tag)
        self.log_box.see('end')
        self.log_box.configure(state='disabled')
        self._logger.info(msg)

    def _clear_log(self):
        self.log_box.configure(state='normal')
        self.log_box.delete('1.0', 'end')
        self.log_box.configure(state='disabled')

    # ══════════════════════════════════════════
    #  MODE
    # ══════════════════════════════════════════
    def _set_mode(self, deploy: bool):
        self.deploying.set(deploy)
        if deploy:
            self.btn_deploy.configure(bg=GREEN, fg='#000')
            self.btn_recover.configure(bg=BORDER, fg=FG_DIM)
            self._log('Mode → DEPLOY', 'ok')
        else:
            self.btn_recover.configure(bg=AMBER, fg='#000')
            self.btn_deploy.configure(bg=BORDER, fg=FG_DIM)
            self._log('Mode → RECOVERY', 'w')

    # ══════════════════════════════════════════
    #  BROWSE
    # ══════════════════════════════════════════
    def _browse(self, var, is_dir):
        if is_dir:
            p = filedialog.askdirectory()
            if p:
                var.set(p.replace('/', '\\'))
        else:
            current = var.get()
            p = filedialog.askdirectory(initialdir=os.path.dirname(current))
            if p:
                var.set(p.replace('/', '\\') + '\\' + os.path.basename(current))

    # ══════════════════════════════════════════
    #  START / STOP
    # ══════════════════════════════════════════
    def _toggle(self):
        if not self.running:
            self.running = True
            self.btn_start.configure(text='■ STOP', bg=RED,
                                     activebackground='#b71c1c', fg='#fff')
            self.status_lbl.configure(text='RUNNING', fg=FG)
            self._canvas.itemconfigure(self._dot, fill=GREEN)
            self._worker = threading.Thread(target=self._watch_loop, daemon=True)
            self._worker.start()
            self._log('Watcher started.', 'ok')
        else:
            self.running = False
            self.btn_start.configure(text='▶ START', bg=GREEN,
                                     activebackground='#00c853', fg='#000')
            self.status_lbl.configure(text='IDLE', fg=FG_DIM)
            self._canvas.itemconfigure(self._dot, fill=FG_DIM)
            self._log('Watcher stopped.', 'w')

    # ══════════════════════════════════════════
    #  FLASH — 10 sec green on dot + STOP button
    # ══════════════════════════════════════════
    def _flash(self, filename=''):
        # turn green
        self._canvas.itemconfigure(self._dot, fill='#ffffff')
        self.btn_start.configure(bg='#00ff88', fg='#000',
                                 activebackground='#00e676')
        self.status_lbl.configure(text='COPYING…', fg=GREEN)

        def _hold():
            self._canvas.itemconfigure(self._dot, fill=GREEN)

        def _restore():
            if self.running:
                self._canvas.itemconfigure(self._dot, fill=GREEN)
                self.btn_start.configure(bg=RED, fg='#fff',
                                         activebackground='#b71c1c')
                self.status_lbl.configure(text='RUNNING', fg=FG)
            else:
                self._canvas.itemconfigure(self._dot, fill=FG_DIM)
            self.counter_lbl.configure(
                text=f'Transfers: {self._transfer_count}')

        if self._flash_job:
            self.after_cancel(self._flash_job)
        self.after(200, _hold)
        self._flash_job = self.after(10_000, _restore)

    # ══════════════════════════════════════════
    #  COPY
    # ══════════════════════════════════════════
    def _copy(self, src, dst):
        copy2(src, dst)
        self._transfer_count += 1
        self.after(0, self._flash, src)
        self._log(f'copied {os.path.basename(src)} → {dst}', 'ok')

    # ══════════════════════════════════════════
    #  DEPLOY / RECOVERY  (Safe Filename Extraction)
    # ══════════════════════════════════════════
    def _deployment_pics(self, movingImage):
        RL     = self.rl_folder_var.get()
        TD     = self.time_dif.get()
        UHD333 = self.uhd333_image_var.get()
        UHD334 = self.uhd334_image_var.get()

        filename = os.path.basename(movingImage)
        runline = re.findall(r'L_\d{3,4}', movingImage)[0][2:]
        dest = RL + '/' + runline + '/Deployed/' + filename
        m = re.search(r'T\d{14}', dest)
        dest = dest[:m.start()+1] + (
            datetime.fromtimestamp(os.path.getctime(movingImage)) +
            timedelta(hours=TD)).strftime('%H%M%S%d%m%Y') + dest[m.end():]
        try:
            self._copy(movingImage, dest)
        except FileNotFoundError:
            os.makedirs(RL + '/' + runline + '/Deployed', exist_ok=True)
            os.makedirs(RL + '/' + runline + '/Recovered', exist_ok=True)
            self._copy(movingImage, dest)

        for suffix, pattern in [('UHD333.png', UHD333), ('UHD334.png', UHD334)]:
            if movingImage[-10:] == suffix:
                pre  = max(glob.glob(pattern), key=os.path.getctime)
                dst2 = RL + '/' + runline + '/Deployed/Pre_' + filename
                m    = re.search(r'T\d{14}', dst2)
                dst2 = dst2[:m.start()+1] + (
                    datetime.fromtimestamp(os.path.getctime(pre)) +
                    timedelta(hours=TD)).strftime('%H%M%S%d%m%Y') + dst2[m.end():]
                self._copy(pre, dst2)

    def _recovery_pics(self, movingImage):
        RL     = self.rl_folder_var.get()
        TD     = self.time_dif.get()
        UHD333 = self.uhd333_image_var.get()
        UHD334 = self.uhd334_image_var.get()

        filename = os.path.basename(movingImage)
        runline = re.findall(r'L_\d{3,4}', movingImage)[0][2:]
        dest = RL + '/' + runline + '/Recovered/Post_' + filename
        m = re.search(r'T\d{14}', dest)
        dest = dest[:m.start()+1] + (
            datetime.fromtimestamp(os.path.getctime(movingImage)) +
            timedelta(hours=TD)).strftime('%H%M%S%d%m%Y') + dest[m.end():]
        try:
            self._copy(movingImage, dest)
        except FileNotFoundError:
            os.makedirs(RL + '/' + runline + '/Deployed', exist_ok=True)
            os.makedirs(RL + '/' + runline + '/Recovered', exist_ok=True)
            self._copy(movingImage, dest)

        for suffix, pattern in [('UHD333.png', UHD333), ('UHD334.png', UHD334)]:
            if movingImage[-10:] == suffix:
                pre  = max(glob.glob(pattern), key=os.path.getctime)
                dst2 = RL + '/' + runline + '/Recovered/' + filename
                m    = re.search(r'T\d{14}', dst2)
                dst2 = dst2[:m.start()+1] + (
                    datetime.fromtimestamp(os.path.getctime(pre)) +
                    timedelta(hours=TD)).strftime('%H%M%S%d%m%Y') + dst2[m.end():]
                self._copy(pre, dst2)

    # ══════════════════════════════════════════
    #  WATCH LOOP
    # ══════════════════════════════════════════
    def _watch_loop(self):
        fix_pattern    = self.fix_image_var.get()
        number_picture = len(glob.glob(fix_pattern))

        while self.running:
            try:
                fix_pattern = self.fix_image_var.get()
                files = glob.glob(fix_pattern)

                if number_picture != len(files):
                    sorted_files = sorted(files, key=os.path.getctime)
                    if self.deploying.get():
                        self._deployment_pics(sorted_files[-1])
                        number_picture += 1
                        if number_picture + 1 == len(files):
                            self._deployment_pics(sorted_files[-2])
                            number_picture += 1
                    else:
                        self._recovery_pics(sorted_files[-1])
                        number_picture += 1
                        if number_picture + 1 == len(files):
                            self._recovery_pics(sorted_files[-2])
                            number_picture += 1
                else:
                    ts = datetime.now().strftime('%H:%M:%S')
                    self.after(0, lambda t=ts: self.status_lbl.configure(
                        text=f'watch {t}'))

            except Exception as e:
                self._log(f'ERROR: {e}', 'err')

            sleep(5)

    # ══════════════════════════════════════════
    #  CLOSE
    # ══════════════════════════════════════════
    def _on_close(self):
        self.running = False
        self.destroy()


if __name__ == '__main__':
    app = App()
    app.mainloop()
