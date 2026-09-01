import os
import re
import glob
import json
import shutil
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, simpledialog
from datetime import datetime

# ─────────────────────────────────────────────
#  DEFAULT PATHS
# ─────────────────────────────────────────────
DEFAULT_FIX_IMAGE    = 'C:/Users/Public/Documents/4D Nav/NavView/OBN013626_Eng10/Local/Online/Data/Node Dashboard/Fix Images/*.png'
DEFAULT_UHD333_IMAGE = 'C:/Users/Public/Documents/4D Nav/NavView/OBN013626_Eng10/Shared/Data/Clips/UHD333 Color/*.png'
DEFAULT_UHD334_IMAGE = 'C:/Users/Public/Documents/4D Nav/NavView/OBN013626_Eng10/Shared/Data/Clips/UHD334 Color/*.png'
DEFAULT_DEST_FOLDER  = 'Z:/Projects/OBN013626_SLB_USA_Engagement10/04_SURVEY/01.Data/01.Node_Pictures'

EVENT_LABEL = 'Position deviation'   # folder name middle segment: FM-055-Position deviation[-reason]

# Report template — kept next to this script. Its 4 "Figure N" captions each
# already have a placeholder picture embedded right before them in the
# template; Generate Report swaps Figures 1-3's picture bytes for the 3 real
# event photos (in place, same zip path — no new library needed) and removes
# Figure 4's picture entirely (left blank for manual insertion later).
REPORT_TEMPLATE = 'IFR-PXGEO-OBN-013626-.docx'

# Remembers the 4 path fields across restarts — plain JSON next to the
# script, written on every change (see App._save_config()).
CONFIG_FILENAME = 'field_memo_config.json'

# Templates for text insertion under the Engagement 10 Position Deviation table
TEMPLATES_FILENAME = 'field_memo_templates.json'

# OOXML namespaces used when editing the report's word/document.xml
NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
DOCX_NS = {'w': NS_W, 'a': NS_A, 'r': NS_R}

BG      = '#0d0f14'
PANEL   = '#13161e'
BORDER  = '#1e2330'
FG      = '#c8cfe0'
FG_DIM  = '#4a5270'
GREEN   = '#00e676'
AMBER   = '#ffb300'
RED     = '#ff1744'
YELLOW  = '#ffd600'   # Generate Report button — distinct from AMBER's warning connotation
FM      = ('Courier New', 9)
FB      = ('Courier New', 11, 'bold')
FT      = ('Courier New', 12, 'bold')

FLASH_INTERVAL_MS = 500  # blink period while a button is armed/waiting for its second click


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Field Memo — Position Deviation Photo Logger')
        self.configure(bg=BG)
        self.minsize(190, 170)
        self.geometry('720x480')
        # Scalable window — unlike the old fixed-size tool, this one resizes freely.
        self.resizable(True, True)

        cfg = self._load_config()
        self.fix_image_var    = tk.StringVar(value=cfg.get('fix_image', DEFAULT_FIX_IMAGE))
        self.uhd333_image_var = tk.StringVar(value=cfg.get('uhd333_image', DEFAULT_UHD333_IMAGE))
        self.uhd334_image_var = tk.StringVar(value=cfg.get('uhd334_image', DEFAULT_UHD334_IMAGE))
        self.dest_folder_var  = tk.StringVar(value=cfg.get('dest_folder', DEFAULT_DEST_FOLDER))
        for var in (self.fix_image_var, self.uhd333_image_var,
                    self.uhd334_image_var, self.dest_folder_var):
            var.trace_add('write', lambda *args: self._save_config())

        # ── event state machine ──
        # active_uhd: None | "333" | "334" — which button (if any) is armed,
        # waiting for its second click. Only one event can be in progress at
        # a time; the other UHD button is disabled while one is active.
        self.active_uhd     = None
        self.current_folder = None
        self.reason_win     = None
        self._flash_job      = None
        self._flash_on       = False
        self.current_report_path = None  # Stores path to generated report for template insertion

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

        # ── paths (collapsible — closed by default) ──
        # Always-visible path rows were the main thing forcing a wide
        # minimum window (each needs room for a label + an editable path +
        # a Browse button). Collapsed by default, the everyday view (just
        # the two UHD buttons + log) can be much smaller; expand only when
        # you actually need to change a folder.
        paths_wrap = tk.Frame(self, bg=BG)
        paths_wrap.grid(row=1, column=0, sticky='ew', padx=10, pady=(2, 6))
        paths_wrap.columnconfigure(0, weight=1)

        self._paths_open = False
        self.paths_toggle = tk.Button(
            paths_wrap, text='▶ PATHS', font=FM, bg=PANEL, fg=FG_DIM,
            relief='flat', bd=0, anchor='w', cursor='hand2',
            activebackground=BORDER, highlightthickness=1,
            highlightbackground=BORDER, padx=8, pady=3,
            command=self._toggle_paths)
        self.paths_toggle.grid(row=0, column=0, sticky='ew')

        self.paths_body = tk.Frame(paths_wrap, bg=BG)
        self.paths_body.columnconfigure(1, weight=1)

        entries = [
            ('Fix Img', self.fix_image_var, False),
            ('UHD333',  self.uhd333_image_var, False),
            ('UHD334',  self.uhd334_image_var, False),
            ('Dest',    self.dest_folder_var, True),
        ]
        for i, (lbl, var, is_dir) in enumerate(entries):
            tk.Label(self.paths_body, text=lbl, font=FM, bg=BG, fg=FG_DIM,
                      anchor='w', width=7
                      ).grid(row=i, column=0, sticky='w', padx=(2, 2), pady=3)
            tk.Entry(self.paths_body, textvariable=var, font=FM, bg=PANEL, fg=FG,
                      insertbackground=FG, relief='flat', bd=0,
                      highlightthickness=1, highlightcolor=GREEN,
                      highlightbackground=BORDER
                      ).grid(row=i, column=1, sticky='ew', pady=3)
            tk.Button(self.paths_body, text='…', font=FM, bg=BORDER, fg=FG, relief='flat',
                      bd=0, cursor='hand2', padx=6,
                      command=lambda v=var, d=is_dir: self._browse(v, d)
                      ).grid(row=i, column=2, padx=(4, 2), pady=3)

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
            bot, text='UHD333', font=FB, bg=PANEL, fg=RED,
            activebackground=BORDER, relief='flat', bd=0, cursor='hand2',
            pady=10, command=lambda: self._on_uhd_click('333'))
        self.btn_uhd333.grid(row=0, column=0, sticky='ew', padx=(0, 4))

        self.btn_uhd334 = tk.Button(
            bot, text='UHD334', font=FB, bg=PANEL, fg=GREEN,
            activebackground=BORDER, relief='flat', bd=0, cursor='hand2',
            pady=10, command=lambda: self._on_uhd_click('334'))
        self.btn_uhd334.grid(row=0, column=1, sticky='ew', padx=(4, 0))

        self._buttons = {'333': self.btn_uhd333, '334': self.btn_uhd334}
        self._idle_fg = {'333': RED, '334': GREEN}  # each button's label color when not flashing

        bottom_row = tk.Frame(self, bg=BG)
        bottom_row.grid(row=4, column=0, sticky='ew', padx=10, pady=(0, 8))
        bottom_row.columnconfigure(0, weight=1)  # spacer — pushes both buttons to the right

        tk.Button(bottom_row, text='GENERATE REPORT', font=FM, bg=BORDER, fg=YELLOW,
                  relief='flat', bd=0, cursor='hand2', padx=8, pady=3,
                  command=self._generate_report).grid(row=0, column=1, padx=(0, 6))
        tk.Button(bottom_row, text='CLR LOG', font=FM, bg=BORDER, fg='#ffffff',
                  relief='flat', bd=0, cursor='hand2', padx=8, pady=3,
                  command=self._clear_log).grid(row=0, column=2)

    def _toggle_paths(self):
        self._paths_open = not self._paths_open
        if self._paths_open:
            self.paths_toggle.configure(text='▼ PATHS')
            self.paths_body.grid(row=1, column=0, sticky='ew', pady=(4, 0))
        else:
            self.paths_toggle.configure(text='▶ PATHS')
            self.paths_body.grid_forget()

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
    #  SETTINGS PERSISTENCE
    # ══════════════════════════════════════════
    def _config_path(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILENAME)

    def _templates_path(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), TEMPLATES_FILENAME)

    def _load_templates(self):
        """Load templates from JSON file."""
        path = self._templates_path()
        if os.path.isfile(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get('templates', [])
            except Exception:
                pass
        return []

    def _save_templates(self, templates):
        """Save templates to JSON file."""
        data = {'templates': templates}
        try:
            with open(self._templates_path(), 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            self._log(f'✗ Could not save templates: {e}', 'err')

    def _extract_fm_number(self, folder_path):
        """Extract FM number from folder path (e.g., FM-001 from 'FM-001-Position deviation')."""
        folder_name = os.path.basename(folder_path)
        m = re.match(r'(FM-\d+)', folder_name)
        if m:
            return m.group(1).replace('FM-', '')  # Return just the number part (001)
        return '000'

    def _get_formatted_date(self):
        """Get current system date in Day/Month/Year format."""
        today = datetime.now()
        return today.strftime('%d/%m/%Y')

    def _load_config(self):
        path = self._config_path()
        if os.path.isfile(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_config(self):
        data = {
            'fix_image':    self.fix_image_var.get(),
            'uhd333_image': self.uhd333_image_var.get(),
            'uhd334_image': self.uhd334_image_var.get(),
            'dest_folder':  self.dest_folder_var.get(),
        }
        try:
            with open(self._config_path(), 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass  # best-effort — a failed save shouldn't interrupt the tool

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

    def _copy_file(self, src, dest_folder, label, dest_name=None):
        name = dest_name if dest_name else os.path.basename(src)
        try:
            shutil.copy2(src, os.path.join(dest_folder, name))
            self._log(f'✓ Copied {label}: {name}', 'ok')
            return True
        except Exception as e:
            self._log(f'✗ Copy failed for {label}: {e}', 'err')
            return False

    def _copy_latest(self, pattern, dest_folder, label, dest_name=None):
        src = self._latest_photo(pattern)
        if not src:
            self._log(f'✗ No photos found for {label} ({pattern})', 'err')
            return None
        self._copy_file(src, dest_folder, label, dest_name=dest_name)
        return src

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

        # The UHD "#2" photo is named after the Fix Image (Pre_<fix filename>),
        # not its own filename — Fix Images use a different "L_..." naming
        # scheme entirely. Resolve the Fix Image's path ONCE and reuse it for
        # both the naming and the actual copy below, so the two can't
        # possibly disagree about which file "the latest Fix Image" was.
        fix_src = self._latest_photo(self.fix_image_var.get())
        if fix_src:
            pre_name = 'Pre_' + os.path.basename(fix_src)
            self._copy_latest(self._uhd_pattern(which), folder_path, f'UHD{which} #2', dest_name=pre_name)
            self._copy_file(fix_src, folder_path, 'Fix Image')
        else:
            self._log(f'✗ No photos found for Fix Image ({self.fix_image_var.get()})', 'err')
            self._log('⚠ No Fix Image found — UHD photo kept its own filename', 'w')
            self._copy_latest(self._uhd_pattern(which), folder_path, f'UHD{which} #2')

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
        for key, btn in self._buttons.items():
            btn.configure(state='normal', bg=PANEL, fg=self._idle_fg[key])

    def _start_flash(self, which):
        self._flash_on = False
        self._do_flash(which)

    def _do_flash(self, which):
        btn = self._buttons[which]
        self._flash_on = not self._flash_on
        if self._flash_on:
            btn.configure(bg=GREEN, fg='#000')
        else:
            btn.configure(bg=PANEL, fg=self._idle_fg[which])
        self._flash_job = self.after(FLASH_INTERVAL_MS, lambda: self._do_flash(which))

    def _stop_flash(self, which):
        if self._flash_job is not None:
            self.after_cancel(self._flash_job)
            self._flash_job = None
        self._buttons[which].configure(bg=PANEL, fg=self._idle_fg[which])

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

    # ══════════════════════════════════════════
    #  GENERATE REPORT
    # ══════════════════════════════════════════
    def _latest_event_folder(self, dest_root):
        """Return (path, name) of the highest-numbered existing
        'FM-###-...' folder in dest_root, or (None, None)."""
        best_num, best_name = -1, None
        if os.path.isdir(dest_root):
            for name in os.listdir(dest_root):
                m = re.match(r'FM-(\d+)-', name)
                if m and int(m.group(1)) > best_num:
                    best_num, best_name = int(m.group(1)), name
        if best_name is None:
            return None, None
        return os.path.join(dest_root, best_name), best_name

    def _identify_event_photos(self, folder_path):
        """Identify (photo1, photo2, fix_photo) among the files already in
        an event folder, using the naming convention _complete_event() set
        up: photo2 is the file prefixed 'Pre_', the fix photo is that same
        name minus the prefix, and whatever's left over is photo1. Any of
        the three may come back None if it can't be determined."""
        files = [f for f in os.listdir(folder_path)
                 if os.path.isfile(os.path.join(folder_path, f))]

        pre_file = next((f for f in files if f.startswith('Pre_')), None)
        fix_file = None
        if pre_file:
            candidate = pre_file[len('Pre_'):]
            if candidate in files:
                fix_file = candidate

        remaining = [f for f in files if f not in (pre_file, fix_file)]
        if len(remaining) > 1:
            remaining.sort(key=lambda f: os.path.getctime(os.path.join(folder_path, f)))
        photo1_file = remaining[0] if remaining else None

        def full(f):
            return os.path.join(folder_path, f) if f else None
        return full(photo1_file), full(pre_file), full(fix_file)

    # ---- docx editing helpers (pure zipfile/ElementTree — no extra library) ----
    def _docx_rel_map(self, docx_path):
        with zipfile.ZipFile(docx_path) as z:
            rels_xml = z.read('word/_rels/document.xml.rels').decode('utf-8')
        rel_root = ET.fromstring(rels_xml)
        return {rel.get('Id'): rel.get('Target') for rel in rel_root}

    def _docx_figure_rid(self, root, caption_substr):
        """Relationship id of the picture immediately preceding the
        paragraph whose text contains caption_substr (e.g. 'Figure 2')."""
        body = root.find('w:body', DOCX_NS)
        paras = list(body.iter('{%s}p' % NS_W))
        for i, p in enumerate(paras):
            texts = ''.join(t.text or '' for t in p.findall('.//w:t', DOCX_NS))
            if caption_substr in texts:
                for j in range(i - 1, -1, -1):
                    blip = paras[j].find('.//a:blip', DOCX_NS)
                    if blip is not None:
                        return blip.get('{%s}embed' % NS_R)
                return None
        return None

    def _docx_run_span_for_rid(self, xml_text, rid):
        """Offsets (start, end) of the whole <w:r>...</w:r> run that embeds
        the picture with the given relationship id, for surgical removal."""
        idx = xml_text.find(f'r:embed="{rid}"')
        if idx < 0:
            return None
        draw_start = xml_text.rfind('<w:drawing', 0, idx)
        draw_end = xml_text.find('</w:drawing>', idx) + len('</w:drawing>')
        run_start = max(xml_text.rfind('<w:r>', 0, draw_start), xml_text.rfind('<w:r ', 0, draw_start))
        run_end = xml_text.find('</w:r>', draw_end) + len('</w:r>')
        return run_start, run_end

    def _build_report(self, template_path, out_path, photo1, photo2, fix_photo):
        with zipfile.ZipFile(template_path) as z:
            doc_xml_bytes = z.read('word/document.xml')
        rel_map = self._docx_rel_map(template_path)
        root = ET.fromstring(doc_xml_bytes)

        fig_rid = {n: self._docx_figure_rid(root, f'Figure {n}') for n in (1, 2, 3, 4)}
        doc_xml_text = doc_xml_bytes.decode('utf-8')

        # Figure 4 — remove its picture entirely, left blank for manual
        # insertion later (per instruction: do not insert anything there).
        rid4 = fig_rid.get(4)
        if rid4:
            span = self._docx_run_span_for_rid(doc_xml_text, rid4)
            if span:
                doc_xml_text = doc_xml_text[:span[0]] + doc_xml_text[span[1]:]

        # Figures 1-3 — swap each placeholder picture's bytes for the real
        # event photo, keeping the same zip path so the template's existing
        # relationships/sizing/captions stay untouched.
        photos = {1: photo1, 2: photo2, 3: fix_photo}
        targets_to_replace = {}
        for n in (1, 2, 3):
            rid, src = fig_rid.get(n), photos.get(n)
            if rid and src and rel_map.get(rid):
                with open(src, 'rb') as f:
                    targets_to_replace['word/' + rel_map[rid]] = f.read()

        with zipfile.ZipFile(template_path) as zin, \
             zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item, doc_xml_text.encode('utf-8'))
                elif item.filename in targets_to_replace:
                    zout.writestr(item, targets_to_replace[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))

    def _insert_template_into_docx(self, docx_path, template_text, fm_number, line_num, station_num, node_num, author):
        """Update the actual placeholders in the template and insert the template text."""
        try:
            ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
            ET.register_namespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
            ET.register_namespace('v', 'urn:schemas-microsoft-com:vml')

            def set_font_12(run, bold=False):
                rpr = run.find('{%s}rPr' % NS_W)
                if rpr is None:
                    rpr = ET.SubElement(run, '{%s}rPr' % NS_W)
                fonts = rpr.find('{%s}rFonts' % NS_W)
                if fonts is None:
                    fonts = ET.SubElement(rpr, '{%s}rFonts' % NS_W)
                fonts.set('{%s}ascii' % NS_W, 'Arial')
                fonts.set('{%s}hAnsi' % NS_W, 'Arial')
                fonts.set('{%s}eastAsia' % NS_W, 'Arial')
                if bold:
                    if rpr.find('{%s}b' % NS_W) is None:
                        ET.SubElement(rpr, '{%s}b' % NS_W)
                    if rpr.find('{%s}bCs' % NS_W) is None:
                        ET.SubElement(rpr, '{%s}bCs' % NS_W)
                else:
                    for tag in ('{%s}b' % NS_W, '{%s}bCs' % NS_W):
                        el = rpr.find(tag)
                        if el is not None:
                            rpr.remove(el)
                sz = rpr.find('{%s}sz' % NS_W)
                if sz is None:
                    sz = ET.SubElement(rpr, '{%s}sz' % NS_W)
                sz.set('{%s}val' % NS_W, '24')
                szcs = rpr.find('{%s}szCs' % NS_W)
                if szcs is None:
                    szcs = ET.SubElement(rpr, '{%s}szCs' % NS_W)
                szcs.set('{%s}val' % NS_W, '24')

            def replace_paragraph_text(paragraph, new_text, bold=False):
                for child in list(paragraph):
                    if child.tag != '{%s}pPr' % NS_W:
                        paragraph.remove(child)
                run = ET.SubElement(paragraph, '{%s}r' % NS_W)
                set_font_12(run, bold=bold)
                text_el = ET.SubElement(run, '{%s}t' % NS_W)
                if new_text and (new_text[0].isspace() or new_text[-1].isspace() or '  ' in new_text):
                    text_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                text_el.text = new_text

            def replace_cell_text(cell, new_text, bold=False):
                # Keep the first existing paragraph's pPr (indent, cnfStyle
                # table-banding) so the inserted text stays aligned with the
                # other rows instead of landing flush-left — matching
                # IFR-PXGEO-OBN-013626-FM-074.docx.
                keep_pPr = None
                for para in list(cell):
                    if para.tag == '{%s}p' % NS_W:
                        if keep_pPr is None:
                            keep_pPr = para.find('{%s}pPr' % NS_W)
                        cell.remove(para)
                para = ET.Element('{%s}p' % NS_W)
                if keep_pPr is not None:
                    para.append(keep_pPr)
                cell.append(para)
                replace_paragraph_text(para, new_text, bold=bold)

            with zipfile.ZipFile(docx_path, 'r') as z:
                doc_xml_bytes = z.read('word/document.xml')

            root = ET.fromstring(doc_xml_bytes)
            body = root.find('w:body', DOCX_NS)
            if body is None:
                self._log('✗ Cannot find document body in Word file', 'err')
                return False

            current_date = self._get_formatted_date()
            day_str, month_str, year_str = current_date.split('/')
            line_text = str(int(line_num))
            station_text = str(int(station_num))
            node_text = str(int(node_num))

            # Edit only the placeholder runs' text in place, so the template's
            # existing tabs / bold / right-alignment (e.g. the tab-aligned
            # "REVISION 1" next to the title, or the tabs that right-align the
            # date) survive untouched — matching IFR-PXGEO-OBN-013626-FM-074.docx.
            for para in body.iter('{%s}p' % NS_W):
                t_elems = para.findall('.//w:t', DOCX_NS)
                txt = ''.join(t.text or '' for t in t_elems)
                if 'IFR-PXGEO-OBN-013626-' in txt and 'REVISION' in txt:
                    for t in t_elems:
                        if t.text and 'IFR-PXGEO-OBN-013626-' in t.text:
                            t.text = t.text.replace(
                                'IFR-PXGEO-OBN-013626-',
                                f'IFR-PXGEO-OBN-013626-FM-{fm_number}',
                            )
                            break
                elif 'IFR-PXGEO-OBN-013626-' in txt:
                    # The "Document Number:" table row carries the same bare
                    # prefix as the title line, just without "REVISION" next
                    # to it — needs the same FM suffix appended.
                    for t in t_elems:
                        if t.text and 'IFR-PXGEO-OBN-013626-' in t.text:
                            t.text = t.text.replace(
                                'IFR-PXGEO-OBN-013626-',
                                f'IFR-PXGEO-OBN-013626-FM-{fm_number}',
                            )
                            break
                elif txt.strip() == 'Day/Month/Year' or ('Day' in txt and 'Month' in txt and 'Year' in txt):
                    for t in t_elems:
                        if t.text == 'Day':
                            t.text = day_str
                        elif t.text == 'Month':
                            t.text = month_str
                        elif t.text == 'Year':
                            t.text = year_str
                elif 'Line ' in txt and 'Station ' in txt and 'Node ' in txt and 'Engagement 10 Position Deviation' in txt:
                    placeholders = [t for t in t_elems if t.text == 'XXX']
                    for t, value in zip(placeholders, (line_text, station_text, node_text)):
                        t.text = value

            # Update the date/author cells in the metadata table.
            tables = body.findall('w:tbl', DOCX_NS)
            for table in tables:
                table_texts = ''.join(t.text or '' for t in table.iter('{%s}t' % NS_W))
                if 'Engagement 10 Position Deviation' in table_texts:
                    rows = table.findall('{%s}tr' % NS_W)
                    # Index into each row's <w:tc> cells specifically, not the
                    # row's raw children — a <w:trPr> sits before them and
                    # shifts plain rows[i][1] indexing onto the label cell
                    # instead of the value cell next to it.
                    if len(rows) > 1:
                        tcs = rows[1].findall('{%s}tc' % NS_W)
                        if len(tcs) > 1:
                            replace_cell_text(tcs[1], current_date, bold=False)
                    if len(rows) > 3:
                        tcs = rows[3].findall('{%s}tc' % NS_W)
                        if len(tcs) > 1:
                            replace_cell_text(tcs[1], author, bold=False)
                    break

            # Insert the template text after the main target table.
            insert_after_table = None
            for table in tables:
                table_texts = ''.join(t.text or '' for t in table.iter('{%s}t' % NS_W))
                if 'Engagement 10 Position Deviation' in table_texts:
                    insert_after_table = table
                    break
            if insert_after_table is None:
                self._log('⚠ "Engagement 10 Position Deviation" table not found', 'w')
                return False

            def make_normal_paragraph(line_text=''):
                p = ET.Element('{%s}p' % NS_W)
                pPr = ET.SubElement(p, '{%s}pPr' % NS_W)
                ET.SubElement(pPr, '{%s}pStyle' % NS_W).set('{%s}val' % NS_W, 'Normal')
                if line_text:
                    run = ET.SubElement(p, '{%s}r' % NS_W)
                    set_font_12(run, bold=False)
                    text_el = ET.SubElement(run, '{%s}t' % NS_W)
                    if line_text[0].isspace() or line_text[-1].isspace() or '  ' in line_text:
                        text_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    text_el.text = line_text
                return p

            # A blank spacer paragraph between the table and the template
            # text, so it doesn't sit flush against the table border.
            # Each Enter the user typed in the editor becomes its own <w:p>,
            # so paragraph breaks in the editor carry through the same way
            # into the generated report.
            insert_idx = list(body).index(insert_after_table) + 1
            body.insert(insert_idx, make_normal_paragraph())
            for offset, line_text in enumerate(template_text.split('\n'), start=1):
                body.insert(insert_idx + offset, make_normal_paragraph(line_text))

            updated_xml = ET.tostring(root, encoding='utf-8')
            with zipfile.ZipFile(docx_path, 'r') as zin, zipfile.ZipFile(docx_path + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == 'word/document.xml':
                        zout.writestr(item, updated_xml)
                    else:
                        zout.writestr(item, zin.read(item.filename))

            os.remove(docx_path)
            os.rename(docx_path + '.tmp', docx_path)

            self._log('✓ Template and metadata inserted into report', 'ok')
            return True
        except Exception as e:
            self._log(f'✗ Failed to insert template: {e}', 'err')
            import traceback
            traceback.print_exc()
            return False

    def _insert_figure4_image(self, docx_path, image_path):
        """Insert image as Figure 4 (Navview Map) into the document.
        This adds the image to the DOCX, updates relationships, and embeds it in the document."""
        try:
            if not os.path.isfile(image_path):
                self._log(f'✗ Image file not found: {image_path}', 'err')
                return False
            
            # Read existing DOCX
            with zipfile.ZipFile(docx_path, 'r') as z:
                doc_xml_bytes = z.read('word/document.xml')
                rels_xml_bytes = z.read('word/_rels/document.xml.rels')
            
            doc_xml_text = doc_xml_bytes.decode('utf-8')
            root = ET.fromstring(doc_xml_bytes)
            rels_root = ET.fromstring(rels_xml_bytes)
            
            body = root.find('w:body', DOCX_NS)
            if body is None:
                self._log('✗ Cannot find document body', 'err')
                return False
            
            # Find Figure 4 caption
            paras = list(body.iter('{%s}p' % NS_W))
            fig4_para_idx = None
            for i, p in enumerate(paras):
                texts = ''.join(t.text or '' for t in p.findall('.//w:t', DOCX_NS))
                if 'Figure 4' in texts:
                    fig4_para_idx = i
                    break
            
            if fig4_para_idx is None:
                self._log('⚠ Figure 4 caption not found in document', 'w')
                return False
            
            # Find the paragraph before Figure 4 caption (where we'll insert the image)
            if fig4_para_idx > 0:
                insert_para = paras[fig4_para_idx - 1]
            else:
                insert_para = paras[fig4_para_idx]
            
            # Generate new relationship ID for the image
            existing_rels = rels_root.findall('{%s}Relationship' % 'http://schemas.openxmlformats.org/package/2006/relationships')
            max_rel_id = 0
            for rel in existing_rels:
                rid = rel.get('Id')
                if rid and rid.startswith('rId'):
                    try:
                        num = int(rid[3:])
                        max_rel_id = max(max_rel_id, num)
                    except:
                        pass
            
            new_rel_id = f'rId{max_rel_id + 1}'
            
            # Add relationship for the image
            NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
            rel_elem = ET.Element('{%s}Relationship' % NS_REL)
            rel_elem.set('Id', new_rel_id)
            rel_elem.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
            
            # Determine image file extension
            _, ext = os.path.splitext(image_path)
            image_filename = f'image{max_rel_id + 1}{ext}'
            rel_elem.set('Target', f'media/{image_filename}')
            rels_root.append(rel_elem)
            
            # Create drawing element with the image
            drawing = ET.Element('{%s}drawing' % NS_W)
            inline = ET.SubElement(drawing, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline')
            inline.set('distT', '0')
            inline.set('distB', '0')
            inline.set('distL', '0')
            inline.set('distR', '0')
            
            # Set extent (size) - 6 inches wide
            extent = ET.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
            extent.set('cx', '5486400')  # 6 inches in EMUs
            extent.set('cy', '4114800')  # proportional height

            # wp:docPr is a REQUIRED child of wp:inline (right after extent,
            # before graphic) per the OOXML schema. Omitting it is exactly
            # what makes Word flag the file as unreadable content needing
            # recovery — it's not optional the way cNvGraphicFramePr is.
            existing_docpr_ids = [
                int(dp.get('id')) for dp in root.iter(
                    '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr')
                if dp.get('id') and dp.get('id').isdigit()
            ]
            new_docpr_id = max(existing_docpr_ids, default=0) + 1
            docPr = ET.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr')
            docPr.set('id', str(new_docpr_id))
            docPr.set('name', f'Picture {new_docpr_id}')

            # Add graphic
            graphic = ET.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
            graphicData = ET.SubElement(graphic, '{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData')
            graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
            
            pic = ET.SubElement(graphicData, '{http://schemas.openxmlformats.org/drawingml/2006/picture}pic')
            nvPicPr = ET.SubElement(pic, '{http://schemas.openxmlformats.org/drawingml/2006/picture}nvPicPr')
            
            cNvPr = ET.SubElement(nvPicPr, '{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPr')
            cNvPr.set('id', str(new_docpr_id))
            cNvPr.set('name', 'Navview Map')
            
            cNvPicPr = ET.SubElement(nvPicPr, '{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPicPr')
            blipFill = ET.SubElement(pic, '{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill')
            blip = ET.SubElement(blipFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
            blip.set('{%s}embed' % NS_R, new_rel_id)
            
            stretch = ET.SubElement(blipFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}stretch')
            fillRect = ET.SubElement(stretch, '{http://schemas.openxmlformats.org/drawingml/2006/main}fillRect')
            
            spPr = ET.SubElement(pic, '{http://schemas.openxmlformats.org/drawingml/2006/picture}spPr')
            xfrm = ET.SubElement(spPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
            off = ET.SubElement(xfrm, '{http://schemas.openxmlformats.org/drawingml/2006/main}off')
            off.set('x', '0')
            off.set('y', '0')
            ext = ET.SubElement(xfrm, '{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
            ext.set('cx', '5486400')
            ext.set('cy', '4114800')
            
            prstGeom = ET.SubElement(spPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom')
            prstGeom.set('prst', 'rect')
            avLst = ET.SubElement(prstGeom, '{http://schemas.openxmlformats.org/drawingml/2006/main}avLst')
            
            # Create a new run with the drawing
            new_run = ET.Element('{%s}r' % NS_W)
            new_run.append(drawing)
            
            # Insert run into a new paragraph before Figure 4 caption
            new_para = ET.Element('{%s}p' % NS_W)
            new_para.append(new_run)
            
            insert_para_index = list(body).index(insert_para)
            body.insert(insert_para_index + 1, new_para)
            
            # Read image file as bytes
            with open(image_path, 'rb') as f:
                image_bytes = f.read()
            
            # Write updated DOCX
            updated_doc_xml = ET.tostring(root, encoding='utf-8')
            updated_rels_xml = ET.tostring(rels_root, encoding='utf-8')
            
            with zipfile.ZipFile(docx_path, 'r') as zin, \
                 zipfile.ZipFile(docx_path + '.tmp', 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == 'word/document.xml':
                        zout.writestr(item, updated_doc_xml)
                    elif item.filename == 'word/_rels/document.xml.rels':
                        zout.writestr(item, updated_rels_xml)
                    else:
                        zout.writestr(item, zin.read(item.filename))
                
                # Add the image file
                zout.writestr(f'word/media/{image_filename}', image_bytes)
            
            os.remove(docx_path)
            os.rename(docx_path + '.tmp', docx_path)
            
            self._log(f'✓ Image inserted as Figure 4', 'ok')
            return True
        except Exception as e:
            self._log(f'✗ Failed to insert image: {e}', 'err')
            return False

    def _generate_report(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(script_dir, REPORT_TEMPLATE)
        if not os.path.isfile(template_path):
            self._log(f'✗ Report template not found next to the script: {REPORT_TEMPLATE}', 'err')
            return

        dest_root = self.dest_folder_var.get()
        folder_path, folder_name = self._latest_event_folder(dest_root)
        if not folder_path:
            self._log('✗ No FM-### event folder found in Destination', 'err')
            return

        photo1, photo2, fix_photo = self._identify_event_photos(folder_path)
        if not (photo1 and photo2 and fix_photo):
            self._log(f'⚠ Could not identify all 3 photos in {folder_name} — '
                       f'report will leave any unmatched figure as the template default', 'w')

        m = re.match(r'(FM-\d+)', folder_name)
        fm_id = m.group(1) if m else folder_name
        base, ext = os.path.splitext(REPORT_TEMPLATE)
        out_path = os.path.join(folder_path, f'{base}{fm_id}{ext}')

        try:
            self._build_report(template_path, out_path, photo1, photo2, fix_photo)
            self._log(f'✓ Generated report: {os.path.basename(out_path)}', 'ok')
            self.current_report_path = out_path
            self._open_template_manager(out_path, folder_path)
        except Exception as e:
            self._log(f'✗ Report generation failed: {e}', 'err')

    def _open_template_manager(self, report_path, folder_path):
        """Open template manager window for adding templates to the report."""
        win = tk.Toplevel(self)
        win.title('Template Manager')
        win.configure(bg=BG)
        win.geometry('640x680')
        win.minsize(480, 500)
        win.transient(self)

        # Extract FM number from folder path
        fm_number = self._extract_fm_number(folder_path)

        # ── Title ──
        tk.Label(win, text='Engagement 10 Templates & Metadata', font=FB, bg=BG, fg=GREEN).pack(padx=12, pady=(12, 6), anchor='w')

        # ── Templates — one button per saved template ──
        # templates: list of {'name': short label shown on the button, 'text': saved body}.
        # selected['index'] tracks which template is currently loaded into the
        # editor below, so Update/Delete know which one to act on.
        templates = self._load_templates()
        selected = {'index': None}
        template_buttons = []
        TEMPLATES_PER_ROW = 4

        tk.Label(win, text='Templates — click to load, edit freely below, then Insert:',
                 font=FM, bg=BG, fg=FG_DIM).pack(padx=12, pady=(6, 2), anchor='w')

        buttons_wrap = tk.Frame(win, bg=BG)
        buttons_wrap.pack(fill='x', padx=12, pady=(0, 4))

        def highlight_selected():
            for i, b in enumerate(template_buttons):
                if i == selected['index']:
                    b.configure(bg=GREEN, fg='#000')
                else:
                    b.configure(bg=PANEL, fg=FG)

        def load_template(i):
            selected['index'] = i
            editor.delete('1.0', 'end')
            editor.insert('1.0', templates[i]['text'])
            highlight_selected()

        def rebuild_template_buttons():
            for child in buttons_wrap.winfo_children():
                child.destroy()
            template_buttons.clear()
            row = None
            for i, tmpl in enumerate(templates):
                if i % TEMPLATES_PER_ROW == 0:
                    row = tk.Frame(buttons_wrap, bg=BG)
                    row.pack(fill='x', pady=2)
                b = tk.Button(row, text=tmpl['name'], font=FM, relief='flat', bd=0,
                              cursor='hand2', padx=10, pady=4, bg=PANEL, fg=FG,
                              command=lambda i=i: load_template(i))
                b.pack(side='left', padx=(0, 4))
                template_buttons.append(b)
            highlight_selected()

        def add_template():
            name = simpledialog.askstring(
                'New Template', 'Short name (e.g. "Slope", "Benthic"):', parent=win)
            if not name or not name.strip():
                return
            templates.append({'name': name.strip(), 'text': ''})
            self._save_templates(templates)
            rebuild_template_buttons()
            load_template(len(templates) - 1)
            editor.focus_set()
            self._log(f'✓ Template "{name.strip()}" added', 'ok')

        def update_template():
            if selected['index'] is None:
                self._log('✗ Select a template first', 'err')
                return
            templates[selected['index']]['text'] = editor.get('1.0', 'end-1c')
            self._save_templates(templates)
            self._log(f'✓ Template "{templates[selected["index"]]["name"]}" updated', 'ok')

        def delete_template():
            if selected['index'] is None:
                self._log('✗ Select a template first', 'err')
                return
            name = templates[selected['index']]['name']
            templates.pop(selected['index'])
            self._save_templates(templates)
            selected['index'] = None
            editor.delete('1.0', 'end')
            rebuild_template_buttons()
            self._log(f'✓ Template "{name}" deleted', 'ok')

        tk.Button(win, text='+ Add Template', font=FM, bg=GREEN, fg='#000', relief='flat',
                  bd=0, cursor='hand2', padx=10, pady=3,
                  command=add_template).pack(padx=12, pady=(0, 6), anchor='w')

        # ── Editable template text — a real multi-line Text widget, so
        # Enter starts a new paragraph here exactly like it will in Word,
        # instead of the single-line Entry that could only hold one line. ──
        tk.Label(win, text='Template Text (Enter = new paragraph in the report):',
                 font=FM, bg=BG, fg=FG_DIM).pack(padx=12, pady=(6, 2), anchor='w')

        editor_frame = tk.Frame(win, bg=BG)
        editor_frame.pack(fill='both', expand=True, padx=12, pady=(0, 6))
        editor_scroll = tk.Scrollbar(editor_frame, bg=BORDER, troughcolor=BG, relief='flat')
        editor_scroll.pack(side='right', fill='y')
        editor = tk.Text(editor_frame, font=FM, bg=PANEL, fg=FG, insertbackground=FG,
                          relief='flat', bd=0, highlightthickness=1, highlightcolor=GREEN,
                          highlightbackground=BORDER, wrap='word', height=8,
                          yscrollcommand=editor_scroll.set)
        editor.pack(side='left', fill='both', expand=True)
        editor_scroll.config(command=editor.yview)

        edit_btn_row = tk.Frame(win, bg=BG)
        edit_btn_row.pack(fill='x', padx=12, pady=(0, 6))
        tk.Button(edit_btn_row, text='Update Template', font=FM, bg=BORDER, fg=YELLOW, relief='flat',
                  bd=0, cursor='hand2', padx=10, command=update_template).pack(side='left', padx=(0, 3))
        tk.Button(edit_btn_row, text='Delete Template', font=FM, bg=RED, fg='#fff', relief='flat',
                  bd=0, cursor='hand2', padx=10, command=delete_template).pack(side='left', padx=3)

        rebuild_template_buttons()

        # ── Metadata Fields ──
        meta_frame = tk.LabelFrame(win, text='Document Metadata', font=FM, bg=BG, fg=FG_DIM, relief='flat', bd=0)
        meta_frame.pack(fill='x', padx=12, pady=6)

        # Line number
        line_frame = tk.Frame(meta_frame, bg=BG)
        line_frame.pack(fill='x', pady=3)
        tk.Label(line_frame, text='Line #:', font=FM, bg=BG, fg=FG_DIM, width=12, anchor='w').pack(side='left')
        line_var = tk.StringVar(value='00001')
        line_entry = tk.Entry(line_frame, textvariable=line_var, font=FM, bg=PANEL, fg=FG, width=10,
                              insertbackground=FG, relief='flat', bd=0, highlightthickness=1, highlightbackground=BORDER)
        line_entry.pack(side='left', padx=(0, 6))

        # Station number
        tk.Label(line_frame, text='Station #:', font=FM, bg=BG, fg=FG_DIM, width=12, anchor='w').pack(side='left')
        station_var = tk.StringVar(value='00001')
        station_entry = tk.Entry(line_frame, textvariable=station_var, font=FM, bg=PANEL, fg=FG, width=10,
                                 insertbackground=FG, relief='flat', bd=0, highlightthickness=1, highlightbackground=BORDER)
        station_entry.pack(side='left', padx=(0, 6))

        # Node number
        tk.Label(line_frame, text='Node #:', font=FM, bg=BG, fg=FG_DIM, width=12, anchor='w').pack(side='left')
        node_var = tk.StringVar(value='00001')
        node_entry = tk.Entry(line_frame, textvariable=node_var, font=FM, bg=PANEL, fg=FG, width=10,
                              insertbackground=FG, relief='flat', bd=0, highlightthickness=1, highlightbackground=BORDER)
        node_entry.pack(side='left')

        # Author name
        author_frame = tk.Frame(meta_frame, bg=BG)
        author_frame.pack(fill='x', pady=3)
        tk.Label(author_frame, text='Author:', font=FM, bg=BG, fg=FG_DIM, width=12, anchor='w').pack(side='left')
        author_var = tk.StringVar()
        author_entry = tk.Entry(author_frame, textvariable=author_var, font=FM, bg=PANEL, fg=FG,
                                insertbackground=FG, relief='flat', bd=0, highlightthickness=1, highlightbackground=BORDER)
        author_entry.pack(side='left', fill='x', expand=True)

        # FM number display
        fm_frame = tk.Frame(meta_frame, bg=BG)
        fm_frame.pack(fill='x', pady=3)
        tk.Label(fm_frame, text='FM #:', font=FM, bg=BG, fg=FG_DIM, width=12, anchor='w').pack(side='left')
        tk.Label(fm_frame, text=fm_number, font=FM, bg=PANEL, fg=GREEN, width=10, anchor='w',
                 relief='flat', bd=0).pack(side='left')
        tk.Label(fm_frame, text='(Auto-detected)', font=FM, bg=BG, fg=FG_DIM).pack(side='left', padx=(6, 0))

        # ── Buttons ──
        btn_frame = tk.Frame(win, bg=BG)
        btn_frame.pack(fill='x', padx=12, pady=(0, 12))

        image_selected = {'path': None}  # Track selected image

        def browse_image():
            """Browse and select image to insert as Figure 4."""
            img_path = filedialog.askopenfilename(
                initialdir=folder_path,
                title='Select Image for Figure 4 (Navview Map)',
                filetypes=[('Image Files', '*.png *.jpg *.jpeg *.bmp'), ('All Files', '*.*')]
            )
            if img_path:
                image_selected['path'] = img_path
                if self._insert_figure4_image(report_path, img_path):
                    self._log(f'✓ Figure 4 image selected and inserted', 'ok')
                    # Update button text to show image is selected
                    browse_btn.config(text=f'✓ Image: {os.path.basename(img_path)[:20]}', fg=GREEN)

        def insert_template():
            text = editor.get('1.0', 'end-1c')
            if not text.strip():
                self._log('✗ No template text to insert — load or type one first', 'err')
                return

            # Validate inputs
            line = line_var.get().strip()
            station = station_var.get().strip()
            node = node_var.get().strip()
            author = author_var.get().strip()

            if not author:
                self._log('✗ Please enter Author name', 'err')
                return

            try:
                int(line)
                int(station)
                int(node)
            except ValueError:
                self._log('✗ Line, Station, and Node must be numeric', 'err')
                return

            if self._insert_template_into_docx(report_path, text, fm_number, line, station, node, author):
                self._log(f'✓ Template and metadata inserted', 'ok')
                win.destroy()

        browse_btn = tk.Button(btn_frame, text='Browse & Insert Image', font=FM, bg=BORDER, fg=YELLOW, relief='flat',
                  bd=0, cursor='hand2', padx=10, command=browse_image)
        browse_btn.pack(side='left', padx=3)

        tk.Button(btn_frame, text='Insert & Close', font=FM, bg=GREEN, fg='#000', relief='flat',
                  bd=0, cursor='hand2', padx=10, command=insert_template).pack(side='right')


if __name__ == '__main__':
    app = App()
    app.mainloop()
