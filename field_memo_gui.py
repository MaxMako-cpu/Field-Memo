import os
import re
import glob
import shutil
import zipfile
import xml.etree.ElementTree as ET
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

# Report template — kept next to this script. Its 4 "Figure N" captions each
# already have a placeholder picture embedded right before them in the
# template; Generate Report swaps Figures 1-3's picture bytes for the 3 real
# event photos (in place, same zip path — no new library needed) and removes
# Figure 4's picture entirely (left blank for manual insertion later).
REPORT_TEMPLATE = 'IFR-PXGEO-OBN-013626-.docx'

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

        bottom_row = tk.Frame(self, bg=BG)
        bottom_row.grid(row=4, column=0, sticky='ew', padx=10, pady=(0, 8))
        bottom_row.columnconfigure(0, weight=1)  # spacer — pushes both buttons to the right

        tk.Button(bottom_row, text='GENERATE REPORT', font=FM, bg=YELLOW, fg='#000',
                  relief='flat', bd=0, cursor='hand2', padx=8, pady=3,
                  command=self._generate_report).grid(row=0, column=1, padx=(0, 6))
        tk.Button(bottom_row, text='CLR LOG', font=FM, bg=BORDER, fg=FG_DIM,
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
        except Exception as e:
            self._log(f'✗ Report generation failed: {e}', 'err')


if __name__ == '__main__':
    app = App()
    app.mainloop()
