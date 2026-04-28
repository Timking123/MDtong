"""GUI 界面 — tkinter 标签页架构（v2 完整优化版）"""

import os
import sys
import re
import subprocess
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from typing import Any

from .parser import parse_markdown
from .generator import SCRIPT_DIR, FORMAT_PRESETS
from .reverse import docx_to_markdown
from .services import ConversionCancelled, ConversionRequest, convert_request, convert_file
from .templates import list_templates, load_template_content
from .validator import format_issues, has_errors, validate_conversion_input
from .ai_polish import get_template_names, detect_doc_type
from .logging_config import get_logger
from .config import (
    load_config, save_config,
    DEFAULT_CONFIG, THEMES, APP_VERSION,
    MAX_RECENT, MAX_HISTORY, DRAFT_FILENAME, MARKDOWN_CHEATSHEET,
    AI_MODELS,
)

logger = get_logger('gui')

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# ── 语法高亮正则 ──

RE_HL_HEADING = re.compile(r'^(#{1,6})\s+(.+)$', re.MULTILINE)
RE_HL_BOLD = re.compile(r'\*\*(.+?)\*\*')
RE_HL_CODE_INLINE = re.compile(r'`([^`]+)`')
RE_HL_LINK = re.compile(r'\[([^\]]+)\]\([^)]+\)')
RE_HL_IMAGE = re.compile(r'!\[([^\]]*)\]\([^)]+\)')
RE_HL_FENCE = re.compile(r'^```\w*$', re.MULTILINE)
RE_HL_TASK = re.compile(r'^\s*[*\-+]\s+\[([ xX])\]\s+', re.MULTILINE)
RE_HL_LIST = re.compile(r'^\s*[*\-+]\s+', re.MULTILINE)
RE_HL_STRIKE = re.compile(r'~~.+?~~')


def _open_path(path):
    """跨平台打开文件或目录"""
    if sys.platform == 'win32':
        os.startfile(path)
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


def _get_ui_font():
    """根据平台返回合适的 UI 字体"""
    if sys.platform == 'win32':
        return 'Microsoft YaHei'
    elif sys.platform == 'darwin':
        return 'PingFang SC'
    return 'Noto Sans CJK SC'


def _get_mono_font():
    """根据平台返回合适的等宽字体"""
    return 'Consolas' if sys.platform == 'win32' else 'Menlo' if sys.platform == 'darwin' else 'Monospace'


UI_FONT = _get_ui_font()
MONO_FONT = _get_mono_font()

APP_PADDING = 14
CARD_PADDING = 10


class App:
    def __init__(self):
        if DND_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        self.root.title('MD通')
        self._cfg = load_config()
        self.style = ttk.Style(self.root)
        self._configure_base_style()

        saved_geo = self._cfg.get('window_geometry', '')
        if saved_geo:
            self.root.geometry(saved_geo)
        else:
            self.root.geometry('1100x700')
        self.root.minsize(800, 550)

        self.msg_queue = queue.Queue()
        self.output_path = None
        self._preview_timer = None
        self._current_save_path = None
        self._loading_editor_content = False
        self._cancel_event = threading.Event()
        self._sync_scroll_enabled = True
        self._find_bar_visible = False
        self._find_matches = []
        self._find_current = -1

        self._build_ui()
        self._bind_shortcuts()
        self._apply_theme(self._cfg.get('theme', 'light'))
        self._check_draft_recovery()
        self._start_auto_save()
        self._setup_dnd()

        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    # ── 通用轮询框架 ──

    def _poll_queue(self, handlers, interval=50):
        """通用消息队列轮询。handlers: {msg_type: callback}"""
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                handler = handlers.get(msg[0])
                if handler:
                    result = handler(msg)
                    if result == 'stop':
                        return
        except queue.Empty:
            pass
        self.root.after(interval, lambda: self._poll_queue(handlers, interval))

    def _configure_base_style(self):
        """初始化跨平台 ttk 主题，后续颜色由 _configure_theme_style 注入。"""
        try:
            self.style.theme_use('clam')
        except tk.TclError:
            pass
        self.root.option_add('*Font', (UI_FONT, 10))
        self.root.option_add('*TCombobox*Listbox.font', (UI_FONT, 10))

    def _theme_palette(self, theme_name=None):
        theme_var = getattr(self, 'theme_var', None)
        theme = theme_name or (theme_var.get() if theme_var is not None else 'light')
        base = THEMES.get(theme, THEMES['light'])
        defaults = {
            'window_bg': '#f4f7fb' if theme == 'light' else '#111827',
            'surface': '#ffffff' if theme == 'light' else '#1f2937',
            'surface_alt': '#eef4ff' if theme == 'light' else '#243244',
            'border': '#d8e2f0' if theme == 'light' else '#374151',
            'muted_fg': '#667085' if theme == 'light' else '#9ca3af',
            'accent': '#2563eb' if theme == 'light' else '#60a5fa',
            'accent_hover': '#1d4ed8' if theme == 'light' else '#3b82f6',
            'accent_fg': '#ffffff' if theme == 'light' else '#0f172a',
            'danger': '#dc2626' if theme == 'light' else '#f87171',
            'success': '#16a34a' if theme == 'light' else '#4ade80',
        }
        return {**defaults, **base}

    def _configure_theme_style(self, theme_name):
        """统一设置窗口、标签页、按钮、表格等控件视觉风格。"""
        p = self._theme_palette(theme_name)
        self.root.configure(bg=p['window_bg'])

        self.style.configure('.', font=(UI_FONT, 10), background=p['window_bg'], foreground=p['fg'])
        self.style.configure('TFrame', background=p['window_bg'])
        self.style.configure('Card.TFrame', background=p['surface'], relief='flat')
        self.style.configure('Hero.TFrame', background=p['surface_alt'])
        self.style.configure('TLabel', background=p['window_bg'], foreground=p['fg'])
        self.style.configure('Card.TLabel', background=p['surface'], foreground=p['fg'])
        self.style.configure('Hero.TLabel', background=p['surface_alt'], foreground=p['fg'])
        self.style.configure('Title.TLabel', background=p['surface_alt'], foreground=p['accent'], font=(UI_FONT, 18, 'bold'))
        self.style.configure('Subtitle.TLabel', background=p['surface_alt'], foreground=p['muted_fg'], font=(UI_FONT, 9))
        self.style.configure('Muted.TLabel', background=p['window_bg'], foreground=p['muted_fg'], font=(UI_FONT, 8))
        self.style.configure('Status.TLabel', background=p['window_bg'], foreground=p['muted_fg'])

        self.style.configure('TButton', padding=(12, 7), relief='flat', borderwidth=1)
        self.style.map(
            'TButton',
            background=[('active', p['surface_alt']), ('pressed', p['border'])],
            foreground=[('disabled', p['muted_fg'])],
        )
        self.style.configure('Accent.TButton', background=p['accent'], foreground=p['accent_fg'], borderwidth=0)
        self.style.map('Accent.TButton', background=[('active', p['accent_hover']), ('pressed', p['accent_hover'])])
        self.style.configure('Danger.TButton', background=p['danger'], foreground='#ffffff', borderwidth=0)
        self.style.map('Danger.TButton', background=[('active', p['danger']), ('pressed', p['danger'])])

        self.style.configure('TMenubutton', padding=(12, 7), background=p['surface'], foreground=p['fg'])
        self.style.configure('TCheckbutton', background=p['window_bg'], foreground=p['fg'])
        self.style.configure('TEntry', fieldbackground=p['surface'], foreground=p['fg'], bordercolor=p['border'], padding=6)
        self.style.configure('TCombobox', fieldbackground=p['surface'], foreground=p['fg'], bordercolor=p['border'], padding=4)
        self.style.configure('TNotebook', background=p['window_bg'], borderwidth=0, tabmargins=(0, 6, 0, 0))
        self.style.configure('TNotebook.Tab', padding=(16, 8), font=(UI_FONT, 10), background=p['surface_alt'], foreground=p['muted_fg'])
        self.style.map('TNotebook.Tab', background=[('selected', p['surface'])], foreground=[('selected', p['accent'])])
        self.style.configure('TLabelframe', background=p['surface'], bordercolor=p['border'], relief='solid')
        self.style.configure('TLabelframe.Label', background=p['surface'], foreground=p['accent'], font=(UI_FONT, 10, 'bold'))
        self.style.configure('Treeview', background=p['surface'], fieldbackground=p['surface'], foreground=p['fg'], rowheight=30, borderwidth=0)
        self.style.configure('Treeview.Heading', background=p['surface_alt'], foreground=p['fg'], font=(UI_FONT, 10, 'bold'), padding=8)
        self.style.map('Treeview', background=[('selected', p['select_bg'])], foreground=[('selected', p['fg'])])
        self.style.configure('TProgressbar', troughcolor=p['surface_alt'], background=p['accent'], bordercolor=p['border'], lightcolor=p['accent'], darkcolor=p['accent'])

    def _style_canvas(self, canvas):
        p = self._theme_palette()
        canvas.configure(bg=p['window_bg'], highlightthickness=0)

    def _style_popup(self, window):
        p = self._theme_palette()
        window.configure(bg=p['window_bg'])

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=APP_PADDING)
        main.pack(fill='both', expand=True)

        # 顶部栏
        top = ttk.Frame(main, style='Hero.TFrame', padding=(18, 14))
        top.pack(fill='x', pady=(0, 12))
        title_box = ttk.Frame(top, style='Hero.TFrame')
        title_box.pack(side='left', fill='x', expand=True)
        ttk.Label(
            title_box, text='MD通',
            style='Title.TLabel',
        ).pack(anchor='w')
        ttk.Label(
            title_box, text=f'Markdown → DOCX 智能转换工具  ·  v{APP_VERSION}',
            style='Subtitle.TLabel',
        ).pack(anchor='w', pady=(3, 0))

        # 深色模式切换
        self.theme_var = tk.StringVar(value=self._cfg.get('theme', 'light'))
        ttk.Checkbutton(
            top, text='深色模式',
            command=self._toggle_theme,
            variable=self.theme_var,
            onvalue='dark', offvalue='light',
        ).pack(side='right', padx=(12, 0))

        # 关于按钮
        ttk.Button(top, text='关于', command=self._show_about, width=7).pack(side='right', padx=5)
        # 语法帮助按钮
        ttk.Button(top, text='语法帮助', command=self._show_cheatsheet, width=10).pack(side='right', padx=5)

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill='both', expand=True)

        # 格式预设变量（标签页间共享）
        fs = self._cfg.get('format_settings', {})
        self.format_preset_var = tk.StringVar(value=fs.get('preset', '默认'))

        self._build_single_tab(self.notebook)
        self._build_batch_tab(self.notebook)
        self._build_history_tab(self.notebook)
        self._build_settings_tab(self.notebook)

    # ── 快捷键绑定 ──

    def _bind_shortcuts(self):
        self.root.bind('<Control-s>', lambda e: self._shortcut_save())
        self.root.bind('<Control-o>', lambda e: self._open_file())
        self.root.bind('<Control-n>', lambda e: self._clear())
        self.root.bind('<F5>', lambda e: self._start())
        self.root.bind('<Control-Shift-P>', lambda e: self._ai_polish())
        self.root.bind('<Control-Shift-p>', lambda e: self._ai_polish())
        self.root.bind('<Control-f>', lambda e: self._show_find())
        self.root.bind('<Control-h>', lambda e: self._show_find(replace=True))
        self.root.bind('<Escape>', lambda e: self._hide_find())
        self.root.bind('<Control-t>', lambda e: self._new_editor_tab())
        self.root.bind('<Control-w>', lambda e: self._close_editor_tab())

    def _shortcut_save(self):
        self._save_current_tab()

    def _get_current_tab_index(self):
        return self.editor_notebook.index(self.editor_notebook.select())

    def _get_current_tab(self):
        return self.editor_tabs[self._get_current_tab_index()]

    def _sync_current_path_from_tab(self):
        tab = self._get_current_tab()
        self._current_save_path = tab.get('path')
        return self._current_save_path

    def _set_tab_dirty(self, tab=None, dirty=True):
        if tab is None:
            tab = self._get_current_tab()
        tab['dirty'] = dirty
        self._refresh_editor_tab_title(tab)

    def _refresh_editor_tab_title(self, tab):
        if not hasattr(self, 'editor_notebook'):
            return
        try:
            idx = self.editor_tabs.index(tab)
        except ValueError:
            return
        title = tab.get('title') or f'文档 {idx + 1}'
        if tab.get('dirty'):
            title += '*'
        self.editor_notebook.tab(tab['frame'], text=title)

    def _set_current_tab_path(self, path):
        tab = self._get_current_tab()
        tab['path'] = path
        tab['title'] = os.path.basename(path) if path else tab.get('title')
        self._current_save_path = path
        self._refresh_editor_tab_title(tab)

    def _save_current_tab(self, save_as=False):
        content = self._get_current_text().strip()
        if not content:
            self.status.set('没有内容可保存')
            return False

        tab = self._get_current_tab()
        path = None if save_as else tab.get('path')
        if not path:
            init_dir = self._cfg.get('last_save_dir', '') or SCRIPT_DIR
            path = filedialog.asksaveasfilename(
                defaultextension='.md',
                filetypes=[('Markdown', '*.md'), ('文本文件', '*.txt')],
                initialdir=init_dir,
            )
            if not path:
                return False

        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        self._set_current_tab_path(path)
        self._set_tab_dirty(tab, False)
        self._cfg['last_save_dir'] = os.path.dirname(path)
        self._add_to_recent(path)
        self.status.set(f'已保存: {os.path.basename(path)}')
        return True

    # ── 快捷格式插入 ──

    def _insert_format(self, prefix, suffix, placeholder='', line_start=False):
        editor = self._get_current_editor()
        try:
            sel_text = editor.get('sel.first', 'sel.last')
            start = editor.index('sel.first')
            editor.delete('sel.first', 'sel.last')
        except tk.TclError:
            sel_text = ''
            start = editor.index('insert')

        text = sel_text or placeholder

        if line_start:
            line = start.split('.')[0]
            col = int(start.split('.')[1])
            if col > 0:
                prefix = '\n' + prefix

        editor.insert(start, prefix + text + suffix)

        if not sel_text and placeholder:
            prefix_idx = editor.index(f'{start}+{len(prefix)}c')
            end_idx = editor.index(f'{prefix_idx}+{len(placeholder)}c')
            editor.tag_remove('sel', '1.0', 'end')
            editor.tag_add('sel', prefix_idx, end_idx)
            editor.mark_set('insert', end_idx)

        editor.focus_set()

    def _insert_table_template(self):
        self._open_table_editor()

    def _open_table_editor(self, initial_rows=3, initial_cols=3):
        """打开表格可视化编辑器窗口"""
        win = tk.Toplevel(self.root)
        win.title('表格编辑器')
        win.geometry('700x500')
        win.transient(self.root)
        self._style_popup(win)

        rows_var = tk.IntVar(value=initial_rows)
        cols_var = tk.IntVar(value=initial_cols)
        cells = []

        ctrl = ttk.Frame(win)
        ctrl.pack(fill='x', padx=10, pady=5)
        ttk.Label(ctrl, text='行数:').pack(side='left')
        ttk.Spinbox(ctrl, from_=1, to=50, textvariable=rows_var, width=4).pack(side='left', padx=(0, 10))
        ttk.Label(ctrl, text='列数:').pack(side='left')
        ttk.Spinbox(ctrl, from_=1, to=20, textvariable=cols_var, width=4).pack(side='left', padx=(0, 10))
        ttk.Button(ctrl, text='更新网格', command=lambda: rebuild()).pack(side='left', padx=5)

        canvas_frame = ttk.Frame(win)
        canvas_frame.pack(fill='both', expand=True, padx=10, pady=5)
        canvas = tk.Canvas(canvas_frame)
        self._style_canvas(canvas)
        v_scroll = ttk.Scrollbar(canvas_frame, orient='vertical', command=canvas.yview)
        h_scroll = ttk.Scrollbar(canvas_frame, orient='horizontal', command=canvas.xview)
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.pack(side='right', fill='y')
        h_scroll.pack(side='bottom', fill='x')
        canvas.pack(side='left', fill='both', expand=True)

        grid_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=grid_frame, anchor='nw')

        def rebuild():
            for w in grid_frame.winfo_children():
                w.destroy()
            cells.clear()

            r = rows_var.get()
            c = cols_var.get()

            ttk.Label(grid_frame, text='', width=3).grid(row=0, column=0)
            for j in range(c):
                ttk.Label(grid_frame, text=f'列{j+1}', font=(UI_FONT, 9, 'bold')).grid(row=0, column=j+1, padx=2, pady=2)

            for i in range(r):
                label = '表头' if i == 0 else f'行{i}'
                ttk.Label(grid_frame, text=label, font=(UI_FONT, 9)).grid(row=i+1, column=0, padx=2, pady=2)
                row_cells = []
                for j in range(c):
                    var = tk.StringVar(value=f'列{j+1}' if i == 0 else '')
                    entry = ttk.Entry(grid_frame, textvariable=var, width=15)
                    entry.grid(row=i+1, column=j+1, padx=2, pady=2)
                    row_cells.append(var)
                cells.append(row_cells)

            grid_frame.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox('all'))

        rebuild()

        def insert_table():
            if not cells:
                return
            lines = []
            header = [v.get() or '' for v in cells[0]]
            lines.append('| ' + ' | '.join(header) + ' |')
            lines.append('| ' + ' | '.join(['---'] * len(header)) + ' |')
            for row in cells[1:]:
                vals = [v.get() or '' for v in row]
                lines.append('| ' + ' | '.join(vals) + ' |')
            md = '\n'.join(lines) + '\n'
            editor = self._get_current_editor()
            pos = editor.index('insert')
            col = int(pos.split('.')[1])
            if col > 0:
                md = '\n' + md
            editor.insert('insert', md)
            editor.focus_set()
            win.destroy()

        bottom = ttk.Frame(win)
        bottom.pack(fill='x', padx=10, pady=(0, 10))
        ttk.Button(bottom, text='插入表格', command=insert_table).pack(side='left', padx=5)
        ttk.Button(bottom, text='取消', command=win.destroy).pack(side='right', padx=5)

    # ── 多标签编辑管理 ──

    def _get_current_text(self):
        """获取当前活动编辑器的文本"""
        return self._get_current_tab()['text'].get('1.0', 'end')

    def _get_current_editor(self):
        """获取当前活动的 Text widget"""
        return self._get_current_tab()['text']

    def _new_editor_tab(self):
        tab_info = self._create_editor_tab(self.editor_notebook)
        tab_info['title'] = f'文档 {len(self.editor_tabs) + 1}'
        self.editor_notebook.add(tab_info['frame'], text=tab_info['title'])
        self.editor_tabs.append(tab_info)
        self.editor_notebook.select(tab_info['frame'])
        self._sync_current_path_from_tab()

    def _close_editor_tab(self):
        if len(self.editor_tabs) <= 1:
            return
        idx = self.editor_notebook.index(self.editor_notebook.select())
        tab = self.editor_tabs[idx]
        if tab.get('dirty'):
            answer = messagebox.askyesnocancel('未保存的更改', f'{tab.get("title", "当前文档")} 有未保存内容，是否保存？')
            if answer is None:
                return
            if answer and not self._save_current_tab():
                return
        self.editor_notebook.forget(idx)
        self.editor_tabs.pop(idx)
        self._sync_current_path_from_tab()

    def _create_editor_tab(self, parent):
        frame = ttk.Frame(parent)

        line_numbers = tk.Text(
            frame, width=4, padx=4, pady=0, takefocus=0,
            border=0, state='disabled', wrap='none',
            font=(MONO_FONT, 10), fg='#999999', bg='#f0f0f0',
        )
        line_numbers.pack(side='left', fill='y')

        text = tk.Text(
            frame, wrap='word', font=(UI_FONT, 10), undo=True,
            padx=12, pady=10, relief='flat', borderwidth=0,
        )
        sb = ttk.Scrollbar(frame, command=text.yview)
        text.configure(yscrollcommand=sb.set)
        text.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        text.tag_configure('hl_heading', foreground='#1a56db', font=(UI_FONT, 10, 'bold'))
        text.tag_configure('hl_bold', font=(UI_FONT, 10, 'bold'))
        text.tag_configure('hl_code', foreground='#c7254e', background='#f9f2f4', font=(MONO_FONT, 10))
        text.tag_configure('hl_link', foreground='#0563C1', underline=True)
        text.tag_configure('hl_image', foreground='#6f42c1')
        text.tag_configure('hl_fence', foreground='#999999', background='#f0f0f0')
        text.tag_configure('hl_task', foreground='#228B22')
        text.tag_configure('hl_list', foreground='#e36209')
        text.tag_configure('hl_strike', foreground='#999999', overstrike=True)
        text.tag_configure('find_highlight', background='#ffff00', foreground='#000000')
        text.tag_configure('find_current', background='#ff8c00', foreground='#ffffff')

        text.bind('<<Modified>>', self._on_text_modified)
        text.bind('<Control-v>', lambda e: self._on_paste_image(text))

        text.bind('<MouseWheel>', lambda e: self._on_editor_scroll())
        text.bind('<B1-Motion>', lambda e: self.root.after(10, self._on_editor_scroll))

        def _sync_line_numbers_scroll(*args):
            line_numbers.yview_moveto(text.yview()[0])

        def _scroll_editor(*args):
            text.yview(*args)
            _sync_line_numbers_scroll()

        text.bind('<KeyRelease>', lambda e: self._update_line_numbers(text, line_numbers))
        text.bind('<MouseWheel>', lambda e: self.root.after(10, lambda: _sync_line_numbers_scroll()), add='+')
        sb.configure(command=_scroll_editor)

        return {'frame': frame, 'text': text, 'line_numbers': line_numbers, 'path': None, 'title': '', 'dirty': False}

    def _update_line_numbers(self, text_widget=None, ln_widget=None):
        if text_widget is None:
            idx = self.editor_notebook.index(self.editor_notebook.select())
            tab = self.editor_tabs[idx]
            text_widget = tab['text']
            ln_widget = tab['line_numbers']
        if ln_widget is None:
            return

        ln_widget.configure(state='normal')
        ln_widget.delete('1.0', 'end')
        line_count = int(text_widget.index('end-1c').split('.')[0])
        lines = '\n'.join(str(i) for i in range(1, line_count + 1))
        ln_widget.insert('1.0', lines)
        ln_widget.configure(state='disabled')
        ln_widget.yview_moveto(text_widget.yview()[0])

    # ── 最近文件管理 ──

    def _on_paste_image(self, text_widget):
        """尝试从剪贴板粘贴图片，成功返回 'break' 阻止默认粘贴"""
        try:
            from PIL import ImageGrab, Image
            img = ImageGrab.grabclipboard()
            if not isinstance(img, Image.Image):
                return
            images_dir = os.path.join(SCRIPT_DIR, 'images')
            os.makedirs(images_dir, exist_ok=True)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            filename = f'paste_{ts}.png'
            filepath = os.path.join(images_dir, filename)
            img.save(filepath, 'PNG')
            rel_path = f'images/{filename}'
            md_img = f'![粘贴图片]({rel_path})'
            text_widget.insert('insert', md_img)
            self.status.set(f'已粘贴图片: {filename}')
            return 'break'
        except ImportError:
            return
        except Exception:
            return

    def _add_to_recent(self, path):
        recent = self._cfg.get('recent_files', [])
        if path in recent:
            recent.remove(path)
        recent.insert(0, path)
        self._cfg['recent_files'] = recent[:MAX_RECENT]
        save_config(self._cfg)
        self._update_recent_menu()

    def _update_recent_menu(self):
        if not hasattr(self, 'recent_menu'):
            return
        self.recent_menu.delete(0, 'end')
        recent = self._cfg.get('recent_files', [])
        if not recent:
            self.recent_menu.add_command(label='（无最近文件）', state='disabled')
            return
        for path in recent:
            name = os.path.basename(path)
            self.recent_menu.add_command(
                label=name, command=lambda p=path: self._open_recent(p),
            )

    def _open_recent(self, path):
        if not os.path.isfile(path):
            self.status.set(f'文件不存在: {path}')
            return
        self._load_file_to_editor(path)

    def _load_file_to_editor(self, path):
        for enc in ('utf-8-sig', 'utf-8', 'gbk'):
            try:
                with open(path, 'r', encoding=enc) as f:
                    content = f.read()
                editor = self._get_current_editor()
                self._loading_editor_content = True
                editor.delete('1.0', 'end')
                editor.insert('1.0', content)
                editor.edit_modified(False)
                self._loading_editor_content = False
                self._set_current_tab_path(path)
                self._set_tab_dirty(self._get_current_tab(), False)
                self._add_to_recent(path)
                self.status.set(f'已打开: {os.path.basename(path)}')
                return True
            except UnicodeDecodeError:
                continue
            finally:
                self._loading_editor_content = False
        self.status.set('错误：无法读取文件编码')
        return False

    # ── 深色模式 ──

    def _toggle_theme(self):
        theme = self.theme_var.get()
        self._apply_theme(theme)
        self._cfg['theme'] = theme
        save_config(self._cfg)

    def _apply_theme(self, theme_name):
        t = THEMES.get(theme_name, THEMES['light'])
        p = self._theme_palette(theme_name)
        self._configure_theme_style(theme_name)

        for tab_info in getattr(self, 'editor_tabs', []):
            text = tab_info['text']
            text.configure(
                bg=t['bg'], fg=t['fg'],
                insertbackground=t['insert_bg'],
                selectbackground=t['select_bg'],
                highlightbackground=p['border'], highlightcolor=p['accent'],
            )
            ln = tab_info.get('line_numbers')
            if ln:
                ln.configure(bg=p['surface_alt'], fg=p['muted_fg'])

        if hasattr(self, 'preview'):
            self.preview.configure(
                bg=t['preview_bg'], fg=t['fg'],
                insertbackground=t['insert_bg'],
                selectbackground=t['select_bg'],
                highlightbackground=p['border'], highlightcolor=p['accent'],
            )
            self.preview.tag_configure('title', foreground=t['heading_fg'])
            self.preview.tag_configure('section', foreground=t['heading_fg'])
            self.preview.tag_configure('sub_header', foreground=t['bold_fg'])
            self.preview.tag_configure('label', foreground=t['bold_fg'])
            self.preview.tag_configure('body', foreground=t['fg'])
            self.preview.tag_configure('table', background=t['code_bg'], foreground=t['fg'])
            self.preview.tag_configure('code_block', background=t['code_bg'], foreground=t['fg'])
            self.preview.tag_configure('task_done', foreground=t['task_fg'])
            self.preview.tag_configure('link', foreground=t['link_fg'])

        for menu_name in ('recent_menu',):
            menu = getattr(self, menu_name, None)
            if menu:
                menu.configure(bg=p['surface'], fg=p['fg'], activebackground=p['select_bg'], activeforeground=p['fg'])

        for canvas_name in ('settings_canvas',):
            canvas = getattr(self, canvas_name, None)
            if canvas:
                self._style_canvas(canvas)

        for tab_info in getattr(self, 'editor_tabs', []):
            self._apply_syntax_highlight(tab_info['text'])

    # ── 单个转换标签页 ──

    def _build_single_tab(self, notebook):
        tab = ttk.Frame(notebook, padding=APP_PADDING)
        notebook.add(tab, text='单个转换')

        # 按钮栏
        btn = ttk.Frame(tab)
        btn.pack(fill='x', pady=(0, 5))
        ttk.Button(btn, text='粘贴剪贴板', command=self._paste).pack(side='left', padx=(0, 5))
        ttk.Button(btn, text='打开文件', command=self._open_file).pack(side='left', padx=(0, 5))
        ttk.Button(btn, text='DOCX → Markdown', command=self._import_docx).pack(side='left', padx=(0, 5))
        self.btn_ai = ttk.Button(btn, text='AI 润色', command=self._ai_polish)
        self.btn_ai.pack(side='left', padx=(0, 5))
        self.ai_tpl_var = tk.StringVar(value='会议纪要')
        ai_tpl_combo = ttk.Combobox(
            btn, textvariable=self.ai_tpl_var, values=get_template_names(),
            state='readonly', width=10,
        )
        ai_tpl_combo.pack(side='left', padx=(0, 5))
        ttk.Button(btn, text='清空', command=self._clear).pack(side='left', padx=(0, 5))

        # 取消按钮（默认隐藏）
        self.btn_cancel = ttk.Button(btn, text='取消', command=self._cancel_operation, style='Danger.TButton')

        # 最近文件下拉
        self.recent_menu = tk.Menu(self.root, tearoff=False)
        recent_btn = ttk.Menubutton(btn, text='最近文件')
        recent_btn['menu'] = self.recent_menu
        recent_btn.pack(side='left', padx=(0, 5))
        self._update_recent_menu()

        # 模板选择
        tpl_frame = ttk.Frame(tab)
        tpl_frame.pack(fill='x', pady=(0, 5))
        ttk.Label(tpl_frame, text='模板:').pack(side='left')
        self.tpl_var = tk.StringVar()
        tpl_names = list_templates()
        self.tpl_combo = ttk.Combobox(
            tpl_frame, textvariable=self.tpl_var, values=tpl_names,
            state='readonly', width=25,
        )
        self.tpl_combo.pack(side='left', padx=5)
        self.tpl_combo.bind('<<ComboboxSelected>>', self._on_template_selected)

        ttk.Label(tpl_frame, text='格式:').pack(side='left', padx=(10, 0))
        format_combo = ttk.Combobox(
            tpl_frame, textvariable=self.format_preset_var,
            values=['默认', '公文', '自定义'],
            state='readonly', width=10,
        )
        format_combo.pack(side='left', padx=5)

        # 快捷键提示
        ttk.Label(
            tpl_frame,
            text='Ctrl+S 保存 | Ctrl+O 打开 | F5 转换 | Ctrl+F 查找 | Ctrl+T 新标签',
            style='Muted.TLabel',
        ).pack(side='right')

        # 快捷格式工具栏
        fmt_bar = ttk.Frame(tab)
        fmt_bar.pack(fill='x', pady=(0, 3))
        for label, callback in [
            ('B', lambda: self._insert_format('**', '**', '加粗文本')),
            ('I', lambda: self._insert_format('*', '*', '斜体文本')),
            ('S', lambda: self._insert_format('~~', '~~', '删除文本')),
            ('H1', lambda: self._insert_format('# ', '', '一级标题', line_start=True)),
            ('H2', lambda: self._insert_format('## ', '', '二级标题', line_start=True)),
            ('H3', lambda: self._insert_format('### ', '', '三级标题', line_start=True)),
            ('链接', lambda: self._insert_format('[', '](url)', '链接文本')),
            ('图片', lambda: self._insert_format('![', '](path)', '图片描述')),
            ('代码', lambda: self._insert_format('`', '`', '代码')),
            ('代码块', lambda: self._insert_format('```\n', '\n```', '代码内容', line_start=True)),
            ('列表', lambda: self._insert_format('- ', '', '列表项', line_start=True)),
            ('编号', lambda: self._insert_format('1. ', '', '列表项', line_start=True)),
            ('任务', lambda: self._insert_format('- [ ] ', '', '待办事项', line_start=True)),
            ('表格', lambda: self._insert_table_template()),
            ('引用', lambda: self._insert_format('> ', '', '引用文本', line_start=True)),
            ('分隔线', lambda: self._insert_format('\n---\n', '', '', line_start=True)),
        ]:
            ttk.Button(fmt_bar, text=label, command=callback, width=max(4, len(label) + 1)).pack(side='left', padx=2, pady=2)

        # 查找替换栏（默认隐藏）
        self._build_find_bar(tab)

        # 底部固定区域
        bottom = ttk.Frame(tab)
        bottom.pack(side='bottom', fill='x', pady=(5, 0))

        # 左右分栏
        paned = ttk.PanedWindow(tab, orient='horizontal')
        paned.pack(fill='both', expand=True, pady=5)

        # 左侧：大纲 + 编辑器
        left = ttk.LabelFrame(paned, text='Markdown 编辑', padding=CARD_PADDING)
        paned.add(left, weight=1)

        left_paned = ttk.PanedWindow(left, orient='horizontal')
        left_paned.pack(fill='both', expand=True)

        # 大纲面板
        outline_frame = ttk.Frame(left_paned)
        left_paned.add(outline_frame, weight=0)
        ttk.Label(outline_frame, text='大纲', font=(UI_FONT, 9, 'bold')).pack(anchor='w', padx=2)
        self.outline_tree = ttk.Treeview(outline_frame, show='tree', selectmode='browse')
        outline_sb = ttk.Scrollbar(outline_frame, command=self.outline_tree.yview)
        self.outline_tree.configure(yscrollcommand=outline_sb.set)
        self.outline_tree.pack(side='left', fill='both', expand=True)
        outline_sb.pack(side='right', fill='y')
        self.outline_tree.column('#0', width=160, minwidth=100)
        self.outline_tree.bind('<<TreeviewSelect>>', self._on_outline_click)

        # 编辑器区域
        editor_frame = ttk.Frame(left_paned)
        left_paned.add(editor_frame, weight=1)

        self.editor_notebook = ttk.Notebook(editor_frame)
        self.editor_notebook.pack(fill='both', expand=True)
        self.editor_tabs = []

        first_tab = self._create_editor_tab(self.editor_notebook)
        self.editor_notebook.add(first_tab['frame'], text='文档 1')
        self.editor_tabs.append(first_tab)

        # 为兼容性保留 self.text 引用
        self.text = first_tab['text']

        # 右侧：实时预览
        right = ttk.LabelFrame(paned, text='排版预览', padding=CARD_PADDING)
        paned.add(right, weight=1)
        self.preview = tk.Text(
            right, wrap='word', font=(UI_FONT, 10), state='disabled',
            padx=12, pady=10, relief='flat', borderwidth=0,
        )
        sb_right = ttk.Scrollbar(right, command=self.preview.yview)
        self.preview.configure(yscrollcommand=sb_right.set)
        self.preview.pack(side='left', fill='both', expand=True)
        sb_right.pack(side='right', fill='y')

        self.preview.tag_configure('title', justify='center', font=(UI_FONT, 14, 'bold'))
        self.preview.tag_configure('section', font=(UI_FONT, 11, 'bold'), foreground='#1a56db')
        self.preview.tag_configure('sub_header', font=(UI_FONT, 10, 'bold'))
        self.preview.tag_configure('label', font=(UI_FONT, 10, 'bold'))
        self.preview.tag_configure('body', font=(UI_FONT, 10))
        self.preview.tag_configure('table', font=(MONO_FONT, 9), background='#f5f5f5')
        self.preview.tag_configure('empty', font=(UI_FONT, 6))
        self.preview.tag_configure('code_block', font=(MONO_FONT, 9), background='#f0f0f0')
        self.preview.tag_configure('task_done', foreground='#228B22')
        self.preview.tag_configure('task_undone', foreground='#666666')
        self.preview.tag_configure('link', foreground='#0563C1', underline=True)
        self.preview.tag_configure('list_item', font=(UI_FONT, 10), lmargin1=20, lmargin2=35)
        self.preview.tag_configure('image_placeholder', foreground='#6f42c1', font=(UI_FONT, 10, 'italic'))

        # 输出目录
        of = ttk.Frame(bottom)
        of.pack(fill='x', pady=2)
        ttk.Label(of, text='输出目录:').pack(side='left')
        output_dir = self._cfg.get('output_dir', '') or self._cfg.get('last_save_dir', '') or SCRIPT_DIR
        self.out_dir_var = tk.StringVar(value=output_dir)
        ttk.Entry(of, textvariable=self.out_dir_var, state='readonly').pack(
            side='left', fill='x', expand=True, padx=5,
        )
        ttk.Button(of, text='选择目录', command=self._browse, width=8).pack(side='right')
        self.also_pdf = tk.BooleanVar(value=False)
        ttk.Checkbutton(of, text='同时生成 PDF', variable=self.also_pdf).pack(side='right', padx=5)
        ttk.Button(of, text='导出 HTML', command=self._export_html, width=8).pack(side='right', padx=5)

        # 进度条
        self.pbar_var = tk.DoubleVar(value=0)
        self.pbar = ttk.Progressbar(bottom, variable=self.pbar_var, maximum=100)
        self.pbar.pack(fill='x', pady=2)

        # 状态栏 + 统计
        status_frame = ttk.Frame(bottom)
        status_frame.pack(fill='x')
        self.status = tk.StringVar(value='请粘贴或输入 Markdown 内容，然后点击「开始转换」（支持拖拽文件）')
        ttk.Label(status_frame, textvariable=self.status, wraplength=500, style='Status.TLabel').pack(side='left')
        self.stats_var = tk.StringVar(value='字符: 0 | 字数: 0 | 行数: 0')
        ttk.Label(
            status_frame, textvariable=self.stats_var,
            style='Muted.TLabel',
        ).pack(side='right')

        # 操作按钮
        bf = ttk.Frame(bottom)
        bf.pack(pady=4)
        self.btn_convert = ttk.Button(bf, text='开始转换', command=self._start, style='Accent.TButton')
        self.btn_convert.pack(side='left', padx=5)
        self.btn_open = ttk.Button(bf, text='打开文件', command=self._open_result, state='disabled')
        self.btn_open.pack(side='left', padx=5)
        self.btn_dir = ttk.Button(bf, text='打开目录', command=self._open_dir, state='disabled')
        self.btn_dir.pack(side='left', padx=5)
        self.btn_pdf = ttk.Button(bf, text='导出 PDF', command=self._export_pdf, state='disabled')
        self.btn_pdf.pack(side='left', padx=5)

    # ── 查找替换栏 ──

    def _build_find_bar(self, parent):
        self.find_frame = ttk.Frame(parent)

        row1 = ttk.Frame(self.find_frame)
        row1.pack(fill='x', pady=2)
        ttk.Label(row1, text='查找:').pack(side='left')
        self.find_var = tk.StringVar()
        self.find_entry = ttk.Entry(row1, textvariable=self.find_var, width=30)
        self.find_entry.pack(side='left', padx=5)
        self.find_entry.bind('<Return>', lambda e: self._find_next())
        ttk.Button(row1, text='上一个', command=self._find_prev, width=6).pack(side='left', padx=2)
        ttk.Button(row1, text='下一个', command=self._find_next, width=6).pack(side='left', padx=2)
        self.find_count_var = tk.StringVar(value='')
        ttk.Label(row1, textvariable=self.find_count_var, style='Muted.TLabel').pack(side='left', padx=5)
        ttk.Button(row1, text='✕', command=self._hide_find, width=3).pack(side='right')

        self.replace_frame = ttk.Frame(self.find_frame)
        row2 = ttk.Frame(self.replace_frame)
        row2.pack(fill='x', pady=2)
        ttk.Label(row2, text='替换:').pack(side='left')
        self.replace_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.replace_var, width=30).pack(side='left', padx=5)
        ttk.Button(row2, text='替换', command=self._replace, width=6).pack(side='left', padx=2)
        ttk.Button(row2, text='全部替换', command=self._replace_all, width=8).pack(side='left', padx=2)

    def _show_find(self, replace=False):
        if not self._find_bar_visible:
            self.find_frame.pack(fill='x', pady=(0, 5), after=self.find_frame.master.winfo_children()[1])
            self._find_bar_visible = True

        if replace:
            self.replace_frame.pack(fill='x')
        else:
            self.replace_frame.pack_forget()

        self.find_entry.focus_set()
        sel = ''
        try:
            editor = self._get_current_editor()
            sel = editor.get('sel.first', 'sel.last')
        except tk.TclError:
            pass
        if sel:
            self.find_var.set(sel)
        self.find_entry.select_range(0, 'end')

    def _hide_find(self):
        if self._find_bar_visible:
            self.find_frame.pack_forget()
            self._find_bar_visible = False
            editor = self._get_current_editor()
            editor.tag_remove('find_highlight', '1.0', 'end')
            editor.tag_remove('find_current', '1.0', 'end')
            self._find_matches = []
            self._find_current = -1

    def _find_all(self):
        editor = self._get_current_editor()
        editor.tag_remove('find_highlight', '1.0', 'end')
        editor.tag_remove('find_current', '1.0', 'end')
        self._find_matches = []
        self._find_current = -1

        query = self.find_var.get()
        if not query:
            self.find_count_var.set('')
            return

        start = '1.0'
        while True:
            pos = editor.search(query, start, stopindex='end', nocase=True)
            if not pos:
                break
            end = f'{pos}+{len(query)}c'
            self._find_matches.append((pos, end))
            editor.tag_add('find_highlight', pos, end)
            start = end

        count = len(self._find_matches)
        self.find_count_var.set(f'{count} 个结果' if count else '无结果')

    def _find_next(self):
        self._find_all()
        if not self._find_matches:
            return
        editor = self._get_current_editor()
        self._find_current = (self._find_current + 1) % len(self._find_matches)
        self._highlight_current_match(editor)

    def _find_prev(self):
        self._find_all()
        if not self._find_matches:
            return
        editor = self._get_current_editor()
        self._find_current = (self._find_current - 1) % len(self._find_matches)
        self._highlight_current_match(editor)

    def _highlight_current_match(self, editor):
        editor.tag_remove('find_current', '1.0', 'end')
        if 0 <= self._find_current < len(self._find_matches):
            pos, end = self._find_matches[self._find_current]
            editor.tag_add('find_current', pos, end)
            editor.see(pos)
            n = self._find_current + 1
            total = len(self._find_matches)
            self.find_count_var.set(f'{n}/{total}')

    def _replace(self):
        if not self._find_matches or self._find_current < 0:
            self._find_next()
            return
        editor = self._get_current_editor()
        pos, end = self._find_matches[self._find_current]
        editor.delete(pos, end)
        editor.insert(pos, self.replace_var.get())
        self._find_all()
        if self._find_matches:
            self._find_current = min(self._find_current, len(self._find_matches) - 1)
            self._highlight_current_match(editor)

    def _replace_all(self):
        query = self.find_var.get()
        replacement = self.replace_var.get()
        if not query:
            return
        editor = self._get_current_editor()
        content = editor.get('1.0', 'end')
        new_content = content.replace(query, replacement)
        if new_content != content:
            editor.delete('1.0', 'end')
            editor.insert('1.0', new_content.rstrip('\n'))
            count = content.count(query)
            self.status.set(f'已替换 {count} 处')
        self._find_all()

    # ── 批量转换标签页 ──

    def _build_batch_tab(self, notebook):
        tab = ttk.Frame(notebook, padding=APP_PADDING)
        notebook.add(tab, text='批量转换')

        # 按钮栏
        btn = ttk.Frame(tab)
        btn.pack(fill='x', pady=(0, 5))
        ttk.Button(btn, text='添加文件', command=self._batch_add).pack(side='left', padx=(0, 5))
        ttk.Button(btn, text='移除选中', command=self._batch_remove).pack(side='left', padx=(0, 5))
        ttk.Button(btn, text='清空列表', command=self._batch_clear).pack(side='left')

        # 取消按钮
        self.btn_batch_cancel = ttk.Button(btn, text='取消', command=self._cancel_operation, style='Danger.TButton')

        # 文件列表
        list_frame = ttk.Frame(tab)
        list_frame.pack(fill='both', expand=True, pady=5)

        columns = ('file', 'status')
        self.batch_tree = ttk.Treeview(list_frame, columns=columns, show='headings', selectmode='extended')
        self.batch_tree.heading('file', text='文件路径')
        self.batch_tree.heading('status', text='状态')
        self.batch_tree.column('file', width=500)
        self.batch_tree.column('status', width=120, anchor='center')
        batch_sb = ttk.Scrollbar(list_frame, command=self.batch_tree.yview)
        self.batch_tree.configure(yscrollcommand=batch_sb.set)
        self.batch_tree.pack(side='left', fill='both', expand=True)
        batch_sb.pack(side='right', fill='y')

        # 输出目录
        out_frame = ttk.Frame(tab)
        out_frame.pack(fill='x', pady=5)
        ttk.Label(out_frame, text='输出目录:').pack(side='left')
        batch_out_dir = self._cfg.get('batch_output_dir', '') or SCRIPT_DIR
        self.batch_out_var = tk.StringVar(value=batch_out_dir)
        ttk.Entry(out_frame, textvariable=self.batch_out_var, state='readonly').pack(
            side='left', fill='x', expand=True, padx=5,
        )
        ttk.Button(out_frame, text='浏览', command=self._batch_browse_dir, width=6).pack(side='right')

        # 选项
        opt_frame = ttk.Frame(tab)
        opt_frame.pack(fill='x', pady=2)
        self.batch_also_pdf = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt_frame, text='同时生成 PDF', variable=self.batch_also_pdf).pack(side='left')

        # 进度条
        self.batch_pbar_var = tk.DoubleVar(value=0)
        self.batch_pbar = ttk.Progressbar(tab, variable=self.batch_pbar_var, maximum=100)
        self.batch_pbar.pack(fill='x', pady=5)

        # 状态
        self.batch_status = tk.StringVar(value='添加 Markdown / DOCX 文件后点击「批量转换」（支持拖拽文件）')
        ttk.Label(tab, textvariable=self.batch_status, wraplength=720, style='Status.TLabel').pack()

        # 操作按钮
        bf = ttk.Frame(tab)
        bf.pack(pady=8)
        self.btn_batch_convert = ttk.Button(bf, text='批量转换', command=self._batch_start, style='Accent.TButton')
        self.btn_batch_convert.pack(side='left', padx=5)
        self.btn_batch_open_dir = ttk.Button(bf, text='打开输出目录', command=self._batch_open_dir, state='disabled')
        self.btn_batch_open_dir.pack(side='left', padx=5)

    # ── 导出历史标签页 ──

    def _build_history_tab(self, notebook):
        tab = ttk.Frame(notebook, padding=APP_PADDING)
        notebook.add(tab, text='历史记录')

        # 按钮栏
        btn = ttk.Frame(tab)
        btn.pack(fill='x', pady=(0, 5))
        ttk.Button(btn, text='刷新', command=self._refresh_history).pack(side='left', padx=(0, 5))
        ttk.Button(btn, text='清空历史', command=self._clear_history).pack(side='left')

        columns = ('time', 'input', 'output')
        self.history_tree = ttk.Treeview(tab, columns=columns, show='headings', selectmode='browse')
        self.history_tree.heading('time', text='时间')
        self.history_tree.heading('input', text='输入')
        self.history_tree.heading('output', text='输出文件')
        self.history_tree.column('time', width=150)
        self.history_tree.column('input', width=300)
        self.history_tree.column('output', width=350)
        history_sb = ttk.Scrollbar(tab, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=history_sb.set)
        self.history_tree.pack(side='left', fill='both', expand=True)
        history_sb.pack(side='right', fill='y')

        self.history_tree.bind('<Double-1>', self._on_history_double_click)

        self._refresh_history()

    def _refresh_history(self):
        for iid in self.history_tree.get_children():
            self.history_tree.delete(iid)
        history = self._cfg.get('export_history', [])
        for entry in reversed(history):
            self.history_tree.insert('', 'end', values=(
                entry.get('time', ''),
                entry.get('input', ''),
                entry.get('output', ''),
            ))

    def _add_to_history(self, input_desc, output_path):
        history = self._cfg.get('export_history', [])
        history.append({
            'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'input': input_desc,
            'output': output_path,
        })
        self._cfg['export_history'] = history[-MAX_HISTORY:]
        save_config(self._cfg)

    def _clear_history(self):
        if messagebox.askyesno('确认', '确定要清空所有导出历史记录吗？'):
            self._cfg['export_history'] = []
            save_config(self._cfg)
            self._refresh_history()

    def _on_history_double_click(self, event):
        sel = self.history_tree.selection()
        if not sel:
            return
        values = self.history_tree.item(sel[0], 'values')
        output = values[2] if len(values) > 2 else ''
        if output and os.path.isfile(output):
            _open_path(output)

    # ── 语法高亮 ──

    def _apply_syntax_highlight(self, text_widget=None):
        if text_widget is None:
            text_widget = self._get_current_editor()
        content = text_widget.get('1.0', 'end')
        for tag in ('hl_heading', 'hl_bold', 'hl_code', 'hl_link', 'hl_image', 'hl_fence', 'hl_task', 'hl_list', 'hl_strike'):
            text_widget.tag_remove(tag, '1.0', 'end')

        fence_ranges = []
        fence_starts = []
        for m in RE_HL_FENCE.finditer(content):
            if len(fence_starts) % 2 == 0:
                fence_starts.append(m.start())
            else:
                fence_starts.append(m.end())
                s, e = fence_starts[-2], fence_starts[-1]
                fence_ranges.append((s, e))

        def in_fence(pos):
            for s, e in fence_ranges:
                if s <= pos <= e:
                    return True
            return False

        for s, e in fence_ranges:
            text_widget.tag_add('hl_fence', f'1.0+{s}c', f'1.0+{e}c')

        for m in RE_HL_HEADING.finditer(content):
            if not in_fence(m.start()):
                text_widget.tag_add('hl_heading', f'1.0+{m.start()}c', f'1.0+{m.end()}c')

        for m in RE_HL_BOLD.finditer(content):
            if not in_fence(m.start()):
                text_widget.tag_add('hl_bold', f'1.0+{m.start()}c', f'1.0+{m.end()}c')

        for m in RE_HL_CODE_INLINE.finditer(content):
            if not in_fence(m.start()):
                text_widget.tag_add('hl_code', f'1.0+{m.start()}c', f'1.0+{m.end()}c')

        for m in RE_HL_IMAGE.finditer(content):
            if not in_fence(m.start()):
                text_widget.tag_add('hl_image', f'1.0+{m.start()}c', f'1.0+{m.end()}c')

        for m in RE_HL_LINK.finditer(content):
            if not in_fence(m.start()):
                text_widget.tag_add('hl_link', f'1.0+{m.start()}c', f'1.0+{m.end()}c')

        for m in RE_HL_TASK.finditer(content):
            if not in_fence(m.start()):
                text_widget.tag_add('hl_task', f'1.0+{m.start()}c', f'1.0+{m.end()}c')

        for m in RE_HL_LIST.finditer(content):
            if not in_fence(m.start()):
                text_widget.tag_add('hl_list', f'1.0+{m.start()}c', f'1.0+{m.start() + 2}c')

        for m in RE_HL_STRIKE.finditer(content):
            if not in_fence(m.start()):
                text_widget.tag_add('hl_strike', f'1.0+{m.start()}c', f'1.0+{m.end()}c')

    # ── 实时预览 + 同步滚动 ──

    def _on_text_modified(self, event=None):
        editor = self._get_current_editor()
        if not editor.edit_modified():
            return
        editor.edit_modified(False)
        if not self._loading_editor_content:
            self._set_tab_dirty(dirty=True)
        self._sync_current_path_from_tab()
        self._update_line_numbers()
        if self._preview_timer:
            self.root.after_cancel(self._preview_timer)
        self._preview_timer = self.root.after(300, self._update_preview_and_highlight)

    def _update_preview_and_highlight(self):
        self._apply_syntax_highlight()
        self._update_preview()
        self._update_stats()
        self._update_outline()

    def _update_stats(self):
        content = self._get_current_text().strip()
        char_count = len(content)
        line_count = content.count('\n') + 1 if content else 0
        word_count = len(content.split()) if content else 0
        cjk_count = len(re.findall(r'[一-鿿㐀-䶿]', content))
        word_count = word_count + cjk_count
        self.stats_var.set(f'字符: {char_count} | 字数: {word_count} | 行数: {line_count}')

    def _update_outline(self):
        tree = self.outline_tree
        tree.delete(*tree.get_children())
        content = self._get_current_text()
        if not content.strip():
            return

        parent_stack = [('', 0)]
        for m in RE_HL_HEADING.finditer(content):
            level = len(m.group(1))
            title = m.group(2).strip()
            line_num = content[:m.start()].count('\n') + 1

            while parent_stack and parent_stack[-1][1] >= level:
                parent_stack.pop()

            parent_id = parent_stack[-1][0] if parent_stack else ''
            node_id = tree.insert(parent_id, 'end', text=title, tags=(str(line_num),))
            parent_stack.append((node_id, level))

    def _on_outline_click(self, event=None):
        sel = self.outline_tree.selection()
        if not sel:
            return
        tags = self.outline_tree.item(sel[0], 'tags')
        if not tags:
            return
        line_num = tags[0]
        text_widget = self._get_current_editor()
        if text_widget:
            text_widget.see(f'{line_num}.0')
            text_widget.mark_set('insert', f'{line_num}.0')
            text_widget.focus_set()


    def _on_editor_scroll(self):
        if not self._sync_scroll_enabled:
            return
        try:
            editor = self._get_current_editor()
            fraction = editor.yview()[0]
            self.preview.yview_moveto(fraction)
        except Exception:
            logger.debug('同步滚动失败', exc_info=True)

    def _update_preview(self):
        content = self._get_current_text().strip()
        self.preview.configure(state='normal')
        self.preview.delete('1.0', 'end')

        if not content:
            self.preview.configure(state='disabled')
            return

        try:
            blocks = parse_markdown(content)
        except Exception:
            self.preview.configure(state='disabled')
            return

        for block in blocks:
            t = block['type']
            if t == 'title':
                self.preview.insert('end', block['text'] + '\n', 'title')
            elif t == 'section_header':
                self.preview.insert('end', block['text'] + '\n', 'section')
            elif t == 'sub_header':
                self.preview.insert('end', block['text'] + '\n', 'sub_header')
            elif t == 'label_content':
                self.preview.insert('end', f'{block["label"]}：', 'label')
                self.preview.insert('end', f'{block["content"]}\n', 'body')
            elif t == 'body':
                self.preview.insert('end', block['text'] + '\n', 'body')
            elif t == 'empty':
                self.preview.insert('end', '\n', 'empty')
            elif t == 'code_block':
                lang = block.get('language', '')
                if lang:
                    self.preview.insert('end', f'[{lang}]\n', 'code_block')
                self.preview.insert('end', block['code'] + '\n', 'code_block')
            elif t == 'image':
                alt = block.get('alt', '')
                path = block.get('path', '')
                self.preview.insert('end', f'[图片: {alt or path}]\n', 'image_placeholder')
            elif t == 'task_item':
                marker = '☑ ' if block['checked'] else '☐ '
                tag = 'task_done' if block['checked'] else 'task_undone'
                self.preview.insert('end', f'{marker}{block["text"]}\n', tag)
            elif t == 'list_item':
                level = block.get('level', 1)
                indent = '  ' * level
                if block.get('ordered'):
                    prefix = f'{block.get("number", "1")}. '
                else:
                    prefix = '• '
                self.preview.insert('end', f'{indent}{prefix}{block["text"]}\n', 'list_item')
            elif t == 'table':
                headers = block['headers']
                rows = block['rows']
                col_widths = [len(h) for h in headers]
                for row in rows:
                    for ci, cell in enumerate(row):
                        if ci < len(col_widths):
                            col_widths[ci] = max(col_widths[ci], len(cell))

                def fmt_row(cells):
                    parts = []
                    for ci, cell in enumerate(cells):
                        w = col_widths[ci] if ci < len(col_widths) else len(cell)
                        parts.append(cell.ljust(w))
                    return '│ ' + ' │ '.join(parts) + ' │'

                sep = '├─' + '─┼─'.join('─' * w for w in col_widths) + '─┤'
                top = '┌─' + '─┬─'.join('─' * w for w in col_widths) + '─┐'
                bot = '└─' + '─┴─'.join('─' * w for w in col_widths) + '─┘'

                self.preview.insert('end', top + '\n', 'table')
                self.preview.insert('end', fmt_row(headers) + '\n', 'table')
                self.preview.insert('end', sep + '\n', 'table')
                for row in rows:
                    padded = row + [''] * max(0, len(headers) - len(row))
                    self.preview.insert('end', fmt_row(padded[:len(headers)]) + '\n', 'table')
                self.preview.insert('end', bot + '\n', 'table')

        self.preview.configure(state='disabled')

    # ── 拖拽支持（tkinterdnd2） ──

    def _setup_dnd(self):
        if not DND_AVAILABLE:
            return
        try:
            drop_register = getattr(self.root, 'drop_target_register', None)
            dnd_bind = getattr(self.root, 'dnd_bind', None)
            if not callable(drop_register) or not callable(dnd_bind):
                return
            drop_register(DND_FILES)
            dnd_bind('<<Drop>>', self._on_dnd_drop)
        except Exception:
            logger.warning('拖拽功能初始化失败', exc_info=True)

    def _parse_dnd_data(self, data):
        """解析 tkinterdnd2 的 event.data，处理带空格路径的 {} 包裹"""
        paths = []
        i = 0
        while i < len(data):
            if data[i] == '{':
                end = data.index('}', i)
                paths.append(data[i + 1:end])
                i = end + 2
            elif data[i] == ' ':
                i += 1
            else:
                end = data.find(' ', i)
                if end == -1:
                    end = len(data)
                paths.append(data[i:end])
                i = end + 1
        return [p for p in paths if p]

    def _on_dnd_drop(self, event):
        try:
            paths = self._parse_dnd_data(event.data)
            if not paths:
                return
            self._dispatch_drop(paths)
        except Exception:
            logger.debug('拖拽事件处理失败', exc_info=True)

    def _dispatch_drop(self, paths):
        """根据当前活动标签页决定拖拽行为"""
        try:
            current_tab = self.notebook.index(self.notebook.select())
            if current_tab == 0:
                self._handle_single_drop(paths)
            elif current_tab == 1:
                self._handle_batch_drop(paths)
        except Exception:
            logger.debug('拖拽分发失败', exc_info=True)

    def _handle_single_drop(self, paths):
        try:
            path = paths[0]
            ext = os.path.splitext(path)[1].lower()
            if ext == '.docx':
                try:
                    md_text = docx_to_markdown(path)
                    editor = self._get_current_editor()
                    editor.delete('1.0', 'end')
                    editor.insert('1.0', md_text)
                    self.status.set(f'已从 DOCX 导入: {os.path.basename(path)}')
                except Exception as e:
                    self.status.set(f'DOCX 导入失败: {e}')
            elif ext in ('.md', '.txt', '.markdown'):
                self._load_file_to_editor(path)
            else:
                self.status.set(f'不支持的文件类型: {ext}')
        except Exception as e:
            self.status.set(f'拖拽处理失败: {e}')

    def _handle_batch_drop(self, paths):
        try:
            for p in paths:
                ext = os.path.splitext(p)[1].lower()
                if ext in ('.md', '.txt', '.markdown', '.docx'):
                    existing = [self.batch_tree.item(iid, 'values')[0] for iid in self.batch_tree.get_children()]
                    if p not in existing:
                        self.batch_tree.insert('', 'end', values=(p, '等待'))
            count = len(self.batch_tree.get_children())
            self.batch_status.set(f'已添加 {count} 个文件')
        except Exception as e:
            self.batch_status.set(f'拖拽添加失败: {e}')

    # ── 取消操作 ──

    def _cancel_operation(self):
        self._cancel_event.set()
        self.status.set('正在取消...')
        self.batch_status.set('正在取消...')

    # ── 单个转换事件 ──

    def _paste(self):
        try:
            clip = self.root.clipboard_get()
            editor = self._get_current_editor()
            editor.delete('1.0', 'end')
            editor.insert('1.0', clip)
            self.status.set(f'已粘贴 {len(clip)} 个字符')
        except tk.TclError:
            self.status.set('剪贴板为空或不包含文本')

    def _open_file(self):
        path = filedialog.askopenfilename(
            filetypes=[('Markdown', '*.md'), ('文本文件', '*.txt'), ('所有文件', '*.*')],
        )
        if not path:
            return
        self._load_file_to_editor(path)

    def _import_docx(self):
        path = filedialog.askopenfilename(
            filetypes=[('Word 文档', '*.docx'), ('所有文件', '*.*')],
        )
        if not path:
            return
        try:
            md_text = docx_to_markdown(path)
            editor = self._get_current_editor()
            editor.delete('1.0', 'end')
            editor.insert('1.0', md_text)
            self.status.set(f'已从 DOCX 导入: {os.path.basename(path)}（{len(md_text)} 字符）')
        except Exception as e:
            self.status.set(f'DOCX 导入失败: {e}')

    def _clear(self):
        editor = self._get_current_editor()
        editor.delete('1.0', 'end')
        self.pbar_var.set(0)
        self.status.set('已清空')
        self.btn_open.configure(state='disabled')
        self.btn_dir.configure(state='disabled')
        self._set_current_tab_path(None)
        self._set_tab_dirty(self._get_current_tab(), False)

    def _on_template_selected(self, event=None):
        name = self.tpl_var.get()
        if not name:
            return
        content = load_template_content(name)
        if content is None:
            self.status.set(f'模板加载失败: {name}')
            return
        editor = self._get_current_editor()
        editor.delete('1.0', 'end')
        editor.insert('1.0', content)
        self._set_current_tab_path(None)
        self._set_tab_dirty(self._get_current_tab(), True)
        self.status.set(f'已加载模板: {name}')

    def _browse(self):
        init_dir = self._cfg.get('output_dir', '') or self._cfg.get('last_save_dir', '') or SCRIPT_DIR
        path = filedialog.askdirectory(initialdir=init_dir, title='选择 DOCX 输出目录')
        if path:
            self.out_dir_var.set(path)
            self._cfg['output_dir'] = path
            self._cfg['last_save_dir'] = path
            save_config(self._cfg)

    def _get_doc_features(self):
        return self._cfg.get('doc_features', {})

    def _on_format_preset_changed(self):
        preset = self.format_preset_var.get()
        is_custom = (preset == '自定义')

        if not is_custom and preset in FORMAT_PRESETS:
            vals = FORMAT_PRESETS[preset]
            self.fmt_title_font.set(vals['title_font'])
            self.fmt_title_size.set(str(vals['title_size']))
            self.fmt_heading_font.set(vals['heading_font'])
            self.fmt_heading_size.set(str(vals['heading_size']))
            self.fmt_body_font.set(vals['body_font'])
            self.fmt_body_size.set(str(vals['body_size']))
            self.fmt_line_spacing_type.set(vals['line_spacing_type'])
            self.fmt_line_spacing_value.set(str(vals['line_spacing_value']))
            self.fmt_first_indent.set(str(vals['first_line_indent_char']))
            self.fmt_margin_top.set(str(vals['margin_top_cm']))
            self.fmt_margin_bottom.set(str(vals['margin_bottom_cm']))
            self.fmt_margin_left.set(str(vals['margin_left_cm']))
            self.fmt_margin_right.set(str(vals['margin_right_cm']))

        for w in self._fmt_custom_widgets:
            if isinstance(w, ttk.Combobox):
                w.configure(state='readonly' if is_custom else 'disabled')
            else:
                w.configure(state='normal' if is_custom else 'disabled')

    def _get_format_settings(self):
        preset = self.format_preset_var.get()
        settings: dict[str, Any] = {'preset': preset}
        tpl = self.fmt_docx_template.get().strip()
        if tpl:
            settings['docx_template'] = tpl
        if preset == '自定义':
            def safe_float(val, fallback):
                try:
                    return float(val)
                except (ValueError, TypeError):
                    return fallback

            def safe_int(val, fallback):
                try:
                    return int(float(val))
                except (ValueError, TypeError):
                    return fallback

            settings['custom_overrides'] = {
                'title_font': self.fmt_title_font.get(),
                'title_size': safe_int(self.fmt_title_size.get(), 14),
                'heading_font': self.fmt_heading_font.get(),
                'heading_size': safe_int(self.fmt_heading_size.get(), 11),
                'body_font': self.fmt_body_font.get(),
                'body_size': safe_int(self.fmt_body_size.get(), 11),
                'line_spacing_type': self.fmt_line_spacing_type.get(),
                'line_spacing_value': safe_float(self.fmt_line_spacing_value.get(), 1.15),
                'first_line_indent_char': safe_int(self.fmt_first_indent.get(), 0),
                'margin_top_cm': safe_float(self.fmt_margin_top.get(), 2.54),
                'margin_bottom_cm': safe_float(self.fmt_margin_bottom.get(), 2.54),
                'margin_left_cm': safe_float(self.fmt_margin_left.get(), 3.18),
                'margin_right_cm': safe_float(self.fmt_margin_right.get(), 3.18),
            }
        return settings

    def _start(self):
        content = self._get_current_text().strip()
        if not content:
            self.status.set('错误：请先输入或粘贴内容')
            return

        output_dir = self.out_dir_var.get().strip() or SCRIPT_DIR
        if not os.path.isdir(output_dir):
            self.status.set(f'错误：输出目录不存在：{output_dir}')
            return
        out_path = None

        self.btn_convert.configure(state='disabled')
        self.btn_open.configure(state='disabled')
        self.btn_dir.configure(state='disabled')
        self.pbar_var.set(0)
        self._cancel_event.clear()
        self.btn_cancel.pack(side='left', padx=5)

        base_dir = os.path.dirname(self._current_save_path) if self._current_save_path else None
        issues = validate_conversion_input(content, base_dir=base_dir, output_path=out_path, output_dir=output_dir)
        if issues:
            if has_errors(issues):
                self.status.set('转换前检查未通过')
                messagebox.showerror('转换前检查', format_issues(issues))
                self.btn_convert.configure(state='normal')
                self.btn_cancel.pack_forget()
                return
            if not messagebox.askyesno('转换前检查', format_issues(issues) + '\n\n是否继续转换？'):
                self.status.set('已取消转换')
                self.btn_convert.configure(state='normal')
                self.btn_cancel.pack_forget()
                return

        doc_features = self._get_doc_features()
        format_settings = self._get_format_settings()

        t = threading.Thread(target=self._do_convert, args=(content, out_path, doc_features, format_settings, output_dir), daemon=True)
        t.start()
        self._poll_convert()

    def _do_convert(self, content, out_path, doc_features, format_settings, output_dir):
        try:
            feat = doc_features if any(doc_features.values()) else None
            request = ConversionRequest(
                text=content,
                output_path=out_path,
                output_dir=output_dir,
                input_path=self._current_save_path,
                doc_features=feat,
                format_settings=format_settings,
                base_dir=os.path.dirname(self._current_save_path) if self._current_save_path else None,
            )
            result = convert_request(
                request,
                progress_cb=lambda p, m: self.msg_queue.put(('p', p, m)),
                cancel_event=self._cancel_event,
            )
            self.msg_queue.put(('done', result.docx_path))
        except ConversionCancelled as e:
            self.msg_queue.put(('cancelled', str(e)))
        except InterruptedError as e:
            self.msg_queue.put(('cancelled', str(e)))
        except Exception as e:
            self.msg_queue.put(('err', str(e)))

    def _poll_convert(self):
        def on_progress(msg):
            self.pbar_var.set(msg[1])
            self.status.set(msg[2])

        def on_done(msg):
            self.output_path = msg[1]
            self.pbar_var.set(100)
            self.status.set(f'转换完成！已保存至: {msg[1]}')
            self.btn_convert.configure(state='normal')
            self.btn_open.configure(state='normal')
            self.btn_dir.configure(state='normal')
            self.btn_pdf.configure(state='normal')
            self.btn_cancel.pack_forget()
            input_desc = self._current_save_path or '编辑器内容'
            self._add_to_history(input_desc, msg[1])
            if self.also_pdf.get():
                self.root.after(100, self._export_pdf)
            return 'stop'

        def on_err(msg):
            self.pbar_var.set(0)
            self.status.set(f'错误: {msg[1]}')
            self.btn_convert.configure(state='normal')
            self.btn_cancel.pack_forget()
            return 'stop'

        def on_cancelled(msg):
            self.pbar_var.set(0)
            self.status.set('转换已取消')
            self.btn_convert.configure(state='normal')
            self.btn_cancel.pack_forget()
            return 'stop'

        self._poll_queue({'p': on_progress, 'done': on_done, 'err': on_err, 'cancelled': on_cancelled})

    def _open_result(self):
        if self.output_path and os.path.isfile(self.output_path):
            _open_path(self.output_path)

    def _open_dir(self):
        if self.output_path:
            _open_path(os.path.dirname(self.output_path))

    def _export_pdf(self):
        if not self.output_path or not os.path.isfile(self.output_path):
            self.status.set('错误：请先完成 DOCX 转换')
            return
        self.btn_pdf.configure(state='disabled')
        self.status.set('正在导出 PDF...')

        t = threading.Thread(target=self._do_export_pdf, daemon=True)
        t.start()
        self._poll_pdf()

    def _do_export_pdf(self):
        try:
            from .pdf_export import export_pdf
            pdf_path = export_pdf(self.output_path)
            self.msg_queue.put(('pdf_done', pdf_path))
        except Exception as e:
            self.msg_queue.put(('pdf_err', str(e)))

    def _poll_pdf(self):
        def on_done(msg):
            self.status.set(f'PDF 导出完成: {msg[1]}')
            self.btn_pdf.configure(state='normal')
            return 'stop'

        def on_err(msg):
            self.status.set(f'PDF 导出失败: {msg[1]}')
            self.btn_pdf.configure(state='normal')
            return 'stop'

        self._poll_queue({'pdf_done': on_done, 'pdf_err': on_err}, interval=100)

    def _export_html(self):
        content = self._get_current_text().strip()
        if not content:
            self.status.set('错误：请先输入或粘贴内容')
            return
        try:
            from .html_export import export_html
            init_dir = self._cfg.get('last_save_dir', '') or SCRIPT_DIR
            path = filedialog.asksaveasfilename(
                defaultextension='.html',
                filetypes=[('HTML 文件', '*.html'), ('所有文件', '*.*')],
                initialdir=init_dir,
            )
            if not path:
                return
            export_html(content, path)
            self.status.set(f'HTML 导出完成: {path}')
            self._add_to_history(self._current_save_path or '编辑器内容', path)
        except ImportError:
            self.status.set('错误：缺少 markdown 库，请运行 pip install markdown')
        except Exception as e:
            self.status.set(f'HTML 导出失败: {e}')

    # ── 批量转换事件 ──

    def _batch_add(self):
        paths = filedialog.askopenfilenames(
            filetypes=[
                ('Markdown', '*.md'),
                ('文本文件', '*.txt'),
                ('Word 文档', '*.docx'),
                ('所有文件', '*.*'),
            ],
        )
        for p in paths:
            existing = [self.batch_tree.item(iid, 'values')[0] for iid in self.batch_tree.get_children()]
            if p not in existing:
                self.batch_tree.insert('', 'end', values=(p, '等待'))
        count = len(self.batch_tree.get_children())
        self.batch_status.set(f'已添加 {count} 个文件')

    def _batch_remove(self):
        for iid in self.batch_tree.selection():
            self.batch_tree.delete(iid)
        count = len(self.batch_tree.get_children())
        self.batch_status.set(f'剩余 {count} 个文件')

    def _batch_clear(self):
        for iid in self.batch_tree.get_children():
            self.batch_tree.delete(iid)
        self.batch_pbar_var.set(0)
        self.batch_status.set('列表已清空')
        self.btn_batch_open_dir.configure(state='disabled')

    def _batch_browse_dir(self):
        init_dir = self._cfg.get('batch_output_dir', '') or SCRIPT_DIR
        path = filedialog.askdirectory(initialdir=init_dir)
        if path:
            self.batch_out_var.set(path)
            self._cfg['batch_output_dir'] = path
            save_config(self._cfg)

    def _batch_start(self):
        items = self.batch_tree.get_children()
        if not items:
            self.batch_status.set('错误：请先添加文件')
            return

        self.btn_batch_convert.configure(state='disabled')
        self.batch_pbar_var.set(0)
        self._cancel_event.clear()
        self.btn_batch_cancel.pack(side='left', padx=5)
        output_dir = self.batch_out_var.get()
        format_settings = self._get_format_settings()
        also_pdf = self.batch_also_pdf.get()

        t = threading.Thread(target=self._do_batch, args=(items, output_dir, format_settings, also_pdf), daemon=True)
        t.start()
        self._poll_batch()

    def _do_batch(self, items, output_dir, format_settings, also_pdf):
        total = len(items)
        success = 0
        fail = 0

        for idx, iid in enumerate(items):
            if self._cancel_event.is_set():
                self.msg_queue.put(('batch_done', success, fail, True))
                return

            file_path = self.batch_tree.item(iid, 'values')[0]
            self.msg_queue.put(('batch_status', iid, '转换中...'))

            try:
                basename = os.path.splitext(os.path.basename(file_path))[0]
                out_path = os.path.join(output_dir, f'{basename}.docx')
                convert_file(
                    file_path,
                    output_path=out_path,
                    output_dir=output_dir,
                    format_settings=format_settings,
                    also_pdf=also_pdf,
                    cancel_event=self._cancel_event,
                )

                self.msg_queue.put(('batch_status', iid, '完成'))
                success += 1
            except Exception as e:
                self.msg_queue.put(('batch_status', iid, f'失败: {e}'))
                fail += 1

            pct = int(100 * (idx + 1) / total)
            self.msg_queue.put(('batch_progress', pct, f'进度 {idx + 1}/{total}'))

        self.msg_queue.put(('batch_done', success, fail, False))

    def _poll_batch(self):
        def on_status(msg):
            self.batch_tree.item(msg[1], values=(self.batch_tree.item(msg[1], 'values')[0], msg[2]))

        def on_progress(msg):
            self.batch_pbar_var.set(msg[1])
            self.batch_status.set(msg[2])

        def on_done(msg):
            self.batch_pbar_var.set(100)
            cancelled = msg[3] if len(msg) > 3 else False
            suffix = '（已取消）' if cancelled else ''
            self.batch_status.set(f'批量转换完成{suffix}！成功 {msg[1]} 个，失败 {msg[2]} 个')
            self.btn_batch_convert.configure(state='normal')
            self.btn_batch_open_dir.configure(state='normal')
            self.btn_batch_cancel.pack_forget()
            return 'stop'

        self._poll_queue({
            'batch_status': on_status,
            'batch_progress': on_progress,
            'batch_done': on_done,
        })

    def _batch_open_dir(self):
        output_dir = self.batch_out_var.get()
        if os.path.isdir(output_dir):
            _open_path(output_dir)

    # ── 设置标签页 ──

    def _build_settings_tab(self, notebook):
        tab = ttk.Frame(notebook, padding=APP_PADDING)
        notebook.add(tab, text='设置')

        canvas = tk.Canvas(tab, highlightthickness=0)
        self.settings_canvas = canvas
        self._style_canvas(canvas)
        scrollbar = ttk.Scrollbar(tab, orient='vertical', command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)

        scroll_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scroll_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        cfg = self._cfg

        # ── 文档功能设置 ──
        doc_frame = ttk.LabelFrame(scroll_frame, text='文档功能', padding=10)
        doc_frame.pack(fill='x', pady=(0, 10), padx=5)

        feat = cfg.get('doc_features', {})

        self.feat_toc = tk.BooleanVar(value=feat.get('toc_enabled', False))
        ttk.Checkbutton(doc_frame, text='生成目录（TOC）— 打开文档后按 F9 更新', variable=self.feat_toc).pack(anchor='w', pady=2)

        self.feat_page_num = tk.BooleanVar(value=feat.get('page_number', True))
        ttk.Checkbutton(doc_frame, text='显示页码', variable=self.feat_page_num).pack(anchor='w', pady=2)

        self.feat_header = tk.BooleanVar(value=feat.get('header_enabled', False))
        ttk.Checkbutton(doc_frame, text='显示页眉', variable=self.feat_header).pack(anchor='w', pady=2)

        hdr_frame = ttk.Frame(doc_frame)
        hdr_frame.pack(fill='x', pady=2, padx=(20, 0))
        ttk.Label(hdr_frame, text='页眉文字:').pack(side='left')
        self.feat_header_text = tk.StringVar(value=feat.get('header_text', ''))
        ttk.Entry(hdr_frame, textvariable=self.feat_header_text, width=40).pack(side='left', padx=5)

        logo_frame = ttk.Frame(doc_frame)
        logo_frame.pack(fill='x', pady=2, padx=(20, 0))
        ttk.Label(logo_frame, text='Logo 图片:').pack(side='left')
        self.feat_logo_path = tk.StringVar(value=feat.get('logo_path', ''))
        ttk.Entry(logo_frame, textvariable=self.feat_logo_path, width=30).pack(side='left', padx=5)
        ttk.Button(logo_frame, text='浏览', command=self._browse_logo, width=6).pack(side='left')

        self.feat_watermark = tk.BooleanVar(value=feat.get('watermark_enabled', False))
        ttk.Checkbutton(doc_frame, text='添加水印', variable=self.feat_watermark).pack(anchor='w', pady=2)

        wm_frame = ttk.Frame(doc_frame)
        wm_frame.pack(fill='x', pady=2, padx=(20, 0))
        ttk.Label(wm_frame, text='水印文字:').pack(side='left')
        self.feat_watermark_text = tk.StringVar(value=feat.get('watermark_text', ''))
        ttk.Entry(wm_frame, textvariable=self.feat_watermark_text, width=30).pack(side='left', padx=5)

        # ── 排版格式设置 ──
        fmt_frame = ttk.LabelFrame(scroll_frame, text='排版格式', padding=10)
        fmt_frame.pack(fill='x', pady=(0, 10), padx=5)

        preset_row = ttk.Frame(fmt_frame)
        preset_row.pack(fill='x', pady=(0, 8))
        ttk.Label(preset_row, text='格式预设:').pack(side='left')
        fmt_preset_combo = ttk.Combobox(
            preset_row, textvariable=self.format_preset_var,
            values=['默认', '公文', '自定义'],
            state='readonly', width=10,
        )
        fmt_preset_combo.pack(side='left', padx=5)
        ttk.Label(
            preset_row, text='选择「自定义」可修改下方参数',
            style='Muted.TLabel',
        ).pack(side='left', padx=10)

        tpl_row = ttk.Frame(fmt_frame)
        tpl_row.pack(fill='x', pady=(0, 8))
        ttk.Label(tpl_row, text='DOCX 模板:').pack(side='left')
        self.fmt_docx_template = tk.StringVar(
            value=self._cfg.get('format_settings', {}).get('docx_template', ''),
        )
        ttk.Entry(tpl_row, textvariable=self.fmt_docx_template, width=40).pack(side='left', padx=5)
        ttk.Button(tpl_row, text='浏览', command=self._browse_docx_template).pack(side='left', padx=(0, 5))
        ttk.Button(tpl_row, text='清除', command=lambda: self.fmt_docx_template.set('')).pack(side='left')
        ttk.Label(
            tpl_row, text='可选：使用企业 .docx 模板的样式',
            style='Muted.TLabel',
        ).pack(side='left', padx=10)

        overrides = self._cfg.get('format_settings', {}).get('custom_overrides', {})
        default_fmt = FORMAT_PRESETS.get(self.format_preset_var.get(), FORMAT_PRESETS['默认'])

        self.fmt_custom_frame = ttk.Frame(fmt_frame)
        self.fmt_custom_frame.pack(fill='x')

        # 字体设置
        font_lf = ttk.LabelFrame(self.fmt_custom_frame, text='字体', padding=8)
        font_lf.pack(fill='x', pady=(0, 5))
        font_grid = ttk.Frame(font_lf)
        font_grid.pack(fill='x')

        common_fonts = ['等线', '宋体', '黑体', '仿宋', '楷体', '微软雅黑', '方正小标宋体',
                        'Calibri', 'Times New Roman', 'Arial']

        ttk.Label(font_grid, text='标题字体:').grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.fmt_title_font = tk.StringVar(value=overrides.get('title_font', default_fmt['title_font']))
        ttk.Combobox(font_grid, textvariable=self.fmt_title_font, values=common_fonts, width=18).grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(font_grid, text='字号:').grid(row=0, column=2, sticky='w', padx=5, pady=2)
        self.fmt_title_size = tk.StringVar(value=str(overrides.get('title_size', default_fmt['title_size'])))
        ttk.Entry(font_grid, textvariable=self.fmt_title_size, width=6).grid(row=0, column=3, padx=5, pady=2)
        ttk.Label(font_grid, text='磅').grid(row=0, column=4, sticky='w')

        ttk.Label(font_grid, text='章节字体:').grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.fmt_heading_font = tk.StringVar(value=overrides.get('heading_font', default_fmt['heading_font']))
        ttk.Combobox(font_grid, textvariable=self.fmt_heading_font, values=common_fonts, width=18).grid(row=1, column=1, padx=5, pady=2)
        ttk.Label(font_grid, text='字号:').grid(row=1, column=2, sticky='w', padx=5, pady=2)
        self.fmt_heading_size = tk.StringVar(value=str(overrides.get('heading_size', default_fmt['heading_size'])))
        ttk.Entry(font_grid, textvariable=self.fmt_heading_size, width=6).grid(row=1, column=3, padx=5, pady=2)
        ttk.Label(font_grid, text='磅').grid(row=1, column=4, sticky='w')

        ttk.Label(font_grid, text='正文字体:').grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.fmt_body_font = tk.StringVar(value=overrides.get('body_font', default_fmt['body_font']))
        ttk.Combobox(font_grid, textvariable=self.fmt_body_font, values=common_fonts, width=18).grid(row=2, column=1, padx=5, pady=2)
        ttk.Label(font_grid, text='字号:').grid(row=2, column=2, sticky='w', padx=5, pady=2)
        self.fmt_body_size = tk.StringVar(value=str(overrides.get('body_size', default_fmt['body_size'])))
        ttk.Entry(font_grid, textvariable=self.fmt_body_size, width=6).grid(row=2, column=3, padx=5, pady=2)
        ttk.Label(font_grid, text='磅').grid(row=2, column=4, sticky='w')

        # 行距设置
        spacing_lf = ttk.LabelFrame(self.fmt_custom_frame, text='段落', padding=8)
        spacing_lf.pack(fill='x', pady=(0, 5))
        sp_grid = ttk.Frame(spacing_lf)
        sp_grid.pack(fill='x')

        ttk.Label(sp_grid, text='行距类型:').grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.fmt_line_spacing_type = tk.StringVar(
            value=overrides.get('line_spacing_type', default_fmt['line_spacing_type']),
        )
        ttk.Combobox(
            sp_grid, textvariable=self.fmt_line_spacing_type,
            values=['multiple', 'fixed'], state='readonly', width=10,
        ).grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(sp_grid, text='（multiple=倍数, fixed=固定磅值）',
                  style='Muted.TLabel').grid(row=0, column=2, columnspan=3, sticky='w', padx=5)

        ttk.Label(sp_grid, text='行距值:').grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.fmt_line_spacing_value = tk.StringVar(
            value=str(overrides.get('line_spacing_value', default_fmt['line_spacing_value'])),
        )
        ttk.Entry(sp_grid, textvariable=self.fmt_line_spacing_value, width=8).grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(sp_grid, text='首行缩进:').grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.fmt_first_indent = tk.StringVar(
            value=str(overrides.get('first_line_indent_char', default_fmt['first_line_indent_char'])),
        )
        ttk.Entry(sp_grid, textvariable=self.fmt_first_indent, width=8).grid(row=2, column=1, padx=5, pady=2)
        ttk.Label(sp_grid, text='字符').grid(row=2, column=2, sticky='w', padx=5)

        # 页边距设置
        margin_lf = ttk.LabelFrame(self.fmt_custom_frame, text='页边距 (cm)', padding=8)
        margin_lf.pack(fill='x', pady=(0, 5))
        mg_grid = ttk.Frame(margin_lf)
        mg_grid.pack(fill='x')

        self.fmt_margin_top = tk.StringVar(value=str(overrides.get('margin_top_cm', default_fmt['margin_top_cm'])))
        self.fmt_margin_bottom = tk.StringVar(value=str(overrides.get('margin_bottom_cm', default_fmt['margin_bottom_cm'])))
        self.fmt_margin_left = tk.StringVar(value=str(overrides.get('margin_left_cm', default_fmt['margin_left_cm'])))
        self.fmt_margin_right = tk.StringVar(value=str(overrides.get('margin_right_cm', default_fmt['margin_right_cm'])))

        ttk.Label(mg_grid, text='上:').grid(row=0, column=0, sticky='w', padx=5, pady=2)
        ttk.Entry(mg_grid, textvariable=self.fmt_margin_top, width=6).grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(mg_grid, text='下:').grid(row=0, column=2, sticky='w', padx=5, pady=2)
        ttk.Entry(mg_grid, textvariable=self.fmt_margin_bottom, width=6).grid(row=0, column=3, padx=5, pady=2)
        ttk.Label(mg_grid, text='左:').grid(row=0, column=4, sticky='w', padx=5, pady=2)
        ttk.Entry(mg_grid, textvariable=self.fmt_margin_left, width=6).grid(row=0, column=5, padx=5, pady=2)
        ttk.Label(mg_grid, text='右:').grid(row=0, column=6, sticky='w', padx=5, pady=2)
        ttk.Entry(mg_grid, textvariable=self.fmt_margin_right, width=6).grid(row=0, column=7, padx=5, pady=2)

        # 收集所有自定义控件用于启用/禁用
        self._fmt_custom_widgets = []
        for frame in (font_grid, sp_grid, mg_grid):
            for child in frame.winfo_children():
                if isinstance(child, (ttk.Entry, ttk.Combobox)):
                    self._fmt_custom_widgets.append(child)

        self.format_preset_var.trace_add('write', lambda *_: self._on_format_preset_changed())
        self._on_format_preset_changed()

        # ── 保存按钮 ──
        ttk.Button(scroll_frame, text='保存设置', command=self._save_settings, style='Accent.TButton').pack(anchor='w', padx=5, pady=10)
        self.settings_status = tk.StringVar(value='')
        ttk.Label(scroll_frame, textvariable=self.settings_status).pack(anchor='w', padx=5)

    def _browse_logo(self):
        path = filedialog.askopenfilename(
            filetypes=[('图片文件', '*.png *.jpg *.jpeg *.bmp'), ('所有文件', '*.*')],
        )
        if path:
            self.feat_logo_path.set(path)

    def _browse_docx_template(self):
        path = filedialog.askopenfilename(
            filetypes=[('Word 模板', '*.docx'), ('所有文件', '*.*')],
        )
        if path:
            self.fmt_docx_template.set(path)

    def _save_settings(self):
        self._cfg['doc_features'] = {
            'toc_enabled': self.feat_toc.get(),
            'header_enabled': self.feat_header.get(),
            'header_text': self.feat_header_text.get().strip(),
            'footer_enabled': True,
            'page_number': self.feat_page_num.get(),
            'logo_path': self.feat_logo_path.get().strip(),
            'watermark_enabled': self.feat_watermark.get(),
            'watermark_text': self.feat_watermark_text.get().strip(),
        }
        self._cfg['format_settings'] = self._get_format_settings()
        save_config(self._cfg)
        self.settings_status.set('设置已保存')

    # ── AI 润色 ──

    def _ai_polish(self):
        content = self._get_current_text().strip()
        if not content:
            self.status.set('错误：请先输入或粘贴内容')
            return

        model = self._cfg.get('ai_model') or (AI_MODELS[0] if AI_MODELS else 'gpt-5.5')
        api_key = self._cfg.get('api_key', '').strip()
        template = self.ai_tpl_var.get()

        if not api_key and not os.getenv('MDTONG_AI_API_KEY') and not os.getenv('ANTHROPIC_API_KEY'):
            self.status.set('错误：请先在 config.json 或环境变量中配置 AI API Key')
            return

        detected = detect_doc_type(content)
        if detected != template:
            if messagebox.askyesno(
                '模板推荐',
                f'根据内容分析，推荐使用「{detected}」模板。\n'
                f'当前选择：{template}\n\n是否切换到推荐模板？',
            ):
                template = detected
                self.ai_tpl_var.set(detected)

        char_count = len(content)
        if not messagebox.askyesno(
            'AI 润色确认',
            f'即将使用 AI 润色当前内容：\n\n'
            f'  模板：{template}\n'
            f'  模型：{model}\n'
            f'  内容长度：{char_count} 字符\n\n'
            f'润色后原始内容将保留在撤销历史中（Ctrl+Z 可恢复）。\n'
            f'确定开始润色吗？',
        ):
            return

        self._ai_backup_content = content

        self.btn_ai.configure(state='disabled')
        self.btn_convert.configure(state='disabled')
        self._cancel_event.clear()
        self.btn_cancel.pack(side='left', padx=5)
        self.status.set(f'AI 润色中（{template}），请稍候...')
        editor = self._get_current_editor()
        editor.delete('1.0', 'end')

        t = threading.Thread(target=self._do_ai_polish, args=(content, model, template, api_key), daemon=True)
        t.start()
        self._poll_ai()

    def _restore_ai_backup(self):
        if hasattr(self, '_ai_backup_content') and self._ai_backup_content:
            editor = self._get_current_editor()
            editor.delete('1.0', 'end')
            editor.insert('1.0', self._ai_backup_content)
            self.status.set('已恢复润色前的原始内容')
            self._ai_backup_content = None

    def _do_ai_polish(self, content, model, template, api_key):
        try:
            from .ai_polish import polish

            def on_chunk(chunk):
                if self._cancel_event.is_set():
                    raise InterruptedError('用户取消')
                self.msg_queue.put(('ai_chunk', chunk))

            _, usage = polish(content, api_key=api_key, model=model, template=template, on_chunk=on_chunk)
            self.msg_queue.put(('ai_done', usage))
        except InterruptedError:
            self.msg_queue.put(('ai_cancelled',))
        except Exception as e:
            self.msg_queue.put(('ai_err', str(e)))

    def _poll_ai(self):
        def on_chunk(msg):
            editor = self._get_current_editor()
            editor.insert('end', msg[1])
            editor.see('end')

        def on_done(msg):
            usage = msg[1] if len(msg) > 1 else {}
            usage_text = ''
            if usage:
                inp = usage.get('input_tokens', 0)
                out = usage.get('output_tokens', 0)
                cache_read = usage.get('cache_read', 0)
                cache_create = usage.get('cache_creation', 0)
                parts = [f'输入 {inp}', f'输出 {out}']
                if cache_read:
                    parts.append(f'缓存命中 {cache_read}')
                if cache_create:
                    parts.append(f'缓存写入 {cache_create}')
                usage_text = f'（Token: {" | ".join(parts)}）'
            self.status.set(f'AI 润色完成！{usage_text}')
            self.btn_ai.configure(state='normal')
            self.btn_convert.configure(state='normal')
            self.btn_cancel.pack_forget()
            polished = self._get_current_text()
            original = getattr(self, '_ai_backup_content', None)
            self._ai_backup_content = None
            if original:
                self._show_diff_btn(original, polished)
            return 'stop'

        def on_cancelled(msg):
            self._restore_ai_backup()
            self.status.set('AI 润色已取消，已恢复原始内容')
            self.btn_ai.configure(state='normal')
            self.btn_convert.configure(state='normal')
            self.btn_cancel.pack_forget()
            return 'stop'

        def on_err(msg):
            self._restore_ai_backup()
            self.status.set(f'AI 润色失败: {msg[1]}（已恢复原始内容）')
            self.btn_ai.configure(state='normal')
            self.btn_convert.configure(state='normal')
            self.btn_cancel.pack_forget()
            return 'stop'

        self._poll_queue({
            'ai_chunk': on_chunk,
            'ai_done': on_done,
            'ai_cancelled': on_cancelled,
            'ai_err': on_err,
        })

    def _show_diff_btn(self, original, polished):
        """在状态栏附近显示「对比原文」按钮"""
        if hasattr(self, '_diff_btn') and self._diff_btn:
            self._diff_btn.destroy()
        btn_frame = self.root.nametowidget(self.btn_ai.winfo_parent())
        self._diff_btn = ttk.Button(
            btn_frame, text='对比原文',
            command=lambda: self._show_diff_window(original, polished),
        )
        self._diff_btn.pack(side='left', padx=5)

    def _show_diff_window(self, original, polished):
        """弹出并排对比窗口"""
        win = tk.Toplevel(self.root)
        win.title('AI 润色对比')
        win.geometry('1200x700')
        win.transient(self.root)
        self._style_popup(win)

        header = ttk.Frame(win)
        header.pack(fill='x', padx=10, pady=5)
        ttk.Label(header, text='原始内容', font=(UI_FONT, 11, 'bold')).pack(side='left', expand=True)
        ttk.Label(header, text='润色后内容', font=(UI_FONT, 11, 'bold')).pack(side='right', expand=True)

        pane = ttk.PanedWindow(win, orient='horizontal')
        pane.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        p = self._theme_palette()
        left = tk.Text(
            pane, wrap='word', font=(MONO_FONT, 10), padx=10, pady=10,
            relief='flat', borderwidth=0, bg=p['bg'], fg=p['fg'],
            insertbackground=p['insert_bg'], selectbackground=p['select_bg'],
        )
        left.insert('1.0', original)
        left.configure(state='disabled')
        pane.add(left, weight=1)

        right = tk.Text(
            pane, wrap='word', font=(MONO_FONT, 10), padx=10, pady=10,
            relief='flat', borderwidth=0, bg=p['bg'], fg=p['fg'],
            insertbackground=p['insert_bg'], selectbackground=p['select_bg'],
        )
        right.insert('1.0', polished)
        right.configure(state='disabled')
        pane.add(right, weight=1)

        bottom = ttk.Frame(win)
        bottom.pack(fill='x', padx=10, pady=(0, 10))
        ttk.Button(
            bottom, text='恢复原文',
            command=lambda: self._restore_from_diff(original, win),
            style='Accent.TButton',
        ).pack(side='left', padx=5)
        ttk.Button(bottom, text='关闭', command=win.destroy).pack(side='right', padx=5)

    def _restore_from_diff(self, original, win):
        """从对比视图恢复原文"""
        editor = self._get_current_editor()
        editor.delete('1.0', 'end')
        editor.insert('1.0', original)
        self.status.set('已恢复原始内容')
        win.destroy()
        if hasattr(self, '_diff_btn') and self._diff_btn:
            self._diff_btn.destroy()
            self._diff_btn = None

    # ── 自动保存草稿 ──

    def _get_draft_path(self):
        return os.path.join(SCRIPT_DIR, DRAFT_FILENAME)

    def _start_auto_save(self):
        self._auto_save_draft()

    def _auto_save_draft(self):
        try:
            content = self._get_current_text().strip()
            if content:
                with open(self._get_draft_path(), 'w', encoding='utf-8') as f:
                    f.write(content)
        except Exception:
            logger.debug('自动保存草稿失败', exc_info=True)
        self.root.after(60000, self._auto_save_draft)

    def _check_draft_recovery(self):
        draft_path = self._get_draft_path()
        if not os.path.isfile(draft_path):
            return
        try:
            with open(draft_path, 'r', encoding='utf-8') as f:
                content = f.read().strip()
            if content:
                if messagebox.askyesno('恢复草稿', '检测到上次未保存的编辑内容，是否恢复？'):
                    editor = self._get_current_editor()
                    editor.delete('1.0', 'end')
                    editor.insert('1.0', content)
                    self.status.set('已恢复草稿内容')
                else:
                    os.remove(draft_path)
        except Exception:
            logger.warning('草稿恢复失败', exc_info=True)

    def _delete_draft(self):
        try:
            draft_path = self._get_draft_path()
            if os.path.isfile(draft_path):
                os.remove(draft_path)
        except Exception:
            logger.debug('删除草稿文件失败', exc_info=True)

    # ── 关于对话框 ──

    def _show_about(self):
        about = tk.Toplevel(self.root)
        about.title('关于 MD通')
        about.geometry('400x350')
        self._style_popup(about)
        about.resizable(False, False)
        about.transient(self.root)
        about.grab_set()

        ttk.Label(about, text='MD通', font=(UI_FONT, 18, 'bold')).pack(pady=(20, 5))
        ttk.Label(about, text=f'版本 {APP_VERSION}', font=(UI_FONT, 10)).pack()
        ttk.Label(about, text='Markdown → DOCX 转换工具', font=(UI_FONT, 10)).pack(pady=(5, 15))

        info = ttk.LabelFrame(about, text='功能特性', padding=10)
        info.pack(fill='x', padx=20, pady=5)
        features = [
            'Markdown 解析与 DOCX 生成',
            'AI 智能润色（Claude API）',
            'DOCX 反向转换',
            '实时预览与语法高亮',
            '批量转换 + PDF 导出',
            '多种格式预设（默认/公文/自定义）',
            '拖拽文件支持',
            'HTML 导出',
        ]
        for f in features:
            ttk.Label(info, text=f'  • {f}', font=(UI_FONT, 9)).pack(anchor='w')

        shortcuts = ttk.LabelFrame(about, text='快捷键', padding=10)
        shortcuts.pack(fill='x', padx=20, pady=5)
        keys = 'Ctrl+S 保存 | Ctrl+O 打开 | Ctrl+N 清空 | F5 转换\nCtrl+F 查找 | Ctrl+H 替换 | Ctrl+T 新标签 | Ctrl+Shift+P 润色'
        ttk.Label(shortcuts, text=keys, font=(UI_FONT, 8)).pack()

        ttk.Button(about, text='确定', command=about.destroy, style='Accent.TButton').pack(pady=10)

    # ── Markdown 语法帮助 ──

    def _show_cheatsheet(self):
        cs = tk.Toplevel(self.root)
        cs.title('Markdown 语法速查')
        cs.geometry('500x600')
        cs.transient(self.root)
        self._style_popup(cs)

        p = self._theme_palette()
        text = tk.Text(
            cs, wrap='word', font=(MONO_FONT, 10), padx=12, pady=12,
            relief='flat', borderwidth=0, bg=p['bg'], fg=p['fg'],
            insertbackground=p['insert_bg'], selectbackground=p['select_bg'],
        )
        sb = ttk.Scrollbar(cs, command=text.yview)
        text.configure(yscrollcommand=sb.set)
        text.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        text.insert('1.0', MARKDOWN_CHEATSHEET)
        text.configure(state='disabled')

    # ── 窗口关闭 ──

    def _on_close(self):
        dirty_tabs = [tab for tab in getattr(self, 'editor_tabs', []) if tab.get('dirty')]
        if dirty_tabs:
            if not messagebox.askyesno('未保存的更改', f'还有 {len(dirty_tabs)} 个文档未保存，确定退出吗？'):
                return
        self._cfg['window_geometry'] = self.root.geometry()
        save_config(self._cfg)
        content = self._get_current_text().strip()
        if not content:
            self._delete_draft()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


def run_gui():
    App().run()
