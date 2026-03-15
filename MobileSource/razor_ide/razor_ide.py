#!/usr/bin/env python3
"""
Razor IDE — A lightweight Python IDE for editing Blazor Razor (.razor) files.
Built with tkinter. Supports syntax highlighting, file tree, tabbed editing,
line numbers, find/replace, and dark/light themes.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont

# Allow running from the razor_ide directory
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from syntax_highlighter import configure_tags, highlight, THEME_DARK, THEME_LIGHT


# ─── Defaults ────────────────────────────────────────────────────────────────

DEFAULT_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    '..', 'HelloWorldBlazor'
)
BG_DARK = '#1E1E1E'
FG_DARK = '#D4D4D4'
BG_LIGHT = '#FFFFFF'
FG_LIGHT = '#000000'
HIGHLIGHT_DELAY_MS = 300


# ─── Line Number Widget ─────────────────────────────────────────────────────

class LineNumbers(tk.Canvas):
    def __init__(self, master, text_widget, **kwargs):
        super().__init__(master, **kwargs)
        self.text_widget = text_widget

    def redraw(self):
        self.delete('all')
        i = self.text_widget.index('@0,0')
        while True:
            dline = self.text_widget.dlineinfo(i)
            if dline is None:
                break
            linenum = str(i).split('.')[0]
            self.create_text(
                self.winfo_width() - 8, dline[1],
                anchor='ne', text=linenum,
                fill='#858585', font=('Consolas', 11)
            )
            i = self.text_widget.index(f'{i}+1line')


# ─── Tab (one open file) ────────────────────────────────────────────────────

class EditorTab:
    """Holds the state for one open file tab."""
    def __init__(self, frame, text_widget, line_numbers, filepath=None):
        self.frame = frame
        self.text = text_widget
        self.line_numbers = line_numbers
        self.filepath = filepath
        self.modified = False


# ─── Main Application ───────────────────────────────────────────────────────

class RazorIDE(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Razor IDE')
        self.geometry('1200x750')
        self.theme = 'dark'
        self.tabs: dict[str, EditorTab] = {}   # notebook tab id -> EditorTab
        self._highlight_job = None

        self._build_menu()
        self._build_toolbar()
        self._build_panes()
        self._build_statusbar()
        self._apply_theme()

        # Load default project tree
        project_dir = os.path.normpath(DEFAULT_DIR)
        if os.path.isdir(project_dir):
            self._populate_tree(project_dir)

        self.protocol('WM_DELETE_WINDOW', self._on_close)

    # ── Menu ─────────────────────────────────────────────────────────────

    def _build_menu(self):
        menubar = tk.Menu(self, tearoff=0)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label='Open File…', accelerator='Ctrl+O', command=self._open_file)
        file_menu.add_command(label='Open Folder…', command=self._open_folder)
        file_menu.add_separator()
        file_menu.add_command(label='Save', accelerator='Ctrl+S', command=self._save_file)
        file_menu.add_command(label='Save As…', command=self._save_as)
        file_menu.add_separator()
        file_menu.add_command(label='Close Tab', accelerator='Ctrl+W', command=self._close_tab)
        file_menu.add_separator()
        file_menu.add_command(label='Exit', command=self._on_close)
        menubar.add_cascade(label='File', menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label='Undo', accelerator='Ctrl+Z', command=lambda: self._active_text_cmd('edit_undo'))
        edit_menu.add_command(label='Redo', accelerator='Ctrl+Y', command=lambda: self._active_text_cmd('edit_redo'))
        edit_menu.add_separator()
        edit_menu.add_command(label='Find / Replace…', accelerator='Ctrl+F', command=self._show_find)
        menubar.add_cascade(label='Edit', menu=edit_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label='Toggle Theme', command=self._toggle_theme)
        view_menu.add_command(label='Increase Font', accelerator='Ctrl++', command=lambda: self._change_font_size(1))
        view_menu.add_command(label='Decrease Font', accelerator='Ctrl+-', command=lambda: self._change_font_size(-1))
        menubar.add_cascade(label='View', menu=view_menu)

        self.config(menu=menubar)

        # Key bindings
        self.bind_all('<Control-o>', lambda e: self._open_file())
        self.bind_all('<Control-s>', lambda e: self._save_file())
        self.bind_all('<Control-w>', lambda e: self._close_tab())
        self.bind_all('<Control-f>', lambda e: self._show_find())
        self.bind_all('<Control-plus>', lambda e: self._change_font_size(1))
        self.bind_all('<Control-minus>', lambda e: self._change_font_size(-1))
        self.bind_all('<Control-equal>', lambda e: self._change_font_size(1))

    # ── Toolbar ──────────────────────────────────────────────────────────

    def _build_toolbar(self):
        self.toolbar = tk.Frame(self, height=32)
        self.toolbar.pack(side='top', fill='x')

        buttons = [
            ('Open', self._open_file),
            ('Save', self._save_file),
            ('Theme', self._toggle_theme),
            ('Find', self._show_find),
        ]
        for label, cmd in buttons:
            b = tk.Button(self.toolbar, text=label, command=cmd,
                          relief='flat', padx=10, pady=2)
            b.pack(side='left', padx=2, pady=2)

    # ── Main Layout ──────────────────────────────────────────────────────

    def _build_panes(self):
        self.paned = tk.PanedWindow(self, orient='horizontal', sashwidth=4)
        self.paned.pack(fill='both', expand=True)

        # File tree
        tree_frame = tk.Frame(self.paned, width=250)
        self.tree = ttk.Treeview(tree_frame, show='tree')
        tree_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side='left', fill='both', expand=True)
        tree_scroll.pack(side='right', fill='y')
        self.tree.bind('<<TreeviewOpen>>', self._on_tree_expand)
        self.tree.bind('<Double-1>', self._on_tree_double_click)
        self.paned.add(tree_frame)

        # Notebook (tabs)
        self.notebook = ttk.Notebook(self.paned)
        self.notebook.pack(fill='both', expand=True)
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)
        self.paned.add(self.notebook)

    # ── Status Bar ───────────────────────────────────────────────────────

    def _build_statusbar(self):
        self.statusbar = tk.Label(self, text='Ready', anchor='w', padx=8)
        self.statusbar.pack(side='bottom', fill='x')

    def _set_status(self, msg):
        self.statusbar.config(text=msg)

    # ── Theme ────────────────────────────────────────────────────────────

    def _apply_theme(self):
        bg = BG_DARK if self.theme == 'dark' else BG_LIGHT
        fg = FG_DARK if self.theme == 'dark' else FG_LIGHT
        toolbar_bg = '#2D2D2D' if self.theme == 'dark' else '#E8E8E8'

        self.configure(bg=bg)
        self.toolbar.configure(bg=toolbar_bg)
        for child in self.toolbar.winfo_children():
            child.configure(bg=toolbar_bg, fg=fg)
        self.statusbar.configure(bg=toolbar_bg, fg=fg)

        # Update all open editor tabs
        for tab in self.tabs.values():
            tab.text.configure(bg=bg, fg=fg, insertbackground=fg,
                               selectbackground='#264F78' if self.theme == 'dark' else '#ADD6FF')
            tab.line_numbers.configure(bg=bg)
            configure_tags(tab.text, self.theme)
            highlight(tab.text)
            tab.line_numbers.redraw()

    def _toggle_theme(self):
        self.theme = 'light' if self.theme == 'dark' else 'dark'
        self._apply_theme()

    # ── Font Size ────────────────────────────────────────────────────────

    def _change_font_size(self, delta):
        for tab in self.tabs.values():
            f = tkfont.Font(font=tab.text['font'])
            new_size = max(8, f.actual()['size'] + delta)
            tab.text.configure(font=('Consolas', new_size))

    # ── File Tree ────────────────────────────────────────────────────────

    def _populate_tree(self, path):
        self.tree.delete(*self.tree.get_children())
        self._tree_root = path
        root_node = self.tree.insert('', 'end', text=os.path.basename(path),
                                     values=[path], open=True)
        self._fill_tree(root_node, path)

    def _fill_tree(self, parent, path):
        try:
            entries = sorted(os.listdir(path))
        except PermissionError:
            return

        # Directories first, then files
        dirs = [e for e in entries if os.path.isdir(os.path.join(path, e)) and not e.startswith('.') and e not in ('bin', 'obj')]
        files = [e for e in entries if os.path.isfile(os.path.join(path, e)) and not e.startswith('.')]

        for d in dirs:
            full = os.path.join(path, d)
            node = self.tree.insert(parent, 'end', text=d, values=[full])
            # Add dummy child so the expand arrow appears
            self.tree.insert(node, 'end', text='')

        for f in files:
            full = os.path.join(path, f)
            self.tree.insert(parent, 'end', text=f, values=[full])

    def _on_tree_expand(self, event):
        node = self.tree.focus()
        children = self.tree.get_children(node)
        # If the only child is the dummy, replace it
        if len(children) == 1 and self.tree.item(children[0], 'text') == '':
            self.tree.delete(children[0])
            path = self.tree.item(node, 'values')[0]
            self._fill_tree(node, path)

    def _on_tree_double_click(self, event):
        node = self.tree.focus()
        values = self.tree.item(node, 'values')
        if not values:
            return
        path = values[0]
        if os.path.isfile(path):
            self._open_file_in_tab(path)

    # ── Tab Management ───────────────────────────────────────────────────

    def _create_editor_tab(self, filepath=None):
        frame = tk.Frame(self.notebook)

        # Line numbers
        line_canvas = LineNumbers(frame, None, width=50)
        line_canvas.pack(side='left', fill='y')

        # Text editor
        text = tk.Text(
            frame, wrap='none', undo=True,
            font=('Consolas', 13), tabs='    ',
            borderwidth=0, padx=8, pady=4,
        )
        # Scrollbars
        vscroll = ttk.Scrollbar(frame, orient='vertical', command=text.yview)
        hscroll = ttk.Scrollbar(frame, orient='horizontal', command=text.xview)
        text.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        vscroll.pack(side='right', fill='y')
        hscroll.pack(side='bottom', fill='x')
        text.pack(side='left', fill='both', expand=True)

        line_canvas.text_widget = text

        # Configure highlighting
        bg = BG_DARK if self.theme == 'dark' else BG_LIGHT
        fg = FG_DARK if self.theme == 'dark' else FG_LIGHT
        text.configure(bg=bg, fg=fg, insertbackground=fg,
                       selectbackground='#264F78' if self.theme == 'dark' else '#ADD6FF')
        line_canvas.configure(bg=bg)
        configure_tags(text, self.theme)

        # Events
        text.bind('<<Modified>>', lambda e, t=text: self._on_text_modified(t))
        text.bind('<KeyRelease>', lambda e: self._schedule_highlight())
        text.bind('<Configure>', lambda e: self._update_line_numbers())
        text.bind('<MouseWheel>', lambda e: self.after_idle(self._update_line_numbers))

        tab_title = os.path.basename(filepath) if filepath else 'Untitled'
        self.notebook.add(frame, text=tab_title)
        self.notebook.select(frame)

        tab_id = str(frame)
        et = EditorTab(frame, text, line_canvas, filepath)
        self.tabs[tab_id] = et
        return et

    def _get_active_tab(self) -> EditorTab | None:
        sel = self.notebook.select()
        return self.tabs.get(str(sel))

    def _on_tab_changed(self, event):
        tab = self._get_active_tab()
        if tab and tab.filepath:
            self._set_status(tab.filepath)
        self._update_line_numbers()

    def _on_text_modified(self, text_widget):
        if text_widget.edit_modified():
            text_widget.edit_modified(False)
            self._update_line_numbers()

    def _schedule_highlight(self):
        if self._highlight_job:
            self.after_cancel(self._highlight_job)
        self._highlight_job = self.after(HIGHLIGHT_DELAY_MS, self._do_highlight)

    def _do_highlight(self):
        tab = self._get_active_tab()
        if tab:
            highlight(tab.text)

    def _update_line_numbers(self):
        tab = self._get_active_tab()
        if tab:
            tab.line_numbers.redraw()

    def _active_text_cmd(self, method):
        tab = self._get_active_tab()
        if tab:
            try:
                getattr(tab.text, method)()
            except tk.TclError:
                pass

    # ── File Operations ──────────────────────────────────────────────────

    def _open_file(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ('Razor Files', '*.razor'),
                ('C# Files', '*.cs'),
                ('XAML Files', '*.xaml'),
                ('All Files', '*.*'),
            ]
        )
        if path:
            self._open_file_in_tab(path)

    def _open_file_in_tab(self, path):
        # Check if already open
        for tab_id, tab in self.tabs.items():
            if tab.filepath and os.path.normpath(tab.filepath) == os.path.normpath(path):
                self.notebook.select(tab.frame)
                return

        et = self._create_editor_tab(filepath=path)
        try:
            with open(path, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            et.text.delete('1.0', 'end')
            et.text.insert('1.0', content)
            et.text.edit_modified(False)
            et.text.edit_reset()
            highlight(et.text)
            self._update_line_numbers()
            self._set_status(path)
        except Exception as e:
            messagebox.showerror('Error', f'Could not open file:\n{e}')

    def _open_folder(self):
        path = filedialog.askdirectory()
        if path:
            self._populate_tree(path)

    def _save_file(self):
        tab = self._get_active_tab()
        if not tab:
            return
        if tab.filepath:
            self._write_file(tab, tab.filepath)
        else:
            self._save_as()

    def _save_as(self):
        tab = self._get_active_tab()
        if not tab:
            return
        path = filedialog.asksaveasfilename(
            defaultextension='.razor',
            filetypes=[
                ('Razor Files', '*.razor'),
                ('C# Files', '*.cs'),
                ('All Files', '*.*'),
            ]
        )
        if path:
            tab.filepath = path
            self.notebook.tab(tab.frame, text=os.path.basename(path))
            self._write_file(tab, path)

    def _write_file(self, tab, path):
        try:
            content = tab.text.get('1.0', 'end-1c')
            with open(path, 'w', encoding='utf-8') as f:
                f.write(content)
            tab.text.edit_modified(False)
            self._set_status(f'Saved: {path}')
        except Exception as e:
            messagebox.showerror('Error', f'Could not save file:\n{e}')

    def _close_tab(self):
        tab = self._get_active_tab()
        if not tab:
            return
        sel = self.notebook.select()
        self.notebook.forget(sel)
        del self.tabs[str(sel)]

    # ── Find / Replace ───────────────────────────────────────────────────

    def _show_find(self):
        tab = self._get_active_tab()
        if not tab:
            return

        win = tk.Toplevel(self)
        win.title('Find / Replace')
        win.geometry('420x150')
        win.transient(self)

        tk.Label(win, text='Find:').grid(row=0, column=0, padx=6, pady=4, sticky='e')
        find_entry = tk.Entry(win, width=35)
        find_entry.grid(row=0, column=1, padx=6, pady=4)
        find_entry.focus_set()

        tk.Label(win, text='Replace:').grid(row=1, column=0, padx=6, pady=4, sticky='e')
        repl_entry = tk.Entry(win, width=35)
        repl_entry.grid(row=1, column=1, padx=6, pady=4)

        def do_find():
            tab.text.tag_remove('found', '1.0', 'end')
            query = find_entry.get()
            if not query:
                return
            idx = '1.0'
            count = 0
            while True:
                idx = tab.text.search(query, idx, nocase=True, stopindex='end')
                if not idx:
                    break
                end_idx = f'{idx}+{len(query)}c'
                tab.text.tag_add('found', idx, end_idx)
                idx = end_idx
                count += 1
            tab.text.tag_configure('found', background='#515C6A' if self.theme == 'dark' else '#FFD700')
            self._set_status(f'Found {count} occurrence(s)')

        def do_replace_all():
            query = find_entry.get()
            replacement = repl_entry.get()
            if not query:
                return
            content = tab.text.get('1.0', 'end-1c')
            new_content = content.replace(query, replacement)
            tab.text.delete('1.0', 'end')
            tab.text.insert('1.0', new_content)
            highlight(tab.text)
            self._set_status('Replaced all occurrences')

        btn_frame = tk.Frame(win)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=8)
        tk.Button(btn_frame, text='Find All', command=do_find).pack(side='left', padx=6)
        tk.Button(btn_frame, text='Replace All', command=do_replace_all).pack(side='left', padx=6)
        tk.Button(btn_frame, text='Close', command=win.destroy).pack(side='left', padx=6)

        win.bind('<Return>', lambda e: do_find())
        win.bind('<Escape>', lambda e: win.destroy())

    # ── Close ────────────────────────────────────────────────────────────

    def _on_close(self):
        self.destroy()


# ─── Entry Point ─────────────────────────────────────────────────────────────

def main():
    app = RazorIDE()
    # Auto-open Home.razor if it exists
    home_razor = os.path.normpath(os.path.join(DEFAULT_DIR, 'Components', 'Pages', 'Home.razor'))
    if os.path.isfile(home_razor):
        app.after(100, lambda: app._open_file_in_tab(home_razor))
    app.mainloop()


if __name__ == '__main__':
    main()
