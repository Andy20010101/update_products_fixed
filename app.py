"""
Products 数据更新工具 - GUI Application
Tkinter wizard-style interface.
"""

import os
import sys
import traceback
from datetime import datetime
from tkinter import (
    Tk, Toplevel, Frame, Label, Button, Entry, Listbox, StringVar,
    Menu, Canvas, filedialog, messagebox, ttk, Text, Scrollbar,
    BOTH, LEFT, RIGHT, TOP, BOTTOM, X, Y, VERTICAL, HORIZONTAL,
    END, DISABLED, NORMAL,
)

if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, APP_DIR)

from engine import (
    SOURCE_DIR, TARGET_DIR, TARGET_FILENAME, TARGET_SHEET,
    SNAPSHOT_DIR_NAME,
    find_latest_products_file, read_products_source,
    analyze_changes, execute_update,
    create_snapshot, list_snapshots, restore_from_snapshot,
    cleanup_snapshots,
    archive_source, UpdateStats, ChangeDetail,
)

# ═══════════════════════════════════════
# Styles & Constants
# ═══════════════════════════════════════

FONT_NORMAL = ('Microsoft YaHei', 10) if os.name == 'nt' else ('TkDefaultFont', 11)
FONT_BOLD = ('Microsoft YaHei', 10, 'bold') if os.name == 'nt' else ('TkDefaultFont', 11, 'bold')
FONT_TITLE = ('Microsoft YaHei', 14, 'bold') if os.name == 'nt' else ('TkDefaultFont', 14, 'bold')
FONT_SMALL = ('Microsoft YaHei', 9) if os.name == 'nt' else ('TkDefaultFont', 9)
FONT_MONO = ('Consolas', 10) if os.name == 'nt' else ('Menlo', 11)

COLOR_BG = '#f0f2f5'
COLOR_HEADER = '#2c3e50'
COLOR_PRIMARY = '#3498db'
COLOR_SUCCESS = '#27ae60'
COLOR_WARNING = '#f39c12'
COLOR_DANGER = '#e74c3c'
COLOR_WHITE = '#ffffff'
COLOR_STEP_ACTIVE = '#3498db'
COLOR_STEP_DONE = '#27ae60'
COLOR_STEP_PENDING = '#bdc3c7'
COLOR_TEXT_DARK = '#2c3e50'
COLOR_TEXT_LIGHT = '#ecf0f1'


# ═══════════════════════════════════════
# Wizard Step Indicator
# ═══════════════════════════════════════

class StepIndicator(Frame):
    STEPS = ['选择文件', '预览变更', '执行更新']

    def __init__(self, parent):
        super().__init__(parent, bg=COLOR_WHITE)
        self.labels = []
        self.separators = []

        for i, name in enumerate(self.STEPS):
            lbl = Label(self, text=f'{i + 1}\n{name}', font=FONT_SMALL,
                        bg=COLOR_WHITE, fg=COLOR_STEP_PENDING,
                        justify='center', width=12)
            lbl.pack(side=LEFT, padx=(10, 0))
            self.labels.append(lbl)

            if i < len(self.STEPS) - 1:
                sep = Label(self, text='─────', font=FONT_SMALL,
                            bg=COLOR_WHITE, fg=COLOR_STEP_PENDING)
                sep.pack(side=LEFT)
                self.separators.append(sep)

    def set_active(self, step: int):
        for i, lbl in enumerate(self.labels):
            if i < step:
                lbl.config(fg=COLOR_STEP_DONE)
            elif i == step:
                lbl.config(fg=COLOR_STEP_ACTIVE, font=FONT_BOLD)
            else:
                lbl.config(fg=COLOR_STEP_PENDING, font=FONT_SMALL)

        for i, sep in enumerate(self.separators):
            sep.config(fg=COLOR_STEP_DONE if i < step else COLOR_STEP_PENDING)


# ═══════════════════════════════════════
# Step 0: File Selection
# ═══════════════════════════════════════

class FileFrame(Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=COLOR_BG)
        self.app = app

        Label(self, text='选择文件', font=FONT_TITLE, bg=COLOR_BG,
              fg=COLOR_HEADER).pack(anchor='w', pady=(10, 20))

        # ── Source file ──
        Label(self, text='源文件 (Products 导出表)', font=FONT_BOLD,
              bg=COLOR_BG).pack(anchor='w', pady=(10, 5))

        f1 = Frame(self, bg=COLOR_WHITE, relief='groove', bd=1)
        f1.pack(fill=X, pady=(0, 5))

        self.source_path_var = StringVar()
        self.source_label = Label(f1, textvariable=self.source_path_var,
                                  font=FONT_SMALL, bg=COLOR_WHITE, fg='#666',
                                  anchor='w', padx=8, pady=6)
        self.source_label.pack(side=LEFT, fill=X, expand=True)

        Button(f1, text='选择文件', command=self._pick_source,
               font=FONT_SMALL, bg=COLOR_PRIMARY, fg=COLOR_WHITE,
               relief='flat', padx=12, pady=4).pack(side=RIGHT, padx=5, pady=4)

        Button(f1, text='自动查找', command=self._auto_source,
               font=FONT_SMALL, relief='flat', padx=12, pady=4).pack(side=RIGHT, pady=4)

        # ── Target file ──
        Label(self, text='目标文件 (分析优化表)', font=FONT_BOLD,
              bg=COLOR_BG).pack(anchor='w', pady=(15, 5))

        f2 = Frame(self, bg=COLOR_WHITE, relief='groove', bd=1)
        f2.pack(fill=X, pady=(0, 5))

        self.target_path_var = StringVar()
        self.target_label = Label(f2, textvariable=self.target_path_var,
                                  font=FONT_SMALL, bg=COLOR_WHITE, fg='#666',
                                  anchor='w', padx=8, pady=6)
        self.target_label.pack(side=LEFT, fill=X, expand=True)

        Button(f2, text='选择文件', command=self._pick_target,
               font=FONT_SMALL, bg=COLOR_PRIMARY, fg=COLOR_WHITE,
               relief='flat', padx=12, pady=4).pack(side=RIGHT, padx=5, pady=4)

        # ── File info ──
        Label(self, text='文件信息', font=FONT_BOLD,
              bg=COLOR_BG).pack(anchor='w', pady=(15, 5))

        info_frame = Frame(self, bg=COLOR_WHITE, relief='groove', bd=1)
        info_frame.pack(fill=BOTH, expand=True, pady=(0, 10))

        self.info_text = Text(info_frame, height=6, font=FONT_SMALL,
                              bg=COLOR_WHITE, fg=COLOR_TEXT_DARK,
                              relief='flat', state=DISABLED, wrap='word')
        self.info_text.pack(fill=BOTH, expand=True, padx=8, pady=8)

        # ── Snapshot quick view ──
        Label(self, text='快照管理', font=FONT_BOLD,
              bg=COLOR_BG).pack(anchor='w', pady=(10, 5))

        snap_frame = Frame(self, bg=COLOR_WHITE, relief='groove', bd=1)
        snap_frame.pack(fill=X, pady=(0, 5))

        self.snap_listbox = Listbox(snap_frame, height=3, font=FONT_SMALL,
                                    relief='flat', exportselection=False)
        self.snap_listbox.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)

        snap_scroll = Scrollbar(snap_frame, orient=VERTICAL, command=self.snap_listbox.yview)
        snap_scroll.pack(side=RIGHT, fill=Y, pady=5)
        self.snap_listbox.config(yscrollcommand=snap_scroll.set)

        snap_btn_frame = Frame(snap_frame, bg=COLOR_WHITE)
        snap_btn_frame.pack(side=RIGHT, padx=5)

        Button(snap_btn_frame, text='撤销 (Undo)', command=self._undo_from_file,
               font=FONT_SMALL, bg=COLOR_DANGER, fg=COLOR_WHITE,
               relief='flat', padx=10, pady=3).pack(pady=2)

        Button(snap_btn_frame, text='清理快照', command=self._cleanup_snapshots,
               font=FONT_SMALL, bg=COLOR_WARNING, fg=COLOR_WHITE,
               relief='flat', padx=10, pady=3).pack(pady=2)

        Button(snap_btn_frame, text='刷新列表', command=self._refresh_snapshots,
               font=FONT_SMALL, relief='flat', padx=10, pady=3).pack(pady=2)

        # Auto-detect on init
        self._auto_source()
        self._auto_target()

    def _auto_source(self):
        latest = find_latest_products_file(SOURCE_DIR)
        if latest:
            self.source_path_var.set(latest)
            self.app.source_path = latest
            self._update_info()
        else:
            self.source_path_var.set('(未找到 Products 文件，请手动选择)')
            self.app.source_path = ''

    def _pick_source(self):
        path = filedialog.askopenfilename(
            title='选择 Products 导出文件',
            filetypes=[('Excel 文件', '*.xls *.xlsx'), ('所有文件', '*.*')],
            initialdir=SOURCE_DIR if os.path.isdir(SOURCE_DIR) else os.path.expanduser('~'),
        )
        if path:
            self.source_path_var.set(path)
            self.app.source_path = path
            self._update_info()

    def _auto_target(self):
        target_path = os.path.join(TARGET_DIR, TARGET_FILENAME)
        if os.path.exists(target_path):
            self.target_path_var.set(target_path)
            self.app.target_path = target_path
        else:
            self.target_path_var.set('(目标文件不存在，请手动选择)')
            self.app.target_path = ''

    def _pick_target(self):
        path = filedialog.askopenfilename(
            title='选择分析优化表',
            filetypes=[('Excel 文件', '*.xlsx'), ('所有文件', '*.*')],
            initialdir=TARGET_DIR if os.path.isdir(TARGET_DIR) else os.path.expanduser('~'),
        )
        if path:
            self.target_path_var.set(path)
            self.app.target_path = path

    def _update_info(self):
        self.info_text.config(state=NORMAL)
        self.info_text.delete('1.0', END)
        path = self.app.source_path
        if path and os.path.exists(path):
            size_kb = os.path.getsize(path) / 1024
            mtime = datetime.fromtimestamp(os.path.getmtime(path))
            self.info_text.insert(END, f'文件名: {os.path.basename(path)}\n')
            self.info_text.insert(END, f'大小: {size_kb:.1f} KB\n')
            self.info_text.insert(END, f'修改时间: {mtime:%Y-%m-%d %H:%M:%S}\n')
            self.info_text.insert(END, f'路径: {os.path.dirname(path)}\n')
        else:
            self.info_text.insert(END, '尚未选择有效的源文件\n')
        self.info_text.config(state=DISABLED)

    def _refresh_snapshots(self):
        self.snap_listbox.delete(0, END)
        if self.app.target_path:
            snapshot_dir = os.path.join(os.path.dirname(self.app.target_path), SNAPSHOT_DIR_NAME)
            snapshots = list_snapshots(snapshot_dir)
            for sp in snapshots:
                mtime = datetime.fromtimestamp(os.path.getmtime(sp))
                self.snap_listbox.insert(END, f'{os.path.basename(sp)}  ({mtime:%Y-%m-%d %H:%M:%S})')
            if not snapshots:
                self.snap_listbox.insert(END, '(暂无快照)')
        else:
            self.snap_listbox.insert(END, '(请先选择目标文件)')

    def _undo_from_file(self):
        if not self.app.target_path or not os.path.exists(self.app.target_path):
            messagebox.showwarning('提示', '请先选择有效的目标文件')
            return

        snapshot_dir = os.path.join(os.path.dirname(self.app.target_path), SNAPSHOT_DIR_NAME)
        snapshots = list_snapshots(snapshot_dir)
        if not snapshots:
            messagebox.showinfo('提示', '没有可用的快照')
            return

        latest = snapshots[0]
        mtime = datetime.fromtimestamp(os.path.getmtime(latest))
        if messagebox.askyesno('确认撤销',
                               f'将恢复到上一次执行前的状态:\n\n'
                               f'{os.path.basename(latest)}\n'
                               f'时间: {mtime:%Y-%m-%d %H:%M:%S}\n\n'
                               f'恢复前会自动备份当前文件，确定继续？'):
            try:
                create_snapshot(self.app.target_path, snapshot_dir)
                restore_from_snapshot(self.app.target_path, latest)
                messagebox.showinfo('完成', f'已从快照恢复:\n{os.path.basename(latest)}')
                self._refresh_snapshots()
            except Exception as e:
                messagebox.showerror('撤销失败', str(e))

    def _cleanup_snapshots(self):
        if not self.app.target_path:
            messagebox.showwarning('提示', '请先选择目标文件')
            return
        snapshot_dir = os.path.join(os.path.dirname(self.app.target_path), SNAPSHOT_DIR_NAME)
        snapshots = list_snapshots(snapshot_dir)
        if len(snapshots) <= 20:
            messagebox.showinfo('提示', f'当前只有 {len(snapshots)} 个快照，无需清理（保留最近 20 个）')
            return

        if messagebox.askyesno('确认清理',
                               f'共 {len(snapshots)} 个快照，将删除最早的 {len(snapshots) - 20} 个，'
                               f'保留最近 20 个。\n\n确定继续？'):
            deleted = cleanup_snapshots(snapshot_dir, keep=20)
            messagebox.showinfo('完成', f'已删除 {deleted} 个旧快照，保留最近 20 个')
            self._refresh_snapshots()

    def refresh_snapshots(self):
        self._refresh_snapshots()


# ═══════════════════════════════════════
# Step 1: Preview Changes
# ═══════════════════════════════════════

class PreviewFrame(Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=COLOR_BG)
        self.app = app

        Label(self, text='预览变更', font=FONT_TITLE, bg=COLOR_BG,
              fg=COLOR_HEADER).pack(anchor='w', pady=(10, 5))

        self.status_label = Label(self, text='点击"分析变更"查看待更新内容',
                                  font=FONT_SMALL, bg=COLOR_BG, fg='#888')
        self.status_label.pack(anchor='w', pady=(0, 10))

        # ── Summary bar ──
        summary_frame = Frame(self, bg=COLOR_WHITE, relief='groove', bd=1)
        summary_frame.pack(fill=X, pady=(0, 10))

        self.summary_vars = {
            'total': StringVar(value='总数: -'),
            'updated': StringVar(value='更新: -'),
            'appended': StringVar(value='追加: -'),
            'skipped': StringVar(value='跳过: -'),
            'deduplicated': StringVar(value='重复ID忽略: -'),
        }
        for key, var in self.summary_vars.items():
            Label(summary_frame, textvariable=var, font=FONT_BOLD,
                  bg=COLOR_WHITE, fg=COLOR_TEXT_DARK,
                  padx=12, pady=6).pack(side=LEFT)

        # ── Changes tree ──
        tree_frame = Frame(self, bg=COLOR_WHITE, relief='groove', bd=1)
        tree_frame.pack(fill=BOTH, expand=True, pady=(0, 10))

        columns = ('product_id', 'action', 'details')
        self.tree = ttk.Treeview(tree_frame, columns=columns,
                                 show='headings', height=12)
        self.tree.heading('product_id', text='产品ID')
        self.tree.heading('action', text='操作')
        self.tree.heading('details', text='变更详情')
        self.tree.column('product_id', width=150, anchor='center')
        self.tree.column('action', width=60, anchor='center')
        self.tree.column('details', width=500)

        tree_scroll = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        tree_scroll.pack(side=RIGHT, fill=Y, pady=5)

        # ── Analyze button ──
        btn_frame = Frame(self, bg=COLOR_BG)
        btn_frame.pack(fill=X)

        Button(btn_frame, text='分析变更', command=self._run_analysis,
               font=FONT_BOLD, bg=COLOR_PRIMARY, fg=COLOR_WHITE,
               relief='flat', padx=20, pady=6).pack(side=LEFT)

    def _run_analysis(self):
        if not self.app.source_path or not os.path.exists(self.app.source_path):
            messagebox.showwarning('提示', '请先在"选择文件"步骤中选择有效的源文件')
            return
        if not self.app.target_path or not os.path.exists(self.app.target_path):
            messagebox.showwarning('提示', '请先在"选择文件"步骤中选择有效的目标文件')
            return

        self.status_label.config(text='正在分析...')
        self.update_idletasks()

        try:
            source_rows = read_products_source(self.app.source_path)
            self.app.source_rows = source_rows
            stats = analyze_changes(self.app.target_path, source_rows)
            self.app.preview_stats = stats

            # Update summary
            self.summary_vars['total'].set(f'总数: {stats.total}')
            self.summary_vars['updated'].set(f'更新: {stats.updated}')
            self.summary_vars['appended'].set(f'追加: {stats.appended}')
            self.summary_vars['skipped'].set(f'跳过: {stats.skipped}')
            self.summary_vars['deduplicated'].set(f'重复ID忽略: {stats.deduplicated}')

            # Populate tree
            self.tree.delete(*self.tree.get_children())
            for change in stats.changes:
                if change.action == '更新':
                    detail_str = '; '.join(
                        f'{field}: {old} → {new}'
                        for field, old, new in change.changes[:5]
                    )
                    if len(change.changes) > 5:
                        detail_str += f' ... 等{len(change.changes)}项'
                    tag = 'updated'
                else:
                    detail_str = '新追加行'
                    tag = 'appended'

                self.tree.insert('', END, values=(change.product_id, change.action, detail_str), tags=(tag,))

            self.tree.tag_configure('updated', background='#fff3cd')
            self.tree.tag_configure('appended', background='#d4edda')

            status = f'分析完成: {stats.updated} 行更新, {stats.appended} 行追加, {stats.skipped} 行跳过'
            if stats.deduplicated:
                status += f'，重复ID忽略 {stats.deduplicated} 行'
            self.status_label.config(text=status)
            if stats.warnings:
                messagebox.showwarning('重复产品ID已处理', '\n'.join(stats.warnings))
        except Exception as e:
            self.status_label.config(text='分析失败')
            messagebox.showerror('分析错误', str(e))


# ═══════════════════════════════════════
# Step 2: Execute Update
# ═══════════════════════════════════════

class ExecuteFrame(Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=COLOR_BG)
        self.app = app
        self.export_done = False

        Label(self, text='执行更新', font=FONT_TITLE, bg=COLOR_BG,
              fg=COLOR_HEADER).pack(anchor='w', pady=(10, 5))

        Label(self, text='确认无误后点击"执行更新"，脚本将修改目标文件并归档源文件。',
              font=FONT_SMALL, bg=COLOR_BG, fg='#888').pack(anchor='w', pady=(0, 15))

        # ── Action buttons ──
        action_frame = Frame(self, bg=COLOR_WHITE, relief='groove', bd=1)
        action_frame.pack(fill=X, pady=(0, 15))

        Label(action_frame, text='操作区', font=FONT_BOLD, bg=COLOR_WHITE,
              fg=COLOR_TEXT_DARK).pack(anchor='w', padx=10, pady=(10, 5))

        btn_row = Frame(action_frame, bg=COLOR_WHITE)
        btn_row.pack(fill=X, padx=10, pady=(5, 10))

        self.execute_btn = Button(btn_row, text='执行更新', command=self._execute,
                                  font=FONT_BOLD, bg=COLOR_SUCCESS, fg=COLOR_WHITE,
                                  relief='flat', padx=20, pady=6)
        self.execute_btn.pack(side=LEFT, padx=(0, 10))

        Button(btn_row, text='创建快照', command=self._manual_snapshot,
               font=FONT_SMALL, relief='flat', padx=12, pady=6).pack(side=LEFT, padx=5)

        Button(btn_row, text='撤销上次更新', command=self._undo_last,
               font=FONT_SMALL, bg=COLOR_DANGER, fg=COLOR_WHITE,
               relief='flat', padx=12, pady=6).pack(side=LEFT, padx=5)

        Button(btn_row, text='清理快照', command=self._cleanup_snapshots,
               font=FONT_SMALL, bg=COLOR_WARNING, fg=COLOR_WHITE,
               relief='flat', padx=12, pady=6).pack(side=LEFT, padx=5)

        # ── Result output ──
        Label(self, text='执行日志', font=FONT_BOLD, bg=COLOR_BG).pack(anchor='w', pady=(5, 5))

        result_frame = Frame(self, bg=COLOR_HEADER, relief='groove', bd=1)
        result_frame.pack(fill=BOTH, expand=True)

        self.result_text = Text(result_frame, height=12, font=FONT_MONO,
                                bg=COLOR_HEADER, fg=COLOR_TEXT_LIGHT,
                                relief='flat', wrap='word', state=DISABLED)
        self.result_text.pack(fill=BOTH, expand=True, padx=10, pady=10)

        result_scroll = ttk.Scrollbar(result_frame, orient=VERTICAL,
                                      command=self.result_text.yview)
        result_scroll.pack(side=RIGHT, fill=Y, pady=10)
        self.result_text.configure(yscrollcommand=result_scroll.set)

    def _log(self, msg: str):
        self.result_text.config(state=NORMAL)
        self.result_text.insert(END, msg + '\n')
        self.result_text.see(END)
        self.result_text.config(state=DISABLED)
        self.update_idletasks()

    def _manual_snapshot(self):
        if not self.app.target_path or not os.path.exists(self.app.target_path):
            messagebox.showwarning('提示', '请先选择有效的目标文件')
            return
        snapshot_dir = os.path.join(os.path.dirname(self.app.target_path), SNAPSHOT_DIR_NAME)
        path = create_snapshot(self.app.target_path, snapshot_dir)
        self._log(f'[快照] {os.path.basename(path)}')
        messagebox.showinfo('完成', f'快照已创建:\n{os.path.basename(path)}')

    def _undo_last(self):
        if not self.app.target_path or not os.path.exists(self.app.target_path):
            messagebox.showwarning('提示', '请先选择有效的目标文件')
            return
        snapshot_dir = os.path.join(os.path.dirname(self.app.target_path), SNAPSHOT_DIR_NAME)
        snapshots = list_snapshots(snapshot_dir)
        if not snapshots:
            messagebox.showinfo('提示', '没有可用的快照')
            return

        latest = snapshots[0]
        mtime = datetime.fromtimestamp(os.path.getmtime(latest))
        if messagebox.askyesno('确认撤销',
                               f'将恢复到上一次执行前的状态:\n\n'
                               f'{os.path.basename(latest)}\n'
                               f'时间: {mtime:%Y-%m-%d %H:%M:%S}\n\n'
                               f'恢复前会自动备份当前文件，确定继续？'):
            try:
                create_snapshot(self.app.target_path, snapshot_dir)
                restore_from_snapshot(self.app.target_path, latest)
                self._log(f'[撤销] 已恢复到: {os.path.basename(latest)}')
                messagebox.showinfo('完成', f'已从快照恢复')
            except Exception as e:
                self._log(f'[错误] 撤销失败: {e}')
                messagebox.showerror('撤销失败', str(e))

    def _execute(self):
        if self.export_done:
            messagebox.showinfo('提示', '已经执行过更新，请重新启动程序以再次操作')
            return
        if not self.app.source_rows:
            messagebox.showwarning('提示', '请先在"预览变更"步骤中完成分析')
            return
        if not self.app.target_path or not os.path.exists(self.app.target_path):
            messagebox.showwarning('提示', '请先选择有效的目标文件')
            return

        stats = self.app.preview_stats
        confirm_msg = (f'确定要执行更新吗？\n\n'
                       f'此操作将:\n'
                       f'1. 更新 {stats.updated} 行已有产品\n'
                       f'2. 追加 {stats.appended} 行新产品\n'
                       f'3. 归档源文件到"已处理"目录\n\n'
                       f'（执行前已自动创建快照）')

        if not messagebox.askyesno('确认执行', confirm_msg):
            return

        self.execute_btn.config(state=DISABLED)
        self._log('=' * 50)
        self._log(f'开始执行更新 - {datetime.now():%Y-%m-%d %H:%M:%S}')
        self._log(f'源文件: {os.path.basename(self.app.source_path)}')
        self._log(f'目标文件: {os.path.basename(self.app.target_path)}')
        self._log('')

        try:
            # Snapshot before execution
            snapshot_dir = os.path.join(os.path.dirname(self.app.target_path), SNAPSHOT_DIR_NAME)
            snap = create_snapshot(self.app.target_path, snapshot_dir)
            self._log(f'[快照] 已创建执行前备份: {os.path.basename(snap)}')

            # Execute
            result = execute_update(self.app.target_path, self.app.source_rows)
            self._log(f'[更新] 完成: {result.updated} 行更新, {result.appended} 行追加, {result.skipped} 行跳过')

            # Archive source
            dest = archive_source(self.app.source_path, SOURCE_DIR)
            self._log(f'[归档] 源文件已归档: {os.path.basename(dest)}')
            self.app.source_path = ''

            self._log('')
            self._log('更新成功完成!')
            self.export_done = True

            messagebox.showinfo('完成', f'更新完成!\n\n'
                                       f'更新: {result.updated} 行\n'
                                       f'追加: {result.appended} 行\n'
                                       f'跳过: {result.skipped} 行')
        except Exception as e:
            self._log(f'[错误] 执行失败: {e}')
            self._log(traceback.format_exc())
            messagebox.showerror('执行错误', str(e))
            self.execute_btn.config(state=NORMAL)

    def _cleanup_snapshots(self):
        if not self.app.target_path or not os.path.exists(self.app.target_path):
            messagebox.showwarning('提示', '请先选择有效的目标文件')
            return
        snapshot_dir = os.path.join(os.path.dirname(self.app.target_path), SNAPSHOT_DIR_NAME)
        snapshots = list_snapshots(snapshot_dir)
        if len(snapshots) <= 20:
            messagebox.showinfo('提示', f'当前只有 {len(snapshots)} 个快照，无需清理（保留最近 20 个）')
            return
        if messagebox.askyesno('确认清理',
                               f'共 {len(snapshots)} 个快照，将删除最早的 {len(snapshots) - 20} 个，'
                               f'保留最近 20 个。\n\n确定继续？'):
            deleted = cleanup_snapshots(snapshot_dir, keep=20)
            messagebox.showinfo('完成', f'已删除 {deleted} 个旧快照，保留最近 20 个')
            self._log(f'[清理] 已删除 {deleted} 个旧快照，保留最近 20 个')

    def load_data(self):
        self.result_text.config(state=NORMAL)
        self.result_text.delete('1.0', END)
        self.result_text.config(state=DISABLED)
        self.export_done = False
        self.execute_btn.config(state=NORMAL)

        if hasattr(self.app, 'preview_stats'):
            stats = self.app.preview_stats
            self._log(f'待执行摘要:')
            self._log(f'  总数: {stats.total}')
            self._log(f'  更新: {stats.updated} 行')
            self._log(f'  追加: {stats.appended} 行')
            self._log(f'  跳过: {stats.skipped} 行')
            self._log(f'  重复ID忽略: {stats.deduplicated} 行')
            for warning in stats.warnings:
                self._log(f'  警告: {warning}')
            self._log('')
            self._log('点击"执行更新"开始操作...')


# ═══════════════════════════════════════
# Main Application
# ═══════════════════════════════════════

class ProductsUpdateApp(Tk):
    def __init__(self):
        super().__init__()

        self.title('Products 数据更新工具')
        self.geometry('900x680')
        self.minsize(760, 520)
        self.configure(bg=COLOR_BG)

        # Application state
        self.source_path: str = ''
        self.target_path: str = ''
        self.source_rows: list = []
        self.preview_stats: UpdateStats = None
        self.current_step = 0

        # Build UI
        self._build_menu()
        self._build_header()
        self._build_step_indicator()
        self._build_navigation()
        self._build_content()

        self.protocol('WM_DELETE_WINDOW', self._on_close)

    # ── Menu ──
    def _build_menu(self):
        menubar = Menu(self)

        file_menu = Menu(menubar, tearoff=0)
        file_menu.add_command(label='选择源文件...', command=self._pick_source)
        file_menu.add_command(label='选择目标文件...', command=self._pick_target)
        file_menu.add_separator()
        file_menu.add_command(label='退出', command=self._on_close)
        menubar.add_cascade(label='文件', menu=file_menu)

        snap_menu = Menu(menubar, tearoff=0)
        snap_menu.add_command(label='查看快照列表', command=self._list_snapshots)
        snap_menu.add_command(label='创建快照', command=self._manual_snapshot)
        snap_menu.add_separator()
        snap_menu.add_command(label='撤销 (Undo)', command=self._undo)
        snap_menu.add_separator()
        snap_menu.add_command(label='清理快照...', command=self._cleanup_snapshots)
        menubar.add_cascade(label='快照', menu=snap_menu)

        help_menu = Menu(menubar, tearoff=0)
        help_menu.add_command(label='关于', command=self._show_about)
        menubar.add_cascade(label='帮助', menu=help_menu)

        self.config(menu=menubar)

    # ── Header ──
    def _build_header(self):
        header = Frame(self, bg=COLOR_HEADER, height=50)
        header.pack(fill=X)
        header.pack_propagate(False)

        Label(header, text='Products 数据更新工具', font=FONT_TITLE,
              bg=COLOR_HEADER, fg=COLOR_WHITE).pack(side=LEFT, padx=20, pady=8)

        self.status_label = Label(header, text='就绪', font=FONT_SMALL,
                                  bg=COLOR_HEADER, fg='#95a5a6')
        self.status_label.pack(side=RIGHT, padx=20, pady=8)

    # ── Step Indicator ──
    def _build_step_indicator(self):
        self.steps = StepIndicator(self)
        self.steps.pack(fill=X, padx=20, pady=(15, 5))

    # ── Scrollable Content ──
    def _build_content(self):
        container = Frame(self, bg=COLOR_BG)
        container.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

        self.content_canvas = Canvas(container, bg=COLOR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient=VERTICAL, command=self.content_canvas.yview)
        self.content_frame = Frame(self.content_canvas, bg=COLOR_BG)

        self.content_frame.bind('<Configure>',
            lambda e: self.content_canvas.configure(scrollregion=self.content_canvas.bbox('all')))

        self.content_canvas.create_window((0, 0), window=self.content_frame, anchor='nw', tags='content')
        self.content_canvas.configure(yscrollcommand=scrollbar.set)

        self.content_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        self.content_canvas.bind('<Enter>', lambda e: self._bind_mousewheel())
        self.content_canvas.bind('<Leave>', lambda e: self._unbind_mousewheel())

        container.bind('<Configure>', lambda e: self.content_canvas.itemconfig(
            'content', width=e.width - scrollbar.winfo_reqwidth()))

        # Create step frames
        self.frames = {
            0: FileFrame(self.content_frame, self),
            1: PreviewFrame(self.content_frame, self),
            2: ExecuteFrame(self.content_frame, self),
        }

        self._show_step(0)

    # ── Navigation Bar ──
    def _build_navigation(self):
        nav = Frame(self, bg=COLOR_WHITE, height=45)
        nav.pack(fill=X, side=BOTTOM)
        nav.pack_propagate(False)

        Button(nav, text='撤销 (Undo)', command=self._undo,
               font=FONT_SMALL, relief='flat', padx=10).pack(side=LEFT, padx=20, pady=8)

        self.prev_btn = Button(nav, text='← 上一步', command=self._prev_step,
                               font=FONT_SMALL, relief='flat', padx=15, pady=5, state=DISABLED)
        self.prev_btn.pack(side=RIGHT, padx=5, pady=8)

        self.next_btn = Button(nav, text='下一步 →', command=self._next_step,
                               font=FONT_BOLD, bg=COLOR_PRIMARY, fg=COLOR_WHITE,
                               relief='flat', padx=20, pady=5)
        self.next_btn.pack(side=RIGHT, padx=5, pady=8)

    # ── Step Navigation ──
    def _show_step(self, step: int):
        for s, frame in self.frames.items():
            if s == step:
                frame.pack(fill=BOTH, expand=True)
            else:
                frame.pack_forget()

        self.current_step = step
        self.steps.set_active(step)

        self.prev_btn.config(state=NORMAL if step > 0 else DISABLED)
        self.next_btn.config(text='下一步 →' if step < 2 else '完成',
                             state=NORMAL)
        self._update_status()
        self.content_canvas.yview_moveto(0)

    def _next_step(self):
        if self.current_step == 0:
            if not self.source_path or not os.path.exists(self.source_path):
                messagebox.showwarning('提示', '请先选择有效的源文件')
                return
            if not self.target_path or not os.path.exists(self.target_path):
                messagebox.showwarning('提示', '请先选择有效的目标文件')
                return
            self._show_step(1)

        elif self.current_step == 1:
            if not hasattr(self, 'source_rows') or not self.source_rows:
                messagebox.showwarning('提示', '请先点击"分析变更"查看预览')
                return
            self.frames[2].load_data()
            self._show_step(2)

        elif self.current_step == 2:
            messagebox.showinfo('提示', '已完成向导')

    def _prev_step(self):
        if self.current_step > 0:
            self._show_step(self.current_step - 1)

    # ── Mouse Wheel ──
    def _bind_mousewheel(self):
        self.content_canvas.bind_all('<MouseWheel>', lambda e: self.content_canvas.yview_scroll(
            int(-1 * (e.delta / 120)), 'units'))

    def _unbind_mousewheel(self):
        self.content_canvas.unbind_all('<MouseWheel>')

    # ── Menu Actions ──
    def _pick_source(self):
        self._show_step(0)
        self.frames[0]._pick_source()

    def _pick_target(self):
        self._show_step(0)
        self.frames[0]._pick_target()

    def _manual_snapshot(self):
        if not self.target_path or not os.path.exists(self.target_path):
            messagebox.showwarning('提示', '请先选择有效的目标文件')
            return
        snapshot_dir = os.path.join(os.path.dirname(self.target_path), SNAPSHOT_DIR_NAME)
        path = create_snapshot(self.target_path, snapshot_dir)
        messagebox.showinfo('完成', f'快照已创建:\n{os.path.basename(path)}')
        if hasattr(self.frames[0], 'refresh_snapshots'):
            self.frames[0].refresh_snapshots()

    def _list_snapshots(self):
        if not self.target_path or not os.path.exists(self.target_path):
            messagebox.showwarning('提示', '请先选择有效的目标文件')
            return

        snapshot_dir = os.path.join(os.path.dirname(self.target_path), SNAPSHOT_DIR_NAME)
        snapshots = list_snapshots(snapshot_dir)

        win = Toplevel(self)
        win.title('快照列表')
        win.geometry('600x400')

        text = Text(win, font=FONT_MONO, wrap='none')
        text.pack(fill=BOTH, expand=True, padx=5, pady=5)

        hs = ttk.Scrollbar(win, orient=HORIZONTAL, command=text.xview)
        vs = ttk.Scrollbar(win, orient=VERTICAL, command=text.yview)
        text.configure(xscrollcommand=hs.set, yscrollcommand=vs.set)
        hs.pack(side=BOTTOM, fill=X)
        vs.pack(side=RIGHT, fill=Y)

        if not snapshots:
            text.insert('1.0', '暂无快照')
        else:
            text.insert('1.0', f'共 {len(snapshots)} 个快照\n')
            text.insert(END, f'位置: {snapshot_dir}\n')
            text.insert(END, '=' * 50 + '\n\n')
            for i, sp in enumerate(snapshots):
                mtime = datetime.fromtimestamp(os.path.getmtime(sp))
                size_kb = os.path.getsize(sp) / 1024
                text.insert(END, f'{i + 1}. {os.path.basename(sp)}\n')
                text.insert(END, f'   时间: {mtime:%Y-%m-%d %H:%M:%S}  |  大小: {size_kb:.1f} KB\n\n')

        text.config(state=DISABLED)

    def _undo(self):
        if not self.target_path or not os.path.exists(self.target_path):
            messagebox.showwarning('提示', '请先选择有效的目标文件')
            return

        snapshot_dir = os.path.join(os.path.dirname(self.target_path), SNAPSHOT_DIR_NAME)
        snapshots = list_snapshots(snapshot_dir)
        if not snapshots:
            messagebox.showinfo('提示', '没有可用的快照')
            return

        latest = snapshots[0]
        mtime = datetime.fromtimestamp(os.path.getmtime(latest))
        if messagebox.askyesno('确认撤销',
                               f'将恢复到上一次执行前的状态:\n\n'
                               f'{os.path.basename(latest)}\n'
                               f'时间: {mtime:%Y-%m-%d %H:%M:%S}\n\n'
                               f'恢复前会自动备份当前文件，确定继续？'):
            try:
                create_snapshot(self.target_path, snapshot_dir)
                restore_from_snapshot(self.target_path, latest)
                messagebox.showinfo('完成', f'已从快照恢复:\n{os.path.basename(latest)}')
                self.status_label.config(text='已撤销')
                if hasattr(self.frames[0], 'refresh_snapshots'):
                    self.frames[0].refresh_snapshots()
            except Exception as e:
                messagebox.showerror('撤销失败', str(e))

    def _cleanup_snapshots(self):
        if not self.target_path or not os.path.exists(self.target_path):
            messagebox.showwarning('提示', '请先选择有效的目标文件')
            return
        snapshot_dir = os.path.join(os.path.dirname(self.target_path), SNAPSHOT_DIR_NAME)
        snapshots = list_snapshots(snapshot_dir)
        if len(snapshots) <= 20:
            messagebox.showinfo('提示', f'当前只有 {len(snapshots)} 个快照，无需清理（保留最近 20 个）')
            return
        if messagebox.askyesno('确认清理',
                               f'共 {len(snapshots)} 个快照，将删除最早的 {len(snapshots) - 20} 个，'
                               f'保留最近 20 个。\n\n确定继续？'):
            deleted = cleanup_snapshots(snapshot_dir, keep=20)
            messagebox.showinfo('完成', f'已删除 {deleted} 个旧快照，保留最近 20 个')
            if hasattr(self.frames[0], 'refresh_snapshots'):
                self.frames[0].refresh_snapshots()

    def _show_about(self):
        messagebox.showinfo('关于', 'Products 数据更新工具 v2.0\n\n'
                                    '功能:\n'
                                    '• 读取 Products 后台导出数据\n'
                                    '• 按产品ID匹配更新分析优化表\n'
                                    '• 追加新产品并更新反馈数量合计\n'
                                    '• 每次执行前自动创建快照\n'
                                    '• 支持撤销 (Undo) 恢复')

    def _update_status(self):
        step_names = ['选择文件', '预览变更', '执行更新']
        self.status_label.config(text=f'步骤 {self.current_step + 1}/3: {step_names[self.current_step]}')

    def _on_close(self):
        self.destroy()


# ═══════════════════════════════════════
# Entry Point
# ═══════════════════════════════════════

def main():
    def _excepthook(exc_type, exc_val, exc_tb):
        tb_str = ''.join(traceback.format_exception(exc_type, exc_val, exc_tb))
        messagebox.showerror('程序错误', f'发生未处理的错误:\n\n{exc_val}\n\n{tb_str}')

    sys.excepthook = _excepthook

    app = ProductsUpdateApp()
    app.mainloop()


if __name__ == '__main__':
    main()
