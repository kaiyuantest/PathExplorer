#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys, os, re, json, math, shutil, tempfile, subprocess
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTextEdit, QPushButton, QLabel, QListWidget, QListWidgetItem,
    QTabWidget, QSplitter, QCheckBox, QAbstractItemView,
    QGroupBox, QStatusBar, QMessageBox,
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QTextCharFormat, QSyntaxHighlighter

HISTORY_FILE = os.path.join(os.path.expanduser("~"), ".path_explorer_history.json")

def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_history(h):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(h, f, ensure_ascii=False, indent=2)

def record_path(h, path):
    if path not in h:
        h[path] = {"count": 0, "last": ""}
    h[path]["count"] += 1
    h[path]["last"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_history(h)


def _build_path_re():
    forbidden = (
        chr(0) + "-" + chr(31) +
        "<>:" + chr(34) + "|?*" +
        chr(0x4e00) + "-" + chr(0x9fff) +
        chr(0xff0c) + chr(0x3002) + chr(0xff01) + chr(0xff1f) +
        chr(0x3001) + chr(0xff1b) + chr(0xff1a) +
        chr(0x2018) + chr(0x2019) + chr(0x201c) + chr(0x201d) +
        chr(0xff08) + chr(0xff09) +
        chr(0x3010) + chr(0x3011) + chr(0x300a) + chr(0x300b) +
        " " + chr(9) + chr(13) + chr(10)
    )
    pc = "[^" + forbidden + "]"
    pattern = (
        r'(?<![' + chr(34) + r'\w])' +
        r'([A-Za-z]:[/\\](?:' + pc + r'*(?:[/\\]' + pc + r'*)*)?)' +
        r'(?![' + chr(34) + r'\w])'
    )
    return re.compile(pattern, re.IGNORECASE)

WINDOWS_PATH_RE = _build_path_re()

def extract_paths(text):
    found, seen = [], set()
    strip_chars = chr(92) + "/.,;:" + chr(34) + chr(39) + " "
    for m in WINDOWS_PATH_RE.finditer(text):
        p = m.group(1).rstrip(strip_chars)
        if len(p) >= 3 and p not in seen:
            seen.add(p)
            found.append(p)
    return found


class PathHighlighter(QSyntaxHighlighter):
    def __init__(self, document):
        super().__init__(document)
        self.fmt = QTextCharFormat()
        self.fmt.setForeground(QColor("#4fc3f7"))
        self.fmt.setFontUnderline(True)
    def highlightBlock(self, text):
        for m in WINDOWS_PATH_RE.finditer(text):
            self.setFormat(m.start(), m.end() - m.start(), self.fmt)


class PathItem(QListWidgetItem):
    def __init__(self, path):
        self.path = path
        is_dir  = os.path.isdir(path)
        is_file = os.path.isfile(path)
        exists  = is_dir or is_file
        if is_dir:    icon, kind = "\U0001f4c1", "[文件夹]"
        elif is_file: icon, kind = "\U0001f4c4", "[文件]  "
        else:         icon, kind = "\u2753",     "[不存在]"
        super().__init__("{} {}  {}".format(icon, kind, path))
        self.setCheckState(Qt.Unchecked)
        if not exists:   self.setForeground(QColor("#888888"))
        elif is_dir:     self.setForeground(QColor("#81c784"))
        else:            self.setForeground(QColor("#4fc3f7"))

    @property
    def is_dir(self):  return os.path.isdir(self.path)
    @property
    def is_file(self): return os.path.isfile(self.path)
    @property
    def exists(self):  return os.path.exists(self.path)


# ── Windows Terminal ──────────────────────────
def find_wt_exe():
    candidates = [
        shutil.which("wt"),
        shutil.which("WindowsTerminal.exe"),
        os.path.expandvars(os.path.join("%LOCALAPPDATA%","Microsoft","WindowsApps","wt.exe")),
        os.path.expandvars(os.path.join("%LOCALAPPDATA%","Microsoft","WindowsApps","WindowsTerminal.exe")),
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return None

def _q(p):
    return p.replace(chr(92), "/")

def build_wt_ps1(wt_exe, dirs, profile="Windows PowerShell", delay_ms=500):
    n = len(dirs)
    if n == 1:   rows, cols = 1, 1
    elif n == 2: rows, cols = 1, 2
    elif n <= 4: rows, cols = 2, math.ceil(n/2)
    elif n <= 8: rows, cols = 2, 4
    else:        rows, cols = 3, 4
    wt = _q(wt_exe)
    dq = chr(34)

    def sp(direction, size, d):
        return [
            "& {dq}{wt}{dq} -w 0 split-pane {dir} -s {sz} -p {dq}{prof}{dq} -d {dq}{d2}{dq}".format(
                dq=dq, wt=wt, dir=direction, sz=size, prof=profile, d2=_q(d)),
            "Start-Sleep -Milliseconds {}".format(delay_ms), "",
        ]

    def fp(target):
        return [
            "& {dq}{wt}{dq} -w 0 focus-pane --target {t}".format(dq=dq, wt=wt, t=target),
            "Start-Sleep -Milliseconds {}".format(delay_ms), "",
        ]

    lines = [
        "$ErrorActionPreference = {dq}Continue{dq}".format(dq=dq), "",
        "& {dq}{wt}{dq} -w new new-tab -p {dq}{prof}{dq} -d {dq}{d0}{dq}".format(
            dq=dq, wt=wt, prof=profile, d0=_q(dirs[0])),
        "Start-Sleep -Milliseconds {}".format(delay_ms * 2), "",
    ]
    row1 = min(cols, n)
    for i in range(1, row1):
        lines += sp("-V", round(1.0-1.0/(row1-i+1),6), dirs[i])
    for row in range(1, rows):
        rd = dirs[row*cols: row*cols+cols]
        if not rd: break
        lines += fp((row-1)*cols)
        lines += sp("-H", round(1.0-1.0/(rows-row+1),6), rd[0])
        for j in range(1, len(rd)):
            lines += sp("-V", round(1.0-1.0/(len(rd)-j+1),6), rd[j])
    lines.append("Write-Host {dq}Done: {n} panes ({r}x{c}){dq}".format(dq=dq, n=n, r=rows, c=cols))
    return "\n".join(lines) + "\n"


# ── 主窗口 ───────────────────────────────────
class PathExplorer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.history   = load_history()
        self.all_paths = []
        self._init_ui()
        self._apply_style()

    def _init_ui(self):
        self.setWindowTitle("路径探索器 PathExplorer")
        self.setMinimumSize(960, 680)
        self.resize(1160, 750)
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(12,12,12,8)
        root.setSpacing(8)
        title = QLabel("\U0001f50d  路径探索器")
        title.setObjectName("title")
        root.addWidget(title)
        splitter = QSplitter(Qt.Vertical)
        root.addWidget(splitter, 1)

        ib_box = QGroupBox("粘贴文本（自动识别路径）")
        ib_box.setObjectName("groupBox")
        ib = QVBoxLayout(ib_box)
        ib.setContentsMargins(8,16,8,8)
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("将包含 Windows 路径的文本粘贴到这里...")
        self.text_edit.setObjectName("textEdit")
        self.text_edit.setMinimumHeight(150)
        self.highlighter = PathHighlighter(self.text_edit.document())
        ib.addWidget(self.text_edit)
        br = QHBoxLayout()
        self.btn_parse = QPushButton("\u26a1  识别路径")
        self.btn_parse.setObjectName("btnPrimary")
        self.btn_parse.clicked.connect(self._parse_paths)
        btn_clr = QPushButton("\U0001f5d1  清空")
        btn_clr.setObjectName("btnSecondary")
        btn_clr.clicked.connect(self.text_edit.clear)
        self.lbl_found = QLabel("")
        self.lbl_found.setObjectName("labelSmall")
        br.addWidget(self.btn_parse); br.addWidget(btn_clr)
        br.addStretch(); br.addWidget(self.lbl_found)
        ib.addLayout(br)
        splitter.addWidget(ib_box)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("tabs")
        splitter.addWidget(self.tabs)
        splitter.setSizes([260,440])
        self.tabs.addTab(self._build_paths_tab(), "\U0001f4cb  路径列表")
        self.tabs.addTab(self._build_history_tab(), "\U0001f4ca  打开历史")
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("就绪  ·  粘贴文本后点击「识别路径」")

    def _build_paths_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(8,8,8,8)
        layout.setSpacing(6)
        top = QHBoxLayout()
        self.chk_all = QCheckBox("全选")
        self.chk_all.setObjectName("checkBox")
        self.chk_all.stateChanged.connect(self._select_all)
        self.chk_only_dir = QCheckBox("只看文件夹")
        self.chk_only_dir.setObjectName("checkBox")
        self.chk_only_dir.stateChanged.connect(self._filter_list)
        self.chk_only_file = QCheckBox("只看文件")
        self.chk_only_file.setObjectName("checkBox")
        self.chk_only_file.stateChanged.connect(self._filter_list)
        top.addWidget(self.chk_all); top.addWidget(self.chk_only_dir)
        top.addWidget(self.chk_only_file); top.addStretch()
        layout.addLayout(top)
        self.list_widget = QListWidget()
        self.list_widget.setObjectName("listWidget")
        self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_widget.itemDoubleClicked.connect(self._open_single)
        layout.addWidget(self.list_widget, 1)
        btn_grid = QHBoxLayout()
        btns = [
            ("\U0001f4c2  打开文件夹", "btnAction",  self._open_selected_dirs),
            ("\U0001f4c4  打开文件",   "btnAction",  self._open_selected_files),
            ("\U0001f5a5  打开CMD",    "btnAction",  self._open_cmd),
            ("\u2b1b  合并到WT",       "btnWT",      self._open_wt),
            ("\u2705  打开全部勾选",   "btnPrimary", self._open_checked),
        ]
        for text, obj, slot in btns:
            b = QPushButton(text); b.setObjectName(obj); b.clicked.connect(slot)
            btn_grid.addWidget(b)
        layout.addLayout(btn_grid)
        return w

    def _build_history_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(8,8,8,8); layout.setSpacing(6)
        top = QHBoxLayout()
        lbl = QLabel("按打开次数排名（双击再次打开）"); lbl.setObjectName("labelSmall")
        btn_r = QPushButton("\U0001f504 刷新"); btn_r.setObjectName("btnSecondary")
        btn_r.clicked.connect(self._refresh_history)
        btn_c = QPushButton("\U0001f5d1 清空历史"); btn_c.setObjectName("btnSecondary")
        btn_c.clicked.connect(self._clear_history)
        top.addWidget(lbl); top.addStretch(); top.addWidget(btn_r); top.addWidget(btn_c)
        layout.addLayout(top)
        self.hist_list = QListWidget(); self.hist_list.setObjectName("listWidget")
        self.hist_list.itemDoubleClicked.connect(self._open_hist_item)
        layout.addWidget(self.hist_list, 1)
        self._refresh_history()
        return w

    def _apply_style(self):
        self.setStyleSheet("""
            QMainWindow, QWidget { background-color: #1a1d23; color: #d0d6e0;
                font-family: "Microsoft YaHei UI","Segoe UI",sans-serif; font-size: 13px; }
            #title { font-size: 20px; font-weight: 700; color: #e8eaf0; padding: 4px 2px 8px 2px; }
            #groupBox { border: 1px solid #2d3240; border-radius: 8px; margin-top: 4px;
                padding: 6px; color: #7a8499; font-size: 12px; }
            #groupBox::title { subcontrol-origin: margin; subcontrol-position: top left;
                padding: 0 6px; left: 10px; }
            #textEdit { background-color: #12141a; border: 1px solid #2a2f3d; border-radius: 6px;
                color: #c8d0e0; padding: 8px; font-family: "Consolas",monospace; font-size: 13px; }
            #listWidget { background-color: #12141a; border: 1px solid #2a2f3d; border-radius: 6px;
                color: #c0c8d8; padding: 4px; font-family: "Consolas",monospace; font-size: 12.5px; }
            #listWidget::item { padding: 5px 8px; border-radius: 4px; }
            #listWidget::item:selected { background-color: #1e3a5f; color: #e8f0ff; }
            #listWidget::item:hover { background-color: #1d2535; }
            QPushButton { border-radius: 6px; padding: 7px 12px; font-weight: 600;
                font-size: 12.5px; min-width: 68px; }
            #btnPrimary { background-color: #1e6fc5; color: #fff; border: none; }
            #btnPrimary:hover { background-color: #2580e0; }
            #btnPrimary:pressed { background-color: #155da0; }
            #btnSecondary { background-color: #252b38; color: #a0aac0; border: 1px solid #333a4d; }
            #btnSecondary:hover { background-color: #2e3548; color: #c0cce0; }
            #btnAction { background-color: #1e2a3a; color: #7ab8e8; border: 1px solid #2a3f58; }
            #btnAction:hover { background-color: #243448; color: #9dd0ff; }
            #btnAction:pressed { background-color: #1a2534; }
            #btnWT { background-color: #1a2e1a; color: #69c469; border: 1px solid #2a4a2a; }
            #btnWT:hover { background-color: #1f381f; color: #88e088; }
            #btnWT:pressed { background-color: #162516; }
            #checkBox { color: #8090a8; spacing: 5px; }
            #checkBox::indicator { width: 14px; height: 14px; border: 1px solid #3a4560;
                border-radius: 3px; background: #12141a; }
            #checkBox::indicator:checked { background-color: #1e6fc5; border-color: #1e6fc5; }
            #labelSmall { color: #6a7a94; font-size: 12px; }
            QTabWidget::pane { border: 1px solid #2a2f3d; border-radius: 8px; background: #1a1d23; }
            QTabBar::tab { background: #141720; color: #6a7a94; border: 1px solid #252b38;
                border-bottom: none; border-radius: 6px 6px 0 0; padding: 7px 18px; margin-right: 3px; }
            QTabBar::tab:selected { background: #1a1d23; color: #d0d8ec; border-color: #2a2f3d; }
            QStatusBar { background: #12141a; color: #55657a; font-size: 11.5px;
                border-top: 1px solid #1e2430; }
            QSplitter::handle { background: #252b38; height: 3px; }
            QScrollBar:vertical { background: #12141a; width: 8px; border-radius: 4px; }
            QScrollBar::handle:vertical { background: #2d3a50; border-radius: 4px; min-height: 20px; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
        """)

    def _parse_paths(self):
        self.all_paths = extract_paths(self.text_edit.toPlainText())
        self._apply_filter()
        n = len(self.all_paths)
        dirs  = sum(1 for p in self.all_paths if os.path.isdir(p))
        files = sum(1 for p in self.all_paths if os.path.isfile(p))
        miss  = n - dirs - files
        self.lbl_found.setText(
            "识别到 {} 条  \U0001f4c1{} 文件夹  \U0001f4c4{} 文件  \u2753{} 不存在".format(n,dirs,files,miss))
        self.status.showMessage("识别完成：共 {} 条路径".format(n))

    def _populate_list(self, paths):
        self.list_widget.clear()
        self.chk_all.blockSignals(True); self.chk_all.setCheckState(Qt.Unchecked); self.chk_all.blockSignals(False)
        for path in paths:
            self.list_widget.addItem(PathItem(path))

    def _apply_filter(self):
        od = self.chk_only_dir.isChecked(); of = self.chk_only_file.isChecked()
        p  = self.all_paths
        if od and of:  f = p
        elif od:       f = [x for x in p if os.path.isdir(x)]
        elif of:       f = [x for x in p if os.path.isfile(x)]
        else:          f = p
        self._populate_list(f)

    def _filter_list(self): self._apply_filter()

    def _select_all(self, state):
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(Qt.Checked if state==Qt.Checked else Qt.Unchecked)

    def _get_checked_items(self):
        return [self.list_widget.item(i) for i in range(self.list_widget.count())
                if self.list_widget.item(i).checkState()==Qt.Checked]

    def _get_selected_items(self): return self.list_widget.selectedItems()

    def _unique_dirs(self, targets):
        seen = {}
        for item in targets:
            if not hasattr(item,'path'): continue
            p = item.path
            if os.path.isfile(p): p = os.path.dirname(p)
            if os.path.isdir(p):  seen[p.lower()] = p
        return list(seen.values())

    def _open_path(self, path):
        if not os.path.exists(path):
            self.status.showMessage("路径不存在: "+path); return
        try:
            os.startfile(path); record_path(self.history, path)
            self._refresh_history(); self.status.showMessage("已打开: "+path)
        except Exception as e:
            self.status.showMessage("打开失败: "+str(e))

    def _open_single(self, item):
        if hasattr(item,'path'): self._open_path(item.path)

    def _open_checked(self):
        items = self._get_checked_items()
        if not items: self.status.showMessage("\u26a0  请先勾选要打开的路径"); return
        for item in items: self._open_path(item.path)

    def _open_selected_dirs(self):
        targets = self._get_checked_items() or self._get_selected_items()
        if not targets: self.status.showMessage("\u26a0  请先勾选或选中路径"); return
        for item in targets:
            if not hasattr(item,'path'): continue
            p = item.path
            if os.path.isdir(p):   self._open_path(p)
            elif os.path.isfile(p): self._open_path(os.path.dirname(p))

    def _open_selected_files(self):
        targets = self._get_checked_items() or self._get_selected_items()
        if not targets: self.status.showMessage("\u26a0  请先勾选或选中路径"); return
        cnt = 0
        for item in targets:
            if hasattr(item,'path') and os.path.isfile(item.path):
                self._open_path(item.path); cnt += 1
        if cnt == 0: self.status.showMessage("勾选项中没有文件（只有文件夹）")

    def _open_cmd(self):
        targets = self._get_checked_items() or self._get_selected_items()
        if not targets: self.status.showMessage("\u26a0  请先勾选或选中路径"); return
        dirs = self._unique_dirs(targets)
        if not dirs: self.status.showMessage("\u26a0  勾选路径中没有可访问的目录"); return
        cnt = 0
        for folder in dirs:
            try:
                subprocess.Popen('start cmd /K "cd /d {}"'.format(folder), shell=True)
                record_path(self.history, folder); cnt += 1
            except Exception as e:
                self.status.showMessage("CMD 打开失败: "+str(e))
        self._refresh_history()
        self.status.showMessage("已打开 {} 个 CMD 窗口（已合并重复目录）".format(cnt))

    def _open_wt(self):
        targets = self._get_checked_items() or self._get_selected_items()
        if not targets: self.status.showMessage("\u26a0  请先勾选或选中路径"); return
        dirs = self._unique_dirs(targets)
        if not dirs: self.status.showMessage("\u26a0  勾选路径中没有可访问的目录"); return
        wt_exe = find_wt_exe()
        if not wt_exe:
            self.status.showMessage("\u274c  未找到 Windows Terminal，请先在微软商店安装"); return
        MAX_PANES = 12
        truncated = len(dirs) > MAX_PANES
        dirs = dirs[:MAX_PANES]
        ps1_text = build_wt_ps1(wt_exe, dirs)
        tmp = tempfile.NamedTemporaryFile(suffix=".ps1", mode="wb", delete=False)
        tmp.write(b"\xef\xbb\xbf"); tmp.write(ps1_text.encode("utf-8")); tmp.close()
        try:
            subprocess.Popen(["powershell","-ExecutionPolicy","Bypass","-WindowStyle","Hidden","-File",tmp.name])
            for d in dirs: record_path(self.history, d)
            self._refresh_history()
            msg = "已在 Windows Terminal 中打开 {} 个 pane".format(len(dirs))
            if truncated: msg += "（超出12个已截断）"
            self.status.showMessage(msg)
        except Exception as e:
            self.status.showMessage("WT 启动失败: "+str(e))

    def _refresh_history(self):
        self.hist_list.clear()
        for rank,(path,info) in enumerate(
                sorted(self.history.items(), key=lambda x:x[1].get("count",0), reverse=True), 1):
            count = info.get("count",0); last = info.get("last","")
            icon = "\U0001f4c1" if os.path.isdir(path) else ("\U0001f4c4" if os.path.isfile(path) else "\u2753")
            label = "#{:<3} {}  {}    （{} 次  ·  {}）".format(rank,icon,path,count,last)
            li = QListWidgetItem(label); li.setData(Qt.UserRole, path)
            if not os.path.exists(path): li.setForeground(QColor("#666677"))
            elif count>=5:               li.setForeground(QColor("#ffb74d"))
            else:                        li.setForeground(QColor("#90caf9"))
            self.hist_list.addItem(li)

    def _open_hist_item(self, item):
        p = item.data(Qt.UserRole)
        if p: self._open_path(p)

    def _clear_history(self):
        reply = QMessageBox.question(self,"确认","确定要清空所有历史记录吗？",
                                     QMessageBox.Yes|QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.history.clear(); save_history(self.history)
            self._refresh_history(); self.status.showMessage("历史记录已清空")


def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = PathExplorer(); win.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()