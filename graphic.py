# graphic.py — 双页面向导式 GUI（PySide6）
# Step 1: 仅路径 -> 自动静默检索 -> 进入 Step 2
# Step 2: 显示“识别结果（带数量）”、选择 Mode，并只展开对应表单
# 改动要点：
#   - 新增：类别规范化映射，兼容“斜撑/桁架/Truss”等写法
#   - 新增：顶部“识别结果”标签条（有什么就展示什么）
#   - 改进：Mode2 的“可包含”行带数量，复选框采用蓝色勾选样式，更显眼

from __future__ import annotations
import os, sys, importlib.util, re, copy
from pathlib import Path
from dataclasses import dataclass
import unicodedata
from collections import defaultdict
from typing import Callable
from PySide6.QtCore import Qt, QSize, QThread, Signal, QSettings, QDate, QPoint, QRect
from PySide6.QtGui import QIcon, QPixmap, QColor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QGroupBox, QFileDialog, QRadioButton, QButtonGroup,
    QCheckBox, QMessageBox, QSpacerItem, QSizePolicy, QStackedWidget, QFrame,
    QComboBox, QScrollArea, QSpinBox, QToolButton,
    QTableWidget, QAbstractItemView, QHeaderView, QDateEdit, QLayout,
    QWidgetItem
)
from functools import partial

# ========= ORF 自搜索导入块 =========
def _load_orf_module():
    mod_name = "ORF"
    try:
        return __import__(mod_name)
    except ModuleNotFoundError:
        pass
    start = Path(__file__).resolve().parent
    candidates = []
    p = start
    for _ in range(7):
        candidates += [
            p / "ORF.py",
            p / "before" / "ORF.py",
            p / "src" / "ORF.py",
            p / "convert" / "src" / "ORF.py",
            p / "new" / "convert" / "src" / "ORF.py",
        ]
        p = p.parent
    for f in candidates:
        if f.exists():
            spec = importlib.util.spec_from_file_location(mod_name, str(f))
            mod = importlib.util.module_from_spec(spec)
            sys.modules[mod_name] = mod
            sys.path.insert(0, str(f.parent))
            spec.loader.exec_module(mod)  # type: ignore
            return mod
    raise ModuleNotFoundError("未找到 ORF.py（已在常见位置搜索）。")

_ORF = _load_orf_module()
probe_categories_from_docx = _ORF.probe_categories_from_docx
export_mode2_noninteractive = _ORF.export_mode2_noninteractive
run_noninteractive = _ORF.run_noninteractive
Mode1ConfigProvider = getattr(_ORF, "Mode1ConfigProvider", None)
run_mode1_with_provider = getattr(_ORF, "run_mode1_with_provider", None)
export_mode1_noninteractive = getattr(_ORF, "export_mode1_noninteractive", None)
export_mode4_noninteractive = getattr(_ORF, "export_mode4_noninteractive", None)
prepare_from_word = getattr(_ORF, "prepare_from_word", None)
_distribute_by_dates = getattr(_ORF, "_distribute_by_dates", None)
normalize_date = getattr(_ORF, "normalize_date", lambda x: x)
_floor_label_from_name = getattr(_ORF, "_floor_label_from_name", None)
_floor_sort_key_by_label = getattr(_ORF, "_floor_sort_key_by_label", None)
BACKEND_TITLE = getattr(_ORF, "TITLE", "原始记录自动填写程序")
ORF_LOADED_FROM = getattr(_ORF, "__file__", None)
# ===================================

DEFAULT_START_DIR = r"E:\pycharm first\pythonProject\原始记录自动填写程序\before"
CANON_KEYS = ["钢柱", "钢梁", "支撑", "网架", "其他"]

# —— 同义词映射（可按你后端真实返回再扩充）——
SYNONYMS = {
    "钢柱": {"钢柱", "柱", "H柱", "钢立柱", "Steel Column", "SC"},
    "钢梁": {"钢梁", "梁", "H梁", "主梁", "次梁", "Steel Beam", "SB"},
    "支撑": {"支撑", "斜撑", "撑", "撑杆", "支撑件", "Brace", "Bracing", "Support"},
    "网架": {"网架", "桁架", "Grid", "Truss", "Space Frame", "框架网架"},
    "其他": {"其他", "其它", "杂项", "附件", "Other"},
}

@dataclass
class DocProbeResult:
    categories: list[str]
    counts: dict

# ---------- 简易流式布局 ----------
class FlowLayout(QLayout):
    def __init__(self, parent=None, margin: int = 0, spacing: int = -1):
        super().__init__(parent)
        self._items: list = []
        if parent is not None:
            self.setContentsMargins(margin, margin, margin, margin)
        self.setSpacing(spacing if spacing >= 0 else 6)

    def __del__(self):
        while self.count():
            item = self.takeAt(0)
            if item is not None:
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()

    def addItem(self, item):
        self._items.append(item)

    def addWidget(self, widget):
        self.addChildWidget(widget)
        self.addItem(QWidgetItem(widget))

    def count(self) -> int:
        return len(self._items)

    def itemAt(self, index: int):
        if 0 <= index < len(self._items):
            return self._items[index]
        return None

    def takeAt(self, index: int):
        if 0 <= index < len(self._items):
            return self._items.pop(index)
        return None

    def expandingDirections(self):
        return Qt.Orientations()

    def hasHeightForWidth(self) -> bool:
        return True

    def heightForWidth(self, width: int) -> int:
        height = self._do_layout(QRect(0, 0, width, 0), True)
        return height

    def setGeometry(self, rect: QRect):
        super().setGeometry(rect)
        self._do_layout(rect, False)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.sizeHint())
        margins = self.contentsMargins()
        size += QSize(margins.left() + margins.right(), margins.top() + margins.bottom())
        return size

    def _do_layout(self, rect: QRect, test_only: bool) -> int:
        x = rect.x()
        y = rect.y()
        line_height = 0
        effective_rect = rect.adjusted(
            self.contentsMargins().left(),
            self.contentsMargins().top(),
            -self.contentsMargins().right(),
            -self.contentsMargins().bottom(),
        )
        x = effective_rect.x()
        y = effective_rect.y()
        for item in self._items:
            wid = item.widget()
            if wid is None or not wid.isVisible():
                hint = item.sizeHint()
            else:
                hint = wid.sizeHint()
            space_x = self.spacing()
            space_y = self.spacing()
            next_x = x + hint.width() + space_x
            if next_x - space_x > effective_rect.right() and line_height > 0:
                x = effective_rect.x()
                y = y + line_height + space_y
                next_x = x + hint.width() + space_x
                line_height = 0
            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), hint))
            x = next_x
            line_height = max(line_height, hint.height())
        return y + line_height - rect.y() + self.contentsMargins().bottom()


# ---------- 后台线程：静默检索 ----------
class ProbeThread(QThread):
    done = Signal(object, object)   # (error, result)

    def __init__(self, path: Path):
        super().__init__()
        self.path = path

    def run(self):
        try:
            info = probe_categories_from_docx(self.path)
            res = DocProbeResult(
                categories=list(info.get("categories", [])),
                counts=dict(info.get("counts", {}))
            )
            self.done.emit(None, res)
        except Exception as e:
            self.done.emit(e, None)

# ---------- UI 小工具 ----------
def hline():
    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setFrameShadow(QFrame.Sunken)
    line.setStyleSheet("color:#e6e6e6;")
    return line

# 规范化：把后端返回的各种写法统一到 CANON_KEYS，并合并数量
def normalize_detected(raw_categories: list[str], raw_counts: dict) -> tuple[dict, dict]:
    present = {k: False for k in CANON_KEYS}
    counts  = {k: 0 for k in CANON_KEYS}

    # 先处理 counts（键也可能是同义词）
    for k, v in (raw_counts or {}).items():
        v_int = 0
        try:
            v_int = int(v or 0)
        except Exception:
            v_int = 0
        mapped = None
        for canon, aliases in SYNONYMS.items():
            if k in aliases:
                mapped = canon
                break
        if mapped is None:
            # 尝试直接匹配规范键
            mapped = k if k in CANON_KEYS else "其他"
        counts[mapped] = counts.get(mapped, 0) + v_int
        if v_int > 0:
            present[mapped] = True

    # 再处理 categories（有的后端只给列表）
    for name in (raw_categories or []):
        mapped = None
        for canon, aliases in SYNONYMS.items():
            if name in aliases:
                mapped = canon
                break
        if mapped is None:
            mapped = name if name in CANON_KEYS else "其他"
        present[mapped] = True

    return present, counts

# ---------- 主窗 ----------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{BACKEND_TITLE} · 图形界面")
        self.resize(1100, 700)

        self.settings = QSettings("ORF", "OriginalRecordFilling")  # 组织, 应用名

        self._theme_options = [
            ("蓝色", "#2d89ef"),
            ("绿色", "#34c759"),
            ("黄色", "#f7b500"),
            ("粉色", "#ff2d55"),
            ("橙色", "#ff9500"),
            ("紫色", "#7e57c2"),
        ]

        self.accent = self.settings.value("ui/themeColor", "#2d89ef", type=str)

        self.doc_path: Path | None = None
        self.present = {k: False for k in CANON_KEYS}
        self.counts  = {k: 0 for k in CANON_KEYS}
        self._m1_day_forms: list[dict] = []
        self._floors_by_cat: dict[str, set[str]] = {}
        self._grouped_cache = None
        self._cf_groups_by_floor: dict[tuple[str, str], list] = {}

        self.m4_strategy = self.settings.value("mode4/strategy", "even", type=str) or "even"
        if self.m4_strategy not in {"even", "quota"}:
            self.m4_strategy = "even"
        self.m4_base_entries: list[tuple[str, int | None]] = []
        self.m4_selected_floors: set[str] = set()
        self.m4_applied_floors: set[str] = set()
        self.m4_overrides: dict[str, list[tuple[str, int | None]]] = {}
        self.m4_all_floors: list[str] = []
        self.m4_floor_buttons: dict[str, QToolButton] = {}
        self.m4_floor_rows: dict[str, dict] = {}
        self.m4_preview_text: str = ""
        self._plan_table_callbacks: dict[QTableWidget, Callable[[], None]] = {}

        self.m4_fallback = self.settings.value("mode4/fallback", "append_last", type=str) or "append_last"
        if self.m4_fallback not in {"append_last", "default", "error"}:
            self.m4_fallback = "append_last"
        self.m4_write_dates = bool(self.settings.value("mode4/writeDates", True, type=bool))
        self.m4_support_strategy = self.settings.value("mode4/supportStrategy", "number", type=str) or "number"
        if self.m4_support_strategy not in {"number", "floor"}:
            self.m4_support_strategy = "number"
        self.m4_net_strategy = self.settings.value("mode4/netStrategy", "number", type=str) or "number"
        if self.m4_net_strategy not in {"number", "floor"}:
            self.m4_net_strategy = "number"
        self.m4_include_support = bool(self.settings.value("mode4/includeSupport", True, type=bool))

        self.stack = QStackedWidget()
        self.page_select = self._build_page_select()
        self.page_modes  = self._build_page_modes()
        self.stack.addWidget(self.page_select)
        self.stack.addWidget(self.page_modes)
        self.setCentralWidget(self.stack)

        self._apply_styles()

    # ====== Page 1：仅路径 ======
    def _build_page_select(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w); lay.setContentsMargins(16,16,16,16); lay.setSpacing(12)

        box = QGroupBox("1. 选择 Word 源文件")
        b = QVBoxLayout(box)
        row = QHBoxLayout()
        self.ed_path = QLineEdit(); self.ed_path.setPlaceholderText("请选择 .docx 文件")
        self.btn_browse = QPushButton("浏览…")
        row.addWidget(self.ed_path, 1); row.addWidget(self.btn_browse, 0)
        b.addLayout(row)

        # 颜色选择行
        row_theme = QHBoxLayout()
        row_theme.addWidget(QLabel("界面颜色"))
        self.cmb_theme = QComboBox()

        for name, hx in self._theme_options:
            pm = QPixmap(14, 14)
            pm.fill(QColor(hx))
            self.cmb_theme.addItem(QIcon(pm), name, hx)

        curr = (self.accent or "").lower()
        idx = next((i for i, (_, hx) in enumerate(self._theme_options) if hx.lower() == curr), 0)
        self.cmb_theme.setCurrentIndex(idx)
        row_theme.addWidget(self.cmb_theme)
        row_theme.addStretch(1)
        b.addLayout(row_theme)

        self.lb_status1 = QLabel("就绪"); self.lb_status1.setStyleSheet("color:#777;")
        b.addWidget(self.lb_status1)
        lay.addWidget(box)

        tip = QLabel(f"后端模块：{ORF_LOADED_FROM or '未知'}"); tip.setStyleSheet("color:#999;")
        lay.addWidget(tip); lay.addStretch(1)

        self.btn_browse.clicked.connect(self._on_browse_and_probe)
        self.cmb_theme.currentIndexChanged.connect(self._on_theme_changed)
        return w

    # ====== Page 2：模式选择 + 表单 ======
    def _build_page_modes(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w); lay.setContentsMargins(16,16,16,16); lay.setSpacing(12)

        header = QHBoxLayout()
        self.btn_back = QPushButton("← 返回选择文件"); self.btn_back.setFixedHeight(32)
        self.lb_file_short = QLabel(""); self.lb_file_short.setStyleSheet("color:#666;")
        header.addWidget(self.btn_back, 0); header.addSpacing(8); header.addWidget(self.lb_file_short, 1)
        lay.addLayout(header)
        lay.addWidget(hline())

        # (A) 识别结果标签条（有什么就展示什么 + 数量）
        self.box_found = QGroupBox("识别结果")
        lf = QHBoxLayout(self.box_found)
        self.lb_found = QLabel("（空）"); self.lb_found.setStyleSheet("color:#555;")
        lf.addWidget(self.lb_found, 1)
        lay.addWidget(self.box_found)

        # (B) 模式选择
        mode_box = QGroupBox("2. 选择处理模式")
        lm = QHBoxLayout(mode_box)
        self.rb_m1 = QRadioButton("Mode 1")
        self.rb_m2 = QRadioButton("Mode 2")
        self.rb_m3 = QRadioButton("Mode 3")
        self.rb_m4 = QRadioButton("Mode 4")
        self.rb_m2.setChecked(True)
        self.rb_m4.setEnabled(True)
        self.grp_mode = QButtonGroup(self)
        for i, rb in enumerate([self.rb_m1, self.rb_m2, self.rb_m3, self.rb_m4], start=1):
            self.grp_mode.addButton(rb, i); lm.addWidget(rb)
        lm.addStretch(1)
        lay.addWidget(mode_box)

        # (C) Mode 1 表单
        self.box_m1 = QGroupBox("3A. Mode 1（日期分桶）")
        lm1 = QVBoxLayout(self.box_m1)
        lm1.setSpacing(12)

        bar = QWidget()
        lb = QHBoxLayout(bar)
        lb.setContentsMargins(0, 0, 0, 0)
        lb.setSpacing(12)
        lb.addWidget(QLabel("天数"))
        self.sp_m1_days = QSpinBox()
        self.sp_m1_days.setRange(1, 30)
        self.sp_m1_days.setValue(1)
        self.sp_m1_days.setFixedWidth(80)
        lb.addWidget(self.sp_m1_days)
        lb.addSpacing(12)
        self.lb_m1_sup = QLabel("支撑分桶")
        self.cmb_m1_sup = QComboBox()
        self.cmb_m1_sup.addItems(["编号", "楼层"])
        self.cmb_m1_sup.setCurrentIndex(0)
        self.lb_m1_net = QLabel("网架分桶")
        self.cmb_m1_net = QComboBox()
        self.cmb_m1_net.addItems(["编号", "楼层"])
        self.cmb_m1_net.setCurrentIndex(0)
        for wdg in (self.lb_m1_sup, self.cmb_m1_sup, self.lb_m1_net, self.cmb_m1_net):
            lb.addWidget(wdg)
        lb.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        lm1.addWidget(bar)

        self.scroll_m1_days = QScrollArea()
        self.scroll_m1_days.setWidgetResizable(True)
        self._m1_days_container = QWidget()
        self._m1_days_layout = QVBoxLayout(self._m1_days_container)
        self._m1_days_layout.setContentsMargins(0, 0, 0, 0)
        self._m1_days_layout.setSpacing(10)
        self.scroll_m1_days.setWidget(self._m1_days_container)
        lm1.addWidget(self.scroll_m1_days, 1)

        row_opts = QWidget()
        lo = QHBoxLayout(row_opts)
        lo.setContentsMargins(0, 0, 0, 0)
        lo.setSpacing(16)
        self.ck_m1_later = QCheckBox("后面的日期优先（推荐）")
        self.ck_m1_later.setChecked(True)
        self.ck_m1_merge = QCheckBox("未分配并入最后一天")
        lo.addWidget(self.ck_m1_later)
        lo.addWidget(self.ck_m1_merge)
        lo.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        lm1.addWidget(row_opts)

        row_go_m1 = QWidget()
        lg = QHBoxLayout(row_go_m1)
        lg.setContentsMargins(0, 0, 0, 0)
        lg.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.btn_run_m1 = QPushButton("生成（日期分桶）")
        self.btn_run_m1.setFixedSize(QSize(180, 36))
        lg.addWidget(self.btn_run_m1)
        lm1.addWidget(row_go_m1)

        # (C) Mode 3 表单
        self.box_m3 = QGroupBox("3A. Mode 3（单日模式）")
        lm3 = QVBoxLayout(self.box_m3)
        row_m3 = QHBoxLayout()
        row_m3.addWidget(QLabel("检测日期"))
        self.ed_m3_date = QLineEdit(); self.ed_m3_date.setPlaceholderText("如：2025-10-13 / 20251013 / 10-13 / 2025年10月13日 …")
        row_m3.addWidget(self.ed_m3_date, 1)
        self.btn_run_m3 = QPushButton("生成（单日）")
        row_m3.addWidget(self.btn_run_m3, 0)
        lm3.addLayout(row_m3)

        # (D) Mode 2 表单
        self.box_m2 = QGroupBox("3B. Mode 2（楼层断点）")
        lm2 = QVBoxLayout(self.box_m2)

        row_bp = QHBoxLayout()
        self.lb_bp_common = QLabel("楼层断点（柱/梁）")
        self.ed_bp_common = QLineEdit(); self.ed_bp_common.setPlaceholderText("例：3 6 10（空=不分段）")
        self.lb_bp_hint = QLabel(""); self.lb_bp_hint.setStyleSheet("color:#888;")
        row_bp.addWidget(self.lb_bp_common)
        row_bp.addWidget(self.ed_bp_common, 1)
        row_bp.addWidget(self.lb_bp_hint)

        row_dt = QHBoxLayout()
        row_dt.addWidget(QLabel("前段日期"))
        self.ed_dt_first = QLineEdit(); self.ed_dt_first.setPlaceholderText("如：2025-08-27")
        row_dt.addWidget(self.ed_dt_first)
        row_dt.addSpacing(16)
        row_dt.addWidget(QLabel("后段日期"))
        self.ed_dt_second = QLineEdit(); self.ed_dt_second.setPlaceholderText("如：2025-09-03")
        row_dt.addWidget(self.ed_dt_second)

        row_inc = QHBoxLayout()
        row_inc.addWidget(QLabel("可包含"))
        self.ck_support = QCheckBox("支撑")   # 数量会在文本里补 "(N)"
        row_inc.addWidget(self.ck_support)
        row_inc.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))

        row_strategy = QHBoxLayout()
        self.lb_sup_strategy = QLabel("支撑分段")
        self.cmb_sup_strategy = QComboBox(); self.cmb_sup_strategy.addItems(["编号", "楼层"])
        self.cmb_sup_strategy.setCurrentIndex(0)
        row_strategy.addWidget(self.lb_sup_strategy)
        row_strategy.addWidget(self.cmb_sup_strategy)
        row_strategy.addSpacing(16)
        self.lb_net_strategy = QLabel("网架分段")
        self.cmb_net_strategy = QComboBox(); self.cmb_net_strategy.addItems(["编号", "楼层"])
        self.cmb_net_strategy.setCurrentIndex(0)
        row_strategy.addWidget(self.lb_net_strategy)
        row_strategy.addWidget(self.cmb_net_strategy)
        row_strategy.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))

        row_sup_bp = QHBoxLayout()
        self.lb_sup_bp = QLabel("支撑断点")
        self.ed_bp_sup = QLineEdit(); self.ed_bp_sup.setPlaceholderText("例：3 6 10（空=不分段）")
        row_sup_bp.addWidget(self.lb_sup_bp)
        row_sup_bp.addWidget(self.ed_bp_sup, 1)

        row_net_bp = QHBoxLayout()
        self.lb_net_bp = QLabel("网架断点")
        self.ed_bp_net = QLineEdit(); self.ed_bp_net.setPlaceholderText("例：10 20 30（空=不分段）")
        row_net_bp.addWidget(self.lb_net_bp)
        row_net_bp.addWidget(self.ed_bp_net, 1)

        row_go = QHBoxLayout()
        row_go.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.btn_run_m2 = QPushButton("生成（楼层断点）")
        self.btn_run_m2.setFixedSize(QSize(160, 36))
        row_go.addWidget(self.btn_run_m2)

        for r in (row_bp, row_dt, row_inc, row_strategy, row_sup_bp, row_net_bp, row_go):
            lm2.addLayout(r)

        self.lb_bp_common.setVisible(False)
        self.ed_bp_common.setVisible(False)
        self.lb_bp_hint.setVisible(False)
        self.lb_sup_bp.setVisible(False)
        self.ed_bp_sup.setVisible(False)
        self.ed_bp_sup.setEnabled(False)
        self.lb_net_bp.setVisible(False)
        self.ed_bp_net.setVisible(False)

        self.ck_support.toggled.connect(lambda on: self.ed_bp_sup.setEnabled(on))
        self.cmb_sup_strategy.currentIndexChanged.connect(self._update_sup_bp_ui)
        self.cmb_net_strategy.currentIndexChanged.connect(self._update_net_bp_ui)
        self._update_sup_bp_ui()
        self._update_net_bp_ui()

                # (E) Mode 4 表单
        self.box_m4 = QGroupBox("3C. Mode 4（多日按楼层计划）")
        lm4 = QHBoxLayout(self.box_m4)
        lm4.setSpacing(16)

        col_left = QVBoxLayout()
        col_left.setSpacing(12)
        lm4.addLayout(col_left, 3)

        self.lb_m4_hint = QLabel("给一批楼层套一张日程表；少数楼层再单独覆写。")
        self.lb_m4_hint.setStyleSheet("color:#555;")
        col_left.addWidget(self.lb_m4_hint)

        self.grp_m4_step1 = QGroupBox("Step 1｜设计划（总表）")
        step1 = QVBoxLayout(self.grp_m4_step1)
        step1.setSpacing(8)

        row_strategy = QHBoxLayout()
        row_strategy.addWidget(QLabel("分配策略"))
        self.rb_m4_even = QRadioButton("均分")
        self.rb_m4_quota = QRadioButton("配额（每日上限）")
        if self.m4_strategy == "quota":
            self.rb_m4_quota.setChecked(True)
        else:
            self.rb_m4_even.setChecked(True)
        row_strategy.addWidget(self.rb_m4_even)
        row_strategy.addWidget(self.rb_m4_quota)
        row_strategy.addStretch(1)
        step1.addLayout(row_strategy)

        self.tbl_m4_base = QTableWidget()
        self._init_plan_table(self.tbl_m4_base, on_change=self._on_base_plan_changed)
        step1.addWidget(self.tbl_m4_base)

        row_plan_btn = QHBoxLayout()
        self.btn_m4_add_base = QPushButton("+ 添加日期")
        self.btn_m4_copy_base = QPushButton("复制上一行")
        self.btn_m4_del_base = QPushButton("删除所选")
        for btn in (self.btn_m4_add_base, self.btn_m4_copy_base, self.btn_m4_del_base):
            row_plan_btn.addWidget(btn)
        row_plan_btn.addStretch(1)
        step1.addLayout(row_plan_btn)

        self.lb_m4_base_hint = QLabel("留空上限 = 均分；最后一天自动吃掉余量。")
        self.lb_m4_base_hint.setStyleSheet("color:#888; font-size:12px;")
        step1.addWidget(self.lb_m4_base_hint)

        col_left.addWidget(self.grp_m4_step1)

        self.grp_m4_step2 = QGroupBox("Step 2｜选楼层并应用")
        step2 = QVBoxLayout(self.grp_m4_step2)
        step2.setSpacing(8)

        ctrl_row = QHBoxLayout()
        self.btn_m4_filter_all = QPushButton("全选")
        self.btn_m4_filter_none = QPushButton("清空")
        self.btn_m4_filter_basement = QPushButton("仅地下")
        self.btn_m4_filter_digits = QPushButton("仅数字层")
        self.btn_m4_filter_me = QPushButton("机房")
        self.btn_m4_filter_roof = QPushButton("屋面")
        for btn in (
            self.btn_m4_filter_all,
            self.btn_m4_filter_none,
            self.btn_m4_filter_basement,
            self.btn_m4_filter_digits,
            self.btn_m4_filter_me,
            self.btn_m4_filter_roof,
        ):
            btn.setFixedHeight(28)
            ctrl_row.addWidget(btn)
        ctrl_row.addStretch(1)
        step2.addLayout(ctrl_row)

        self.m4_floor_chip_container = QWidget()
        self.m4_floor_chips = FlowLayout(self.m4_floor_chip_container)
        self.m4_floor_chips.setContentsMargins(0, 0, 0, 0)
        self.m4_floor_chips.setSpacing(6)
        self.m4_floor_chip_container.setLayout(self.m4_floor_chips)
        step2.addWidget(self.m4_floor_chip_container)

        self.lb_m4_floors = QLabel("")
        self.lb_m4_floors.setStyleSheet("color:#888; font-size:12px;")
        step2.addWidget(self.lb_m4_floors)

        row_m4_cats = QHBoxLayout()
        row_m4_cats.addWidget(QLabel("类别"))
        self.sw_m4_cat_gz = QCheckBox("钢柱")
        self.sw_m4_cat_gl = QCheckBox("钢梁")
        self.sw_m4_cat_sup = QCheckBox("支撑")
        self.sw_m4_cat_net = QCheckBox("网架")
        for sw in (
            self.sw_m4_cat_gz,
            self.sw_m4_cat_gl,
            self.sw_m4_cat_sup,
            self.sw_m4_cat_net,
        ):
            row_m4_cats.addWidget(sw)
        row_m4_cats.addStretch(1)
        step2.addLayout(row_m4_cats)

        self.btn_m4_apply_plan = QPushButton("将上方计划应用到这些楼层")
        self.btn_m4_apply_plan.setFixedHeight(34)
        step2.addWidget(self.btn_m4_apply_plan)

        self.scroll_m4_applied = QScrollArea()
        self.scroll_m4_applied.setWidgetResizable(True)
        self.w_m4_applied = QWidget()
        self.lay_m4_applied = QVBoxLayout(self.w_m4_applied)
        self.lay_m4_applied.setContentsMargins(0, 0, 0, 0)
        self.lay_m4_applied.setSpacing(6)
        self.lay_m4_applied.addStretch(1)
        self.scroll_m4_applied.setWidget(self.w_m4_applied)
        step2.addWidget(self.scroll_m4_applied)

        col_left.addWidget(self.grp_m4_step2, 1)

        self.grp_m4_step3 = QGroupBox("Step 3｜未分配处理")
        step3 = QVBoxLayout(self.grp_m4_step3)
        step3.setSpacing(8)

        row_fb = QHBoxLayout()
        row_fb.addWidget(QLabel("未分配"))
        self.rb_m4_fb_append = QRadioButton("并入最后一天")
        self.rb_m4_fb_default = QRadioButton("回落到 Mode 1（日期分桶）")
        self.rb_m4_fb_error = QRadioButton("停止并提示")
        fb_map = {
            "append_last": self.rb_m4_fb_append,
            "default": self.rb_m4_fb_default,
            "error": self.rb_m4_fb_error,
        }
        fb_map.get(self.m4_fallback, self.rb_m4_fb_append).setChecked(True)
        for rb in (self.rb_m4_fb_append, self.rb_m4_fb_default, self.rb_m4_fb_error):
            row_fb.addWidget(rb)
        row_fb.addStretch(1)
        step3.addLayout(row_fb)

        self.w_m4_default = QWidget()
        lay_def = QVBoxLayout(self.w_m4_default)
        lay_def.setContentsMargins(0, 0, 0, 0)
        lay_def.setSpacing(6)
        self.tbl_m4_default = QTableWidget()
        self._init_plan_table(self.tbl_m4_default, on_change=self._on_default_plan_changed)
        lay_def.addWidget(self.tbl_m4_default)
        row_def_btn = QHBoxLayout()
        self.btn_m4_def_add = QPushButton("+ 添加日期")
        self.btn_m4_def_copy = QPushButton("复制上一行")
        self.btn_m4_def_del = QPushButton("删除所选")
        for btn in (self.btn_m4_def_add, self.btn_m4_def_copy, self.btn_m4_def_del):
            row_def_btn.addWidget(btn)
        row_def_btn.addStretch(1)
        lay_def.addLayout(row_def_btn)
        step3.addWidget(self.w_m4_default)

        row_support = QHBoxLayout()
        self.ck_m4_support = QCheckBox("包含支撑")
        self.ck_m4_support.setChecked(self.m4_include_support)
        self.lb_m4_sup_strategy = QLabel("支撑分桶")
        self.cmb_m4_sup_strategy = QComboBox()
        self.cmb_m4_sup_strategy.addItems(["按编号", "按楼层"])
        self.cmb_m4_sup_strategy.setCurrentIndex(1 if self.m4_support_strategy == "floor" else 0)
        self.lb_m4_net_strategy = QLabel("网架分桶")
        self.cmb_m4_net_strategy = QComboBox()
        self.cmb_m4_net_strategy.addItems(["按编号", "按楼层"])
        self.cmb_m4_net_strategy.setCurrentIndex(1 if self.m4_net_strategy == "floor" else 0)
        for w in (
            self.ck_m4_support,
            self.lb_m4_sup_strategy,
            self.cmb_m4_sup_strategy,
            self.lb_m4_net_strategy,
            self.cmb_m4_net_strategy,
        ):
            row_support.addWidget(w)
        row_support.addStretch(1)
        step3.addLayout(row_support)

        self.ck_m4_write_dates = QCheckBox("写入日期到表头")
        self.ck_m4_write_dates.setChecked(self.m4_write_dates)
        step3.addWidget(self.ck_m4_write_dates)

        row_go_m4 = QHBoxLayout()
        self.btn_run_m4 = QPushButton("生成（Mode 4）")
        self.btn_run_m4.setFixedSize(QSize(200, 40))
        row_go_m4.addWidget(self.btn_run_m4)
        self.lb_m4_summary = QLabel("")
        self.lb_m4_summary.setStyleSheet("color:#555;")
        row_go_m4.addStretch(1)
        row_go_m4.addWidget(self.lb_m4_summary)
        step3.addLayout(row_go_m4)

        col_left.addWidget(self.grp_m4_step3)

        preview_box = QGroupBox("实时预览")
        lp = QVBoxLayout(preview_box)
        lp.setSpacing(8)
        self.lb_m4_preview = QLabel("请先设计划")
        self.lb_m4_preview.setWordWrap(True)
        self.lb_m4_preview.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        lp.addWidget(self.lb_m4_preview)
        lm4.addWidget(preview_box, 2)

# 容器：只显示当前模式对应的表单
        self.panel_host = QWidget()
        self.panel_wrap = QVBoxLayout(self.panel_host)
        self.panel_wrap.setContentsMargins(0, 0, 0, 0)
        self.panel_wrap.setSpacing(8)
        self.panel_wrap.addWidget(self.box_m1)
        self.panel_wrap.addWidget(self.box_m2)  # 默认显示 M2
        self.panel_wrap.addWidget(self.box_m3)
        self.panel_wrap.addWidget(self.box_m4)
        self.box_m1.setVisible(False)
        self.box_m3.setVisible(False)
        self.box_m4.setVisible(False)

        lay.addWidget(self.panel_host)
        lay.addStretch(1)



        lay.addWidget(hline())
        self.status = QLabel("准备就绪"); self.status.setStyleSheet("color:#555;")
        lay.addWidget(self.status)

        # 事件
        self.btn_back.clicked.connect(self._go_back_to_select)
        self.grp_mode.idToggled.connect(self._on_mode_switched)
        self.sp_m1_days.valueChanged.connect(self._on_days_changed)
        self.btn_run_m1.clicked.connect(self._on_run_mode1)
        self.btn_run_m2.clicked.connect(self._on_run_mode2)
        self.btn_run_m3.clicked.connect(self._on_run_mode3)
        self.btn_run_m4.clicked.connect(self._on_run_mode4)
        self.rb_m4_even.toggled.connect(lambda on: on and self._on_m4_strategy_changed("even"))
        self.rb_m4_quota.toggled.connect(lambda on: on and self._on_m4_strategy_changed("quota"))
        self.btn_m4_add_base.clicked.connect(lambda: self._plan_table_add_row(self.tbl_m4_base))
        self.btn_m4_copy_base.clicked.connect(lambda: self._plan_table_copy_last(self.tbl_m4_base))
        self.btn_m4_del_base.clicked.connect(lambda: self._plan_table_remove_selected(self.tbl_m4_base))
        self.btn_m4_filter_all.clicked.connect(lambda: self._m4_apply_filter("all"))
        self.btn_m4_filter_none.clicked.connect(lambda: self._m4_apply_filter("none"))
        self.btn_m4_filter_basement.clicked.connect(lambda: self._m4_apply_filter("basement"))
        self.btn_m4_filter_digits.clicked.connect(lambda: self._m4_apply_filter("digits"))
        self.btn_m4_filter_me.clicked.connect(lambda: self._m4_apply_filter("me"))
        self.btn_m4_filter_roof.clicked.connect(lambda: self._m4_apply_filter("roof"))
        self.btn_m4_apply_plan.clicked.connect(self._on_apply_plan_clicked)
        self.btn_m4_def_add.clicked.connect(lambda: self._plan_table_add_row(self.tbl_m4_default))
        self.btn_m4_def_copy.clicked.connect(lambda: self._plan_table_copy_last(self.tbl_m4_default))
        self.btn_m4_def_del.clicked.connect(lambda: self._plan_table_remove_selected(self.tbl_m4_default))
        self.rb_m4_fb_append.toggled.connect(lambda _: self._on_m4_fallback_changed())
        self.rb_m4_fb_default.toggled.connect(lambda _: self._on_m4_fallback_changed())
        self.rb_m4_fb_error.toggled.connect(lambda _: self._on_m4_fallback_changed())
        self.ck_m4_support.toggled.connect(self._on_m4_support_toggled)
        self.cmb_m4_sup_strategy.currentIndexChanged.connect(self._on_support_strategy_changed)
        self.cmb_m4_net_strategy.currentIndexChanged.connect(self._on_net_strategy_changed)
        self.ck_m4_write_dates.toggled.connect(self._on_m4_write_dates_changed)
        for sw in (self.sw_m4_cat_gz, self.sw_m4_cat_gl, self.sw_m4_cat_sup, self.sw_m4_cat_net):
            sw.toggled.connect(self._on_categories_changed)

        self._apply_detection_to_mode1_ui()
        self._on_support_strategy_changed()
        self._on_net_strategy_changed()
        self._on_m4_fallback_changed()
        self._refresh_m4_default_visibility()
        self._refresh_m4_strategy_ui()

        return w

        # ====== 样式（增加 QCheckBox 的蓝色勾） ======
    def _on_theme_changed(self, idx: int):
        hx = self.cmb_theme.itemData(idx)
        if isinstance(hx, str) and hx.startswith("#"):
            self.accent = hx
            self._apply_styles()  # 重新套样式
            self.settings.setValue("ui/themeColor", self.accent)
            self.settings.sync()

    def _apply_styles(self):
        c = self.accent
        self.setStyleSheet(f"""
    QWidget {{ background:#ffffff; color:#333; font-size:14px; }}
                QGroupBox {{
                    border:1px solid #e7e7e7; border-radius:12px; margin-top:12px; padding:12px;
                    font-weight:600;
                }}
                QGroupBox::title {{ subcontrol-origin: margin; left:12px; padding:0 6px; background:transparent; }}
                QLineEdit {{
                    height:34px; border:1px solid #d9d9d9; border-radius:8px; padding:4px 10px; background:#fafafa;
                }}
                QPushButton {{
                    height:34px; border:1px solid #d9d9d9; border-radius:10px; background:#f6f6f6; padding:0 12px;
                }}
                QPushButton:hover {{ background:#efefef; }}

                /* —— 单选圆点 —— */
                QRadioButton {{ spacing:8px; }}
                QRadioButton::indicator {{
                    width:14px; height:14px; border-radius:7px;
                    border:2px solid #9aa0a6; background:#fff; margin-right:6px;
                }}
                QRadioButton::indicator:hover {{ border-color:{c}; }}
                QRadioButton::indicator:checked {{
                    background:{c}; border:2px solid {c};
                }}
                QRadioButton:checked {{ color:{c}; font-weight:700; }}

                /* —— 复选框 —— */
                QCheckBox::indicator {{
                    width:16px; height:16px; border-radius:4px;
                    border:2px solid #9aa0a6; background:#fff; margin-right:6px;
                }}
                QCheckBox::indicator:hover {{ border-color:{c}; }}
                QCheckBox::indicator:checked {{
                    image: none; background:{c}; border:2px solid {c};
                }}
        """)

    def closeEvent(self, e):
        self.settings.setValue("ui/themeColor", self.accent)
        self.settings.sync()
        super().closeEvent(e)

    # ====== Step1：选择并静默检索 ======
    def _on_browse_and_probe(self):
        start_dir = DEFAULT_START_DIR if Path(DEFAULT_START_DIR).exists() else str(Path.cwd())
        file, _ = QFileDialog.getOpenFileName(self, "选择 Word 文件", start_dir, "Word 文档 (*.docx)")
        if not file:
            return
        self.ed_path.setText(file)
        fp = Path(file)
        if not (fp.exists() and fp.suffix.lower() == ".docx"):
            QMessageBox.warning(self, "提示", "请选择有效的 .docx 文件。")
            return

        self.doc_path = fp
        self._grouped_cache = None
        self._floors_by_cat = {}
        self._reset_m4_plan_state()
        self.lb_status1.setText("🔎 正在分析文档…")
        self.btn_browse.setEnabled(False)

        self.th = ProbeThread(fp)
        self.th.done.connect(self._on_probe_done_step1)
        self.th.start()

    def _on_probe_done_step1(self, err, res: DocProbeResult | None):
        self.btn_browse.setEnabled(True)
        if err:
            QMessageBox.critical(self, "检索失败", f"读取文档出错：\n{err}")
            self.lb_status1.setText("❌ 检索失败，请重新选择文件。")
            return

        self.present, self.counts = normalize_detected(res.categories, res.counts)

        # 切到 Step 2，并按检索结果刷新 UI
        self._apply_detection_to_mode1_ui()
        self._apply_detection_to_mode2_ui()
        self._ensure_floor_info()
        self._apply_detection_to_mode4_ui()
        self._update_m4_floor_hint()
        self._refresh_found_bar()
        self.lb_file_short.setText(f"文件：{self.doc_path.name}")
        self.status.setText("✅ 已分析完成，可选择模式继续")
        self.stack.setCurrentIndex(1)

    # ====== Step2：模式切换 & 表单显隐 ======
    def _on_mode_switched(self, _id: int, checked: bool):
        if not checked:
            return
        current = self.grp_mode.checkedButton()
        self.box_m1.setVisible(current is self.rb_m1)
        self.box_m2.setVisible(current is self.rb_m2)
        self.box_m3.setVisible(current is self.rb_m3)
        self.box_m4.setVisible(current is self.rb_m4)


    # 顶部“识别结果”标签条
    def _refresh_found_bar(self):
        parts = []
        for k in CANON_KEYS:
            if self.present.get(k, False):
                num = self.counts.get(k, 0)
                parts.append(f"{k}（{num}）" if num else f"{k}")
        self.lb_found.setText("、".join(parts) if parts else "未识别到有效构件")

    def _apply_detection_to_mode1_ui(self):
        if not hasattr(self, "sp_m1_days"):
            return

        sup_ok = self.present.get("支撑", False)
        net_ok = self.present.get("网架", False)

        self.lb_m1_sup.setVisible(sup_ok)
        self.cmb_m1_sup.setVisible(sup_ok)
        if not sup_ok:
            self.cmb_m1_sup.setCurrentIndex(0)

        self.lb_m1_net.setVisible(net_ok)
        self.cmb_m1_net.setVisible(net_ok)
        if not net_ok:
            self.cmb_m1_net.setCurrentIndex(0)

        self._rebuild_m1_day_forms(self.sp_m1_days.value())

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
                continue
            child_layout = item.layout()
            if child_layout is not None:
                self._clear_layout(child_layout)

    def _on_days_changed(self, n: int):
        self._rebuild_m1_day_forms(n)

    def _rebuild_m1_day_forms(self, n: int):
        if not hasattr(self, "_m1_days_layout"):
            return

        self._clear_layout(self._m1_days_layout)
        self._m1_day_forms = []

        rule_placeholder = "例：1-3 5 屋面；* 或 全部=全接收；空=不接收"
        date_placeholder = "支持 2025-10-16 / 20251016 / 10-16 / 2025年10月16日"

        for idx in range(max(0, n)):
            box = QGroupBox(f"Day #{idx + 1}")
            box_lay = QVBoxLayout(box)
            box_lay.setContentsMargins(12, 12, 12, 12)
            box_lay.setSpacing(10)

            def add_rule_row(label_text: str, placeholder: str = "") -> QLineEdit:
                row = QWidget()
                row_lay = QHBoxLayout(row)
                row_lay.setContentsMargins(0, 0, 0, 0)
                row_lay.setSpacing(8)
                lb = QLabel(label_text)
                lb.setMinimumWidth(120)
                row_lay.addWidget(lb, 0)
                edit = QLineEdit()
                if placeholder:
                    edit.setPlaceholderText(placeholder)
                row_lay.addWidget(edit, 1)
                box_lay.addWidget(row)
                return edit

            ed_date = add_rule_row("日期", date_placeholder)
            form_entry: dict[str, QLineEdit] = {"date": ed_date}

            for part in ("钢柱", "钢梁", "支撑"):
                if self.present.get(part, False):
                    form_entry[part] = add_rule_row(f"{part} 规则", rule_placeholder)

            if self.present.get("网架", False):
                form_entry["网架_xx"] = add_rule_row("网架（XX）", rule_placeholder)
                form_entry["网架_fg"] = add_rule_row("网架（FG）", rule_placeholder)
                form_entry["网架_sx"] = add_rule_row("网架（SX）", rule_placeholder)
                form_entry["网架_gen"] = add_rule_row("网架（泛称）", rule_placeholder)

            self._m1_days_layout.addWidget(box)
            self._m1_day_forms.append(form_entry)

        self._m1_days_layout.addSpacerItem(
            QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding)
        )

    @staticmethod
    def _to_rule(value: str) -> dict:
        raw_text = (value or "").strip()
        normalized = unicodedata.normalize("NFKC", raw_text).strip()
        if not normalized:
            return {"enabled": False, "ranges": None}

        normalized_cf = normalized.casefold()
        if normalized in {"*", "全部", "所有"} or normalized_cf == "all":
            return {"enabled": True, "ranges": [], "explicit_all": True}

        return {"enabled": True, "ranges": normalized}

    def _apply_detection_to_mode2_ui(self):
        gz_ok = self.present.get("钢柱", False)
        gl_ok = self.present.get("钢梁", False)

        show_common = gz_ok or gl_ok

        if not show_common:
            self.box_m2.setDisabled(True)
            self.status.setText("⚠ 未识别到钢柱/钢梁，Mode 2 可能不可用。")
        else:
            self.box_m2.setDisabled(False)

        hint = "未识别到钢柱/钢梁"
        if gz_ok and gl_ok:
            hint = "识别到：钢柱 + 钢梁（共用断点）"
        elif gz_ok:
            hint = "识别到：钢柱（共用断点）"
        elif gl_ok:
            hint = "识别到：钢梁（共用断点）"

        self.lb_bp_hint.setText(hint)
        self.lb_bp_common.setVisible(show_common)
        self.ed_bp_common.setVisible(show_common)
        self.lb_bp_hint.setVisible(show_common)

        sup_ok = self.present.get("支撑", False)
        num_sup = self.counts.get("支撑", 0)
        self.ck_support.setVisible(sup_ok)
        self.ck_support.setEnabled(sup_ok)
        self.ck_support.setChecked(sup_ok)
        self.ck_support.setText("支撑" if num_sup == 0 else f"支撑（{num_sup}）")
        self.lb_sup_strategy.setVisible(sup_ok)
        self.cmb_sup_strategy.setVisible(sup_ok)
        self.lb_sup_bp.setVisible(sup_ok)
        self.ed_bp_sup.setVisible(sup_ok)
        if sup_ok:
            self.ed_bp_sup.setEnabled(self.ck_support.isChecked())
        else:
            self.cmb_sup_strategy.setCurrentIndex(0)
            self.ed_bp_sup.setEnabled(False)
            self.ed_bp_sup.clear()

        net_ok = self.present.get("网架", False)
        self.lb_net_strategy.setVisible(net_ok)
        self.cmb_net_strategy.setVisible(net_ok)
        self.lb_net_bp.setVisible(net_ok)
        self.ed_bp_net.setVisible(net_ok)
        if not net_ok:
            self.cmb_net_strategy.setCurrentIndex(0)
            self.ed_bp_net.clear()

            self._update_sup_bp_ui()
            self._update_net_bp_ui()

    def _update_sup_bp_ui(self):
        if not hasattr(self, "cmb_sup_strategy"):
            return
        if self.cmb_sup_strategy.currentIndex() == 1:
            self.lb_sup_bp.setText("支撑断点（楼层）")
            self.ed_bp_sup.setPlaceholderText("例：3 6 10（空=不分段）")
        else:
            self.lb_sup_bp.setText("支撑断点（编号）")
            self.ed_bp_sup.setPlaceholderText("例：10 20 30（空=不分段）")

    def _update_net_bp_ui(self):
        if not hasattr(self, "cmb_net_strategy"):
            return
        if self.cmb_net_strategy.currentIndex() == 1:
            self.lb_net_bp.setText("网架断点（楼层）")
            self.ed_bp_net.setPlaceholderText("例：3 6 10（空=不分段）")
        else:
            self.lb_net_bp.setText("网架断点（编号）")
            self.ed_bp_net.setPlaceholderText("例：10 20 30（空=不分段）")

    def _ensure_floor_info(self):
        if not hasattr(self, "lb_m4_floors"):
            return
        if self.doc_path is None or prepare_from_word is None:
            self._floors_by_cat = {}
            self._cf_groups_by_floor = {}
            return
        if self._grouped_cache is not None and self._floors_by_cat:
            return
        try:
            grouped, _cats = prepare_from_word(self.doc_path)
        except Exception:
            self._grouped_cache = None
            self._floors_by_cat = {}
            self._cf_groups_by_floor = {}
            return
        self._grouped_cache = grouped
        floors: dict[str, set[str]] = {}
        cf_groups: dict[tuple[str, str], list] = defaultdict(list)
        for cat, groups in (grouped or {}).items():
            labels = set()
            for g in groups:
                name = ""
                try:
                    name = g.get("name", "")  # type: ignore[call-arg]
                except Exception:
                    name = ""
                label = None
                if _floor_label_from_name:
                    try:
                        label = _floor_label_from_name(name)
                    except Exception:
                        label = None
                if label and label != "F?":
                    labels.add(label)
                    cf_groups[(cat, label)].append(g)
            if labels:
                floors[cat] = labels
        self._floors_by_cat = floors
        self._cf_groups_by_floor = cf_groups

    def _apply_detection_to_mode4_ui(self):
        if not hasattr(self, "m4_floor_chips"):
            return
        gz_ok = self.present.get("钢柱", False)
        gl_ok = self.present.get("钢梁", False)
        sup_ok = self.present.get("支撑", False)
        net_ok = self.present.get("网架", False)

        for ok, widget in [
            (gz_ok, getattr(self, "sw_m4_cat_gz", None)),
            (gl_ok, getattr(self, "sw_m4_cat_gl", None)),
            (sup_ok, getattr(self, "sw_m4_cat_sup", None)),
            (net_ok, getattr(self, "sw_m4_cat_net", None)),
        ]:
            if isinstance(widget, QCheckBox):
                widget.setVisible(ok)
                widget.setEnabled(ok)
                if ok and not widget.isChecked():
                    widget.setChecked(True)
                if not ok:
                    widget.setChecked(False)

        if isinstance(self.ck_m4_support, QCheckBox):
            self.ck_m4_support.setVisible(sup_ok)
            self.ck_m4_support.setEnabled(sup_ok)
            self.ck_m4_support.setChecked(self.m4_include_support and sup_ok)
            self._on_m4_support_toggled(self.ck_m4_support.isChecked())

        if isinstance(self.cmb_m4_sup_strategy, QComboBox):
            self.cmb_m4_sup_strategy.setCurrentIndex(1 if self.m4_support_strategy == "floor" else 0)
        if isinstance(self.cmb_m4_net_strategy, QComboBox):
            self.cmb_m4_net_strategy.setCurrentIndex(1 if self.m4_net_strategy == "floor" else 0)

        sorter = _floor_sort_key_by_label or (lambda x: x)
        floors_set: set[str] = set()
        for cat in ("钢柱", "钢梁", "支撑", "网架"):
            floors_set.update(self._floors_by_cat.get(cat, set()))
        floors = sorted(floors_set, key=sorter)
        self.m4_all_floors = floors
        self._rebuild_m4_floor_chips(floors)
        self._m4_set_selected_floors(set(floors))
        self.m4_applied_floors = set(floors)
        self._sync_m4_applied_rows()
        self._update_m4_floor_hint()
        self._refresh_m4_default_visibility()
        self._refresh_m4_strategy_ui()
        self._refresh_m4_preview()

    def _update_m4_floor_hint(self):
        if not hasattr(self, "lb_m4_floors"):
            return
        if not self._floors_by_cat:
            self.lb_m4_floors.setText("（楼层信息将在读取后显示）")
            return
        sorter = _floor_sort_key_by_label or (lambda x: x)
        parts = []
        for cat in ("钢柱", "钢梁", "支撑", "网架"):
            floors = sorted(self._floors_by_cat.get(cat, []), key=sorter)
            if floors:
                parts.append(f"{cat}：{' '.join(floors)}")
        self.lb_m4_floors.setText(" | ".join(parts))

    def _init_plan_table(self, table: QTableWidget, *, on_change: Callable[[], None] | None = None):
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["日期", "上限"])
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        table.verticalHeader().setVisible(False)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setAlternatingRowColors(True)
        if on_change:
            self._plan_table_callbacks[table] = on_change
        else:
            self._plan_table_callbacks.pop(table, None)

    def _plan_table_add_row(self, table: QTableWidget, entry: tuple[str, int] | None = None, *, suppress_trigger: bool = False):
        row = table.rowCount()
        table.insertRow(row)

        date_edit = QDateEdit()
        date_edit.setCalendarPopup(True)
        date_edit.setDisplayFormat("yyyy-M-d")
        if entry and entry[0]:
            parsed = QDate.fromString(entry[0], "yyyy-M-d")
            if parsed.isValid():
                date_edit.setDate(parsed)
            else:
                date_edit.setDate(QDate.currentDate())
        else:
            date_edit.setDate(QDate.currentDate())

        limit_spin = QSpinBox()
        limit_spin.setRange(0, 999999)
        limit_spin.setSingleStep(5)
        limit_spin.setSpecialValueText("不限")
        if entry is not None:
            try:
                limit_spin.setValue(max(0, int(entry[1])))
            except Exception:
                limit_spin.setValue(0)
        else:
            limit_spin.setValue(0)

        date_edit.dateChanged.connect(lambda _=None, tbl=table: self._trigger_plan_change(tbl))
        limit_spin.valueChanged.connect(lambda _=None, tbl=table: self._trigger_plan_change(tbl))

        table.setCellWidget(row, 0, date_edit)
        table.setCellWidget(row, 1, limit_spin)
        table.setRowHeight(row, 32)
        if not suppress_trigger:
            self._trigger_plan_change(table)

    def _plan_table_set_entries(self, table: QTableWidget, entries: list[tuple[str, int]], *, suppress_trigger: bool = False):
        block = table.blockSignals(True)
        table.setRowCount(0)
        for entry in entries or []:
            self._plan_table_add_row(table, entry, suppress_trigger=True)
        table.blockSignals(block)
        if not suppress_trigger:
            self._trigger_plan_change(table)

    def _plan_table_collect(self, table: QTableWidget) -> list[tuple[str, int]]:
        results: list[tuple[str, int]] = []
        for row in range(table.rowCount()):
            date_edit = table.cellWidget(row, 0)
            limit_widget = table.cellWidget(row, 1)
            if not isinstance(date_edit, QDateEdit) or not isinstance(limit_widget, QSpinBox):
                continue
            date_str = date_edit.date().toString("yyyy-M-d")
            results.append((date_str, int(limit_widget.value())))
        return results

    def _plan_table_remove_selected(self, table: QTableWidget):
        selected_rows = sorted({idx.row() for idx in table.selectedIndexes()}, reverse=True)
        if not selected_rows and table.rowCount() > 0:
            selected_rows = [table.rowCount() - 1]
        for row in selected_rows:
            table.removeRow(row)
        self._trigger_plan_change(table)

    def _plan_table_copy_last(self, table: QTableWidget):
        if table.rowCount() == 0:
            self._plan_table_add_row(table)
            return
        last = table.rowCount() - 1
        date_edit = table.cellWidget(last, 0)
        limit_widget = table.cellWidget(last, 1)
        entry: tuple[str, int] | None = None
        if isinstance(date_edit, QDateEdit) and isinstance(limit_widget, QSpinBox):
            entry = (date_edit.date().toString("yyyy-M-d"), int(limit_widget.value()))
        self._plan_table_add_row(table, entry)

    def _trigger_plan_change(self, table: QTableWidget):
        cb = self._plan_table_callbacks.get(table)
        if cb:
            cb()
    def _rebuild_m4_floor_chips(self, floors: list[str]):
        if not hasattr(self, "m4_floor_chips"):
            return
        while self.m4_floor_chips.count():
            item = self.m4_floor_chips.takeAt(0)
            if item:
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
        self.m4_floor_buttons = {}
        for floor in floors:
            btn = QToolButton()
            btn.setText(floor)
            btn.setCheckable(True)
            btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
            btn.setMinimumWidth(56)
            btn.toggled.connect(lambda on, name=floor: self._m4_on_floor_chip_toggled(name, on))
            self.m4_floor_chips.addWidget(btn)
            self.m4_floor_buttons[floor] = btn
        if floors:
            self._m4_set_selected_floors(self.m4_selected_floors & set(floors))

    def _reset_m4_plan_state(self):
        self.m4_selected_floors = set()
        self.m4_applied_floors = set()
        self.m4_overrides = {}
        self.m4_floor_buttons = {}
        self.m4_floor_rows = {}
        self.m4_preview_text = ""
        self.m4_base_entries = []
        if hasattr(self, "tbl_m4_base"):
            self._plan_table_set_entries(self.tbl_m4_base, [], suppress_trigger=True)
            self._plan_table_add_row(self.tbl_m4_base, suppress_trigger=True)
            self._trigger_plan_change(self.tbl_m4_base)
        if hasattr(self, "tbl_m4_default"):
            self._plan_table_set_entries(self.tbl_m4_default, [], suppress_trigger=True)
        if hasattr(self, "lay_m4_applied"):
            while self.lay_m4_applied.count() > 1:
                item = self.lay_m4_applied.takeAt(0)
                if item and item.widget():
                    item.widget().deleteLater()
        self._refresh_m4_summary_label()
        self._refresh_m4_preview()

    def _m4_apply_filter(self, kind: str):
        if not self.m4_all_floors:
            return
        kind = (kind or "all").lower()
        if kind == "none":
            selected: set[str] = set()
        elif kind == "basement":
            selected = {f for f in self.m4_all_floors if f.upper().startswith("B")}
        elif kind == "digits":
            selected = {f for f in self.m4_all_floors if re.fullmatch(r"\d+F", f.upper())}
        elif kind == "me":
            selected = {f for f in self.m4_all_floors if "机房" in f}
        elif kind == "roof":
            selected = {f for f in self.m4_all_floors if "屋面" in f}
        else:
            selected = set(self.m4_all_floors)
        self._m4_set_selected_floors(selected)

    def _m4_set_selected_floors(self, floors: set[str]):
        floors = floors & set(self.m4_all_floors)
        self.m4_selected_floors = floors
        for floor, btn in self.m4_floor_buttons.items():
            block = btn.blockSignals(True)
            btn.setChecked(floor in floors)
            btn.blockSignals(block)

    def _m4_on_floor_chip_toggled(self, name: str, checked: bool):
        if checked:
            self.m4_selected_floors.add(name)
        else:
            self.m4_selected_floors.discard(name)

    def _create_floor_row(self, floor: str) -> QFrame:
        name = floor

        frame = QFrame()
        frame.setFrameShape(QFrame.StyledPanel)
        frame_lay = QVBoxLayout(frame)
        frame_lay.setContentsMargins(12, 8, 12, 8)
        frame_lay.setSpacing(6)

        header = QHBoxLayout()
        badge = QLabel(floor)
        badge.setStyleSheet("background:#eef2ff; padding:2px 10px; border-radius:10px; font-weight:600;")
        header.addWidget(badge)
        summary = QLabel("未设置")
        summary.setStyleSheet("color:#555;")
        header.addWidget(summary, 1)
        toggle = QToolButton()
        toggle.setText("自定义")
        toggle.setCheckable(True)
        header.addWidget(toggle)
        header.addStretch(1)
        frame_lay.addLayout(header)

        editor = QWidget()
        editor_lay = QVBoxLayout(editor)
        editor_lay.setContentsMargins(0, 0, 0, 0)
        editor_lay.setSpacing(6)
        table = QTableWidget()
        self._init_plan_table(table, on_change=lambda floor=name: self._on_floor_plan_changed(floor))
        editor_lay.addWidget(table)
        btn_row = QHBoxLayout()
        btn_add = QPushButton("+ 添加日期")
        btn_copy = QPushButton("复制上一行")
        btn_del = QPushButton("删除所选")
        for btn in (btn_add, btn_copy, btn_del):
            btn_row.addWidget(btn)
        btn_row.addStretch(1)
        editor_lay.addLayout(btn_row)
        editor.setVisible(False)
        frame_lay.addWidget(editor)

        btn_add.clicked.connect(lambda _, tbl=table: self._plan_table_add_row(tbl))
        btn_copy.clicked.connect(lambda _, tbl=table: self._plan_table_copy_last(tbl))
        btn_del.clicked.connect(lambda _, tbl=table: self._plan_table_remove_selected(tbl))
        toggle.toggled.connect(partial(self._on_floor_customize_toggled, name))

        self.m4_floor_rows[name] = {
            "frame": frame,
            "summary": summary,
            "toggle": toggle,
            "table": table,
            "editor": editor,
        }
        return frame

    def _sync_m4_applied_rows(self):
        if not hasattr(self, "lay_m4_applied"):
            return
        current = set(self.m4_floor_rows.keys())
        for floor in list(current - self.m4_applied_floors):
            info = self.m4_floor_rows.pop(floor, None)
            if info and info.get("frame"):
                info["frame"].setParent(None)
                info["frame"].deleteLater()
        for floor in sorted(self.m4_applied_floors):
            if floor not in self.m4_floor_rows:
                frame = self._create_floor_row(floor)
                self.lay_m4_applied.insertWidget(self.lay_m4_applied.count() - 1, frame)
        for floor in self.m4_applied_floors:
            self._update_floor_summary(floor)
        self._refresh_m4_strategy_ui()
        self._refresh_m4_summary_label()

    def _update_floor_summary(self, floor: str):
        info = self.m4_floor_rows.get(floor)
        if not info:
            return
        summary = info.get("summary")
        if not isinstance(summary, QLabel):
            return
        entries = self.m4_overrides.get(floor) or self.m4_base_entries
        if not entries:
            summary.setText("未设置")
            return
        text = self._format_plan_summary(entries)
        if floor in self.m4_overrides:
            summary.setText(f"自定义：{text}")
        else:
            summary.setText(f"继承：{text}")

    def _format_plan_summary(self, entries: list[tuple[str, int]]) -> str:
        parts = []
        for date, limit in entries:
            if self.m4_strategy == "even":
                parts.append(f"{date}(均分)")
            else:
                parts.append(f"{date}({limit if limit > 0 else '不限'})")
        return "，".join(parts)

    def _on_floor_customize_toggled(self, name: str, checked: bool):
        info = self.m4_floor_rows.get(name)
        if not info:
            return
        editor = info.get("editor")
        toggle = info.get("toggle")
        table = info.get("table")
        if isinstance(toggle, QToolButton):
            toggle.setText("收起" if checked else "自定义")
        if isinstance(editor, QWidget):
            editor.setVisible(checked)
        if not isinstance(table, QTableWidget):
            return
        if checked:
            entries = self.m4_overrides.get(name) or (self.m4_base_entries if self.m4_base_entries else [])
            self._plan_table_set_entries(table, entries, suppress_trigger=True)
            self._trigger_plan_change(table)
        else:
            self._plan_table_set_entries(table, [], suppress_trigger=True)
            self.m4_overrides.pop(name, None)
            self._trigger_plan_change(table)
        self._update_floor_summary(name)

    def _on_floor_plan_changed(self, floor: str):
        info = self.m4_floor_rows.get(floor)
        if not info:
            return
        table = info.get("table")
        if not isinstance(table, QTableWidget):
            return
        entries = self._plan_table_collect(table)
        if entries and entries != self.m4_base_entries:
            self.m4_overrides[floor] = entries
        else:
            self.m4_overrides.pop(floor, None)
        self._update_floor_summary(floor)
        self._refresh_m4_preview()

    def _on_base_plan_changed(self):
        if not hasattr(self, "tbl_m4_base"):
            return
        self.m4_base_entries = self._plan_table_collect(self.tbl_m4_base)
        for floor in self.m4_applied_floors:
            if floor not in self.m4_overrides:
                self._update_floor_summary(floor)
        self._refresh_m4_strategy_ui()
        self._refresh_m4_preview()

    def _on_default_plan_changed(self):
        self._refresh_m4_preview()

    def _on_apply_plan_clicked(self):
        if not self.m4_selected_floors:
            QMessageBox.warning(self, "提示", "请至少选择一个楼层。")
            return
        base_entries = self._plan_table_collect(self.tbl_m4_base)
        if not base_entries:
            QMessageBox.warning(self, "提示", "请先在上方设置日期计划。")
            return
        self.m4_base_entries = base_entries
        self.m4_applied_floors = set(self.m4_selected_floors)
        for floor in list(self.m4_overrides.keys()):
            if floor not in self.m4_applied_floors:
                self.m4_overrides.pop(floor, None)
        self._sync_m4_applied_rows()
        self._refresh_m4_preview()

    def _on_m4_strategy_changed(self, mode: str):
        if mode not in {"even", "quota"}:
            return
        if self.m4_strategy == mode:
            return
        self.m4_strategy = mode
        self.settings.setValue("mode4/strategy", mode)
        self._refresh_m4_strategy_ui()
        self._refresh_m4_preview()

    def _refresh_m4_strategy_ui(self):
        is_quota = self.m4_strategy == "quota"
        if hasattr(self, "lb_m4_base_hint"):
            if is_quota:
                self.lb_m4_base_hint.setText("配额：填入每日上限；最后一天自动吃掉余量。")
            else:
                self.lb_m4_base_hint.setText("均分：系统按天平均分配，最后一天自动吃掉余量。")
        tables: list[QTableWidget] = []
        if hasattr(self, "tbl_m4_base"):
            tables.append(self.tbl_m4_base)
        if hasattr(self, "tbl_m4_default"):
            tables.append(self.tbl_m4_default)
        tables += [info.get("table") for info in self.m4_floor_rows.values()]
        for table in tables:
            if isinstance(table, QTableWidget):
                table.setColumnHidden(1, not is_quota)
        self._refresh_m4_summary_label()

    def _refresh_m4_default_visibility(self):
        if not hasattr(self, "w_m4_default"):
            return
        show = self.m4_fallback == "default"
        self.w_m4_default.setVisible(show)
        if show and hasattr(self, "tbl_m4_default") and self.tbl_m4_default.rowCount() == 0:
            self._plan_table_add_row(self.tbl_m4_default)

    def _on_m4_fallback_changed(self):
        prev = self.m4_fallback
        if hasattr(self, "rb_m4_fb_default") and self.rb_m4_fb_default.isChecked():
            self.m4_fallback = "default"
        elif hasattr(self, "rb_m4_fb_error") and self.rb_m4_fb_error.isChecked():
            self.m4_fallback = "error"
        else:
            self.m4_fallback = "append_last"
        if prev != self.m4_fallback:
            self.settings.setValue("mode4/fallback", self.m4_fallback)
        self._refresh_m4_default_visibility()
        self._refresh_m4_preview()

    def _on_support_strategy_changed(self):
        if not hasattr(self, "cmb_m4_sup_strategy"):
            return
        self.m4_support_strategy = "floor" if self.cmb_m4_sup_strategy.currentIndex() == 1 else "number"
        self.settings.setValue("mode4/supportStrategy", self.m4_support_strategy)
        self._refresh_m4_preview()

    def _on_net_strategy_changed(self):
        if not hasattr(self, "cmb_m4_net_strategy"):
            return
        self.m4_net_strategy = "floor" if self.cmb_m4_net_strategy.currentIndex() == 1 else "number"
        self.settings.setValue("mode4/netStrategy", self.m4_net_strategy)
        self._refresh_m4_preview()

    def _on_m4_support_toggled(self, checked: bool):
        sup_ok = self.present.get("支撑", False)
        self.m4_include_support = checked and sup_ok
        self.settings.setValue("mode4/includeSupport", self.m4_include_support)
        if hasattr(self, "lb_m4_sup_strategy"):
            self.lb_m4_sup_strategy.setVisible(sup_ok)
            self.cmb_m4_sup_strategy.setVisible(sup_ok)
            self.lb_m4_sup_strategy.setEnabled(self.m4_include_support)
            self.cmb_m4_sup_strategy.setEnabled(self.m4_include_support)
        if hasattr(self, "sw_m4_cat_sup") and sup_ok:
            self.sw_m4_cat_sup.setEnabled(self.m4_include_support)
            if not self.m4_include_support:
                self.sw_m4_cat_sup.setChecked(False)
            elif not self.sw_m4_cat_sup.isChecked():
                self.sw_m4_cat_sup.setChecked(True)
        self._refresh_m4_preview()

    def _on_m4_write_dates_changed(self, checked: bool):
        self.m4_write_dates = bool(checked)
        self.settings.setValue("mode4/writeDates", self.m4_write_dates)

    def _on_categories_changed(self):
        self._refresh_m4_summary_label()
        self._refresh_m4_preview()

    def _current_categories(self) -> list[str]:
        cats: list[str] = []
        mapping = [
            ("钢柱", getattr(self, "sw_m4_cat_gz", None)),
            ("钢梁", getattr(self, "sw_m4_cat_gl", None)),
            ("支撑", getattr(self, "sw_m4_cat_sup", None)),
            ("网架", getattr(self, "sw_m4_cat_net", None)),
        ]
        for name, widget in mapping:
            if isinstance(widget, QCheckBox) and widget.isVisible() and widget.isEnabled() and widget.isChecked():
                cats.append(name)
        if not self.m4_include_support and "支撑" in cats:
            cats.remove("支撑")
        return cats

    def _convert_entries_for_strategy(self, entries: list[tuple[str, int]]) -> list[tuple[str, int | None]]:
        if self.m4_strategy == "even":
            return [(date, None) for date, _ in entries]
        return [(date, int(limit)) for date, limit in entries]

    def _collect_m4_plan_from_ui(self) -> dict:
        categories = self._current_categories()
        if not categories or not self.m4_applied_floors:
            return {}
        base_entries = self.m4_base_entries or self._plan_table_collect(self.tbl_m4_base)
        if not base_entries:
            return {}
        plan_base = self._convert_entries_for_strategy(base_entries)
        plan: dict[str, dict[str, list[tuple[str, int | None]]]] = {}
        for cat in categories:
            plan[cat] = {}
            for floor in sorted(self.m4_applied_floors):
                override = self.m4_overrides.get(floor)
                if override:
                    plan[cat][floor] = self._convert_entries_for_strategy(override)
                else:
                    plan[cat][floor] = list(plan_base)
        return plan

    def _refresh_m4_summary_label(self):
        if not hasattr(self, "lb_m4_summary"):
            return
        cat_count = len(self._current_categories())
        floor_count = len(self.m4_applied_floors)
        if not cat_count or not floor_count:
            self.lb_m4_summary.setText("尚未应用楼层计划")
        else:
            self.lb_m4_summary.setText(f"将应用到 {floor_count} 个楼层，涵盖 {cat_count} 类")

    def _simulate_distribution(self, items, plan_entries):
        if not _distribute_by_dates:
            return []
        try:
            return _distribute_by_dates(items, plan_entries)
        except Exception:
            return []

    def _refresh_m4_preview(self):
        if not hasattr(self, "lb_m4_preview"):
            return
        plan = self._collect_m4_plan_from_ui()
        if not plan:
            self.lb_m4_preview.setText("请先设计划并应用到楼层。")
            self._refresh_m4_summary_label()
            return
        summary_by_date: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
        leftover: dict[str, int] = defaultdict(int)
        for (cat, floor), items in self._cf_groups_by_floor.items():
            if cat not in plan:
                continue
            plan_for_floor = plan[cat].get(floor)
            if not plan_for_floor:
                leftover[cat] += len(items)
                continue
            distribution = self._simulate_distribution(list(items), plan_for_floor)
            assigned = 0
            for date, slice_items in distribution:
                assigned += len(slice_items)
                summary_by_date[date][cat] += len(slice_items)
            if len(items) > assigned:
                leftover[cat] += len(items) - assigned
        if not summary_by_date:
            self.lb_m4_preview.setText("暂无可分配的构件。")
        else:
            def _date_key(value: str):
                qd = QDate.fromString(value, "yyyy-M-d")
                return qd.toJulianDay() if qd.isValid() else value
            lines: list[str] = []
            for idx, date in enumerate(sorted(summary_by_date.keys(), key=_date_key)):
                cats_line = "｜".join(f"{cat}{cnt}组" for cat, cnt in summary_by_date[date].items() if cnt)
                if not cats_line:
                    cats_line = "无任务"
                lines.append(f"第{idx + 1}天（{date}）：{cats_line}")
            left_total = sum(leftover.values())
            if left_total:
                parts = [f"{cat}{cnt}组" for cat, cnt in leftover.items() if cnt]
                policy = {
                    "append_last": "（生成时并入最后一天）",
                    "default": "（生成时使用默认计划）",
                    "error": "（生成时会终止提示）",
                }.get(self.m4_fallback, "")
                lines.append(f"未分配：{left_total}组（{'、'.join(parts)}）{policy}")
            self.lb_m4_preview.setText("\n".join(lines))
        self._refresh_m4_summary_label()

    def _on_run_mode4(self):
        if not export_mode4_noninteractive:
            QMessageBox.critical(self, "提示", "后端暂不支持 Mode 4 生成接口。")
            return
        if not self.doc_path:
            QMessageBox.warning(self, "提示", "请先选择 Word 源文件。")
            return

        plan = self._collect_m4_plan_from_ui()
        if not plan:
            QMessageBox.warning(self, "提示", "请先完成计划并应用到楼层。")
            return

        sup_strategy = self.m4_support_strategy
        net_strategy = self.m4_net_strategy
        fallback = self.m4_fallback

        default_entries = None
        if fallback == "default":
            raw_default = self._plan_table_collect(self.tbl_m4_default)
            if not raw_default:
                QMessageBox.warning(self, "提示", "请填写未分配时使用的默认计划。")
                return
            default_entries = self._convert_entries_for_strategy(raw_default)

        include_support = self.m4_include_support

        self.status.setText("⏳ 正在生成（Mode 4）…")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            xlsx, word = export_mode4_noninteractive(
                src_docx=str(self.doc_path),
                meta={},
                plan=plan,
                include_support=include_support,
                support_strategy=sup_strategy,
                net_strategy=net_strategy,
                fallback=fallback,
                default_entries=default_entries,
                write_dates_to_header=self.m4_write_dates,
            )
            QMessageBox.information(
                self,
                "完成",
                f"✅ 生成完成！\nExcel：{xlsx}\n汇总Word：{word}",
            )
            self.status.setText("✅ Mode 4 完成")
        except Exception as e:
            QMessageBox.critical(
                self,
                "失败",
                f"生成失败：\n{e}",
            )
            self.status.setText("❌ 生成失败")
        finally:
            QApplication.restoreOverrideCursor()

    # ====== 生成：Mode 3 ======
    def _on_run_mode3(self):
        if not self.doc_path:
            QMessageBox.warning(self, "提示", "请先选择 Word 源文件。"); return
        dt = (self.ed_m3_date.text() or "").strip()
        meta = {}
        self.status.setText("⏳ 正在生成（单日模式）…")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            out = run_noninteractive(src_path=str(self.doc_path), mode=3, meta=meta, single_date=dt)
            xlsx = out.get("excel"); word = out.get("word")
            QMessageBox.information(self, "完成", f"✅ 生成完成！\nExcel：{xlsx}\n汇总Word：{word}")
            self.status.setText("✅ 单日模式完成")
        except Exception as e:
            QMessageBox.critical(self, "失败", f"生成失败：\n{e}")
            self.status.setText("❌ 生成失败")
        finally:
            QApplication.restoreOverrideCursor()

    # ====== 生成：Mode 2 ======
    def _on_run_mode2(self):
        if not self.doc_path:
            QMessageBox.warning(self, "提示", "请先选择 Word 源文件。");
            return

        bp_common = (self.ed_bp_common.text() or "").strip() if self.ed_bp_common.isVisible() else ""
        bp_sup = ""
        if self.ed_bp_sup.isVisible() and self.ed_bp_sup.isEnabled():
            bp_sup = (self.ed_bp_sup.text() or "").strip()
        bp_net = (self.ed_bp_net.text() or "").strip() if self.ed_bp_net.isVisible() else ""
        dt_first = (self.ed_dt_first.text() or "").strip()
        dt_second = (self.ed_dt_second.text() or "").strip()

        inc_support = self.ck_support.isVisible() and self.ck_support.isChecked()

        sup_strategy = "number"
        if self.cmb_sup_strategy.isVisible() and self.cmb_sup_strategy.currentIndex() == 1:
            sup_strategy = "floor"

        net_strategy = "number"
        if self.cmb_net_strategy.isVisible() and self.cmb_net_strategy.currentIndex() == 1:
            net_strategy = "floor"

        meta = {}

        self.status.setText("⏳ 正在生成（楼层断点）…")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            out = export_mode2_noninteractive(
                src_docx=str(self.doc_path),
                meta=meta,
                breaks_gz=bp_common,
                breaks_gl=bp_common,
                breaks_support=bp_sup,
                breaks_net=bp_net,
                date_first=dt_first,
                date_second=dt_second,
                include_support=inc_support,
                support_strategy=sup_strategy,
                net_strategy=net_strategy,
            )
            xlsx = out.get("excel");
            word = out.get("word")
            if xlsx:
                QMessageBox.information(self, "完成", f"✅ 生成完成！\nExcel：{xlsx}\n汇总Word：{word}")
            self.status.setText("✅ 楼层断点完成")
        except Exception as e:
            QMessageBox.critical(self, "失败", f"生成失败：\n{e}")
            self.status.setText("❌ 生成失败")
        finally:
            QApplication.restoreOverrideCursor()


def main():
    try:
        from PySide6.QtCore import Qt as _Qt
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            getattr(_Qt, "HighDpiScaleFactorRoundingPolicy").PassThrough
        )
    except Exception:
        pass

    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
