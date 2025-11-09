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
from PySide6.QtCore import Qt, QSize, QThread, Signal, QSettings, QDate, QPoint, QRect
from PySide6.QtGui import QIcon, QPixmap, QColor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QGroupBox, QFileDialog, QRadioButton, QButtonGroup,
    QCheckBox, QMessageBox, QSpacerItem, QSizePolicy, QStackedWidget, QFrame,
    QComboBox, QScrollArea, QSpinBox, QToolButton, QListWidget,
    QListWidgetItem, QTableWidget, QAbstractItemView, QHeaderView, QDateEdit, QLayout,
    QWidgetItem
)

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

        self.m4_selected_floors: set[str] = set()
        self.m4_shared_mode: bool = True
        self.m4_entries_shared: list[tuple[str, int | None]] = []
        self.m4_entries_by_floor: dict[str, list[tuple[str, int | None]]] = {}
        self._cur_m4_floor: str | None = None
        self.m4_all_floors: list[str] = []
        self.m4_floor_buttons: dict[str, QToolButton] = {}

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
        lm4 = QVBoxLayout(self.box_m4)
        lm4.setSpacing(10)

        self.lb_m4_hint = QLabel("请选择楼层并为所需类别配置日期与上限计划。")
        self.lb_m4_hint.setStyleSheet("color:#555;")
        lm4.addWidget(self.lb_m4_hint)

        row_m4_floor_ctrl = QHBoxLayout()
        row_m4_floor_ctrl.addWidget(QLabel("楼层"))
        self.btn_m4_floor_all = QPushButton("全选")
        self.btn_m4_floor_none = QPushButton("全不选")
        self.btn_m4_floor_base = QPushButton("仅 B 层")
        self.btn_m4_floor_std = QPushButton("标准层")
        for btn in (
            self.btn_m4_floor_all,
            self.btn_m4_floor_none,
            self.btn_m4_floor_base,
            self.btn_m4_floor_std,
        ):
            btn.setFixedHeight(28)
            row_m4_floor_ctrl.addWidget(btn)
        row_m4_floor_ctrl.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        lm4.addLayout(row_m4_floor_ctrl)

        self.m4_floor_chip_container = QWidget()
        self.m4_floor_chips = FlowLayout(self.m4_floor_chip_container)
        self.m4_floor_chips.setContentsMargins(0, 0, 0, 0)
        self.m4_floor_chips.setSpacing(6)
        self.m4_floor_chip_container.setLayout(self.m4_floor_chips)
        lm4.addWidget(self.m4_floor_chip_container)

        self.lb_m4_floors = QLabel("")
        self.lb_m4_floors.setStyleSheet("color:#888; font-size:12px;")
        lm4.addWidget(self.lb_m4_floors)

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
        row_m4_cats.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        lm4.addLayout(row_m4_cats)

        row_m4_opts = QHBoxLayout()
        self.lb_m4_sup_strategy = QLabel("支撑分段")
        self.cmb_m4_sup_strategy = QComboBox(); self.cmb_m4_sup_strategy.addItems(["编号", "楼层"])
        self.lb_m4_net_strategy = QLabel("网架分段")
        self.cmb_m4_net_strategy = QComboBox(); self.cmb_m4_net_strategy.addItems(["编号", "楼层"])
        self.ck_m4_support = QCheckBox("包含支撑")
        self.ck_m4_support.setChecked(True)
        for wdg in (
            self.lb_m4_sup_strategy,
            self.cmb_m4_sup_strategy,
            self.lb_m4_net_strategy,
            self.cmb_m4_net_strategy,
            self.ck_m4_support,
        ):
            row_m4_opts.addWidget(wdg)
        row_m4_opts.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        lm4.addLayout(row_m4_opts)

        plan_box = QGroupBox("计划编辑")
        plan_lay = QVBoxLayout(plan_box)
        plan_lay.setSpacing(8)

        row_mode = QHBoxLayout()
        self.rb_m4_shared = QRadioButton("共用计划")
        self.rb_m4_byfloor = QRadioButton("分楼层计划")
        self.rb_m4_shared.setChecked(True)
        row_mode.addWidget(self.rb_m4_shared)
        row_mode.addWidget(self.rb_m4_byfloor)
        row_mode.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        plan_lay.addLayout(row_mode)

        plan_body = QHBoxLayout()
        self.lv_m4_floors = QListWidget()
        self.lv_m4_floors.setSelectionMode(QAbstractItemView.SingleSelection)
        self.lv_m4_floors.setFixedWidth(140)
        plan_body.addWidget(self.lv_m4_floors)

        self.tbl_m4_plan = QTableWidget()
        self._init_plan_table(self.tbl_m4_plan)
        plan_body.addWidget(self.tbl_m4_plan, 1)
        plan_lay.addLayout(plan_body)

        row_plan_btn = QHBoxLayout()
        self.btn_m4_addrow = QPushButton("+ 添加日期")
        self.btn_m4_delrow = QPushButton("- 删除所选")
        self.btn_m4_even = QPushButton("均分上限")
        self.btn_m4_copy2all = QPushButton("复制到已选楼层")
        for btn in (
            self.btn_m4_addrow,
            self.btn_m4_delrow,
            self.btn_m4_even,
            self.btn_m4_copy2all,
        ):
            row_plan_btn.addWidget(btn)
        row_plan_btn.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        plan_lay.addLayout(row_plan_btn)
        self.btn_m4_copy2all.hide()

        lm4.addWidget(plan_box)

        row_m4_fallback = QHBoxLayout()
        row_m4_fallback.addWidget(QLabel("未分配处理"))
        self.cmb_m4_fallback = QComboBox()
        self.cmb_m4_fallback.addItems(["并入最后一天", "使用默认计划", "报错"])
        row_m4_fallback.addWidget(self.cmb_m4_fallback)
        row_m4_fallback.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        lm4.addLayout(row_m4_fallback)

        self.w_m4_default = QWidget()
        lay_def = QVBoxLayout(self.w_m4_default)
        lay_def.setContentsMargins(0, 0, 0, 0)
        lay_def.setSpacing(8)
        self.tbl_m4_default = QTableWidget()
        self._init_plan_table(self.tbl_m4_default)
        lay_def.addWidget(self.tbl_m4_default)
        row_def_btn = QHBoxLayout()
        self.btn_m4_def_add = QPushButton("+ 添加日期")
        self.btn_m4_def_del = QPushButton("- 删除所选")
        self.btn_m4_def_even = QPushButton("均分上限")
        for btn in (self.btn_m4_def_add, self.btn_m4_def_del, self.btn_m4_def_even):
            row_def_btn.addWidget(btn)
        row_def_btn.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        lay_def.addLayout(row_def_btn)
        lm4.addWidget(self.w_m4_default)
        self.w_m4_default.setVisible(False)

        row_go_m4 = QHBoxLayout()
        row_go_m4.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.btn_run_m4 = QPushButton("生成（Mode 4）")
        self.btn_run_m4.setFixedSize(QSize(180, 36))
        row_go_m4.addWidget(self.btn_run_m4)
        lm4.addLayout(row_go_m4)

        # 容器：只显示当前模式对应的表单
        self.panel_wrap = QVBoxLayout()
        self.panel_wrap.addWidget(self.box_m1)
        self.panel_wrap.addWidget(self.box_m2)  # 默认显示 M2
        self.panel_wrap.addWidget(self.box_m3)
        self.panel_wrap.addWidget(self.box_m4)
        self.box_m1.setVisible(False)
        self.box_m3.setVisible(False)
        self.box_m4.setVisible(False)

        lay.addLayout(self.panel_wrap)
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
        self.cmb_m4_fallback.currentIndexChanged.connect(self._on_m4_fallback_changed)
        self.ck_m4_support.toggled.connect(self._on_m4_support_toggled)
        self.btn_m4_floor_all.clicked.connect(self._m4_select_all_floors)
        self.btn_m4_floor_none.clicked.connect(self._m4_clear_all_floors)
        self.btn_m4_floor_base.clicked.connect(self._m4_select_basement_only)
        self.btn_m4_floor_std.clicked.connect(self._m4_select_standard_only)
        self.rb_m4_shared.toggled.connect(self._on_shared_mode_changed)
        self.rb_m4_byfloor.toggled.connect(self._on_shared_mode_changed)
        self.lv_m4_floors.itemSelectionChanged.connect(self._on_floor_selected_change)
        self.btn_m4_addrow.clicked.connect(lambda: self._plan_table_add_row(self.tbl_m4_plan))
        self.btn_m4_delrow.clicked.connect(lambda: self._plan_table_remove_selected(self.tbl_m4_plan))
        self.btn_m4_even.clicked.connect(lambda: self._on_even_clicked(self.tbl_m4_plan))
        self.btn_m4_copy2all.clicked.connect(self._on_copy_to_all)
        self.btn_m4_def_add.clicked.connect(lambda: self._plan_table_add_row(self.tbl_m4_default))
        self.btn_m4_def_del.clicked.connect(lambda: self._plan_table_remove_selected(self.tbl_m4_default))
        self.btn_m4_def_even.clicked.connect(lambda: self._on_even_clicked(self.tbl_m4_default))

        self._apply_detection_to_mode1_ui()
        self._on_m4_support_toggled(self.ck_m4_support.isChecked())
        self._on_m4_fallback_changed(self.cmb_m4_fallback.currentIndex())
        self._on_shared_mode_changed()

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
            return
        if self._grouped_cache is not None and self._floors_by_cat:
            return
        try:
            grouped, _cats = prepare_from_word(self.doc_path)
        except Exception:
            self._grouped_cache = None
            self._floors_by_cat = {}
            return
        self._grouped_cache = grouped
        floors: dict[str, set[str]] = {}
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
            if labels:
                floors[cat] = labels
        self._floors_by_cat = floors

    def _init_plan_table(self, table: QTableWidget):
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["日期", "上限", "不限"])
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        table.verticalHeader().setVisible(False)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.SingleSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setAlternatingRowColors(True)

    def _plan_table_add_row(self, table: QTableWidget, entry: tuple[str, int | None] | None = None):
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
        unlimited_check = QCheckBox()

        if entry is not None and entry[1] is not None:
            limit_spin.setValue(int(entry[1]))
            unlimited_check.setChecked(False)
            limit_spin.setEnabled(True)
        else:
            if entry and entry[1] is not None:
                limit_spin.setValue(int(entry[1]))
            else:
                limit_spin.setValue(0)
            unlimited_check.setChecked(entry is not None and entry[1] is None)
            limit_spin.setEnabled(not unlimited_check.isChecked())

        def _on_toggle_unlimited(on: bool, spin: QSpinBox = limit_spin):
            spin.setEnabled(not on)

        unlimited_check.toggled.connect(_on_toggle_unlimited)

        table.setCellWidget(row, 0, date_edit)
        table.setCellWidget(row, 1, limit_spin)
        table.setCellWidget(row, 2, unlimited_check)
        table.setRowHeight(row, 32)

    def _plan_table_set_entries(self, table: QTableWidget, entries: list[tuple[str, int | None]]):
        table.setRowCount(0)
        for entry in entries or []:
            self._plan_table_add_row(table, entry)

    def _plan_table_collect(self, table: QTableWidget) -> list[tuple[str, int | None]]:
        results: list[tuple[str, int | None]] = []
        for row in range(table.rowCount()):
            date_edit = table.cellWidget(row, 0)
            limit_widget = table.cellWidget(row, 1)
            unlimited_widget = table.cellWidget(row, 2)
            if not isinstance(date_edit, QDateEdit) or not isinstance(limit_widget, QSpinBox) or not isinstance(unlimited_widget, QCheckBox):
                continue
            date_str = date_edit.date().toString("yyyy-M-d")
            if unlimited_widget.isChecked():
                results.append((date_str, None))
            else:
                results.append((date_str, int(limit_widget.value())))
        return results

    def _plan_table_remove_selected(self, table: QTableWidget):
        selected_rows = sorted({idx.row() for idx in table.selectedIndexes()}, reverse=True)
        if not selected_rows and table.rowCount() > 0:
            selected_rows = [table.rowCount() - 1]
        for row in selected_rows:
            table.removeRow(row)

    def _on_even_clicked(self, table: QTableWidget):
        limited_rows: list[tuple[int, QSpinBox]] = []
        total = 0
        for row in range(table.rowCount()):
            limit_widget = table.cellWidget(row, 1)
            unlimited_widget = table.cellWidget(row, 2)
            if not isinstance(limit_widget, QSpinBox) or not isinstance(unlimited_widget, QCheckBox):
                continue
            if unlimited_widget.isChecked():
                continue
            limited_rows.append((row, limit_widget))
            total += int(limit_widget.value())
        if not limited_rows or total <= 0:
            return
        base = total // len(limited_rows)
        extra = total % len(limited_rows)
        for idx, (_row, spin) in enumerate(limited_rows):
            spin.setValue(base + (1 if idx < extra else 0))

    def _save_current_floor_entries(self):
        if not hasattr(self, "tbl_m4_plan"):
            return
        if self.m4_shared_mode:
            self.m4_entries_shared = self._plan_table_collect(self.tbl_m4_plan)
        else:
            if self._cur_m4_floor:
                self.m4_entries_by_floor[self._cur_m4_floor] = self._plan_table_collect(self.tbl_m4_plan)

    def _load_entries_for_current_floor(self):
        if not hasattr(self, "tbl_m4_plan"):
            return
        if self.m4_shared_mode:
            self._plan_table_set_entries(self.tbl_m4_plan, self.m4_entries_shared)
        else:
            if self._cur_m4_floor:
                entries = self.m4_entries_by_floor.get(self._cur_m4_floor, [])
                self._plan_table_set_entries(self.tbl_m4_plan, entries)
            else:
                self._plan_table_set_entries(self.tbl_m4_plan, [])

    def _current_selected_floor(self) -> str | None:
        if not hasattr(self, "lv_m4_floors"):
            return None
        item = self.lv_m4_floors.currentItem()
        return item.text() if item else None

    def _refresh_m4_floor_list(self):
        if not hasattr(self, "lv_m4_floors"):
            return
        sorter = _floor_sort_key_by_label or (lambda x: x)
        available = sorted(self.m4_selected_floors, key=sorter)
        previous = self._cur_m4_floor if self._cur_m4_floor in self.m4_selected_floors else None
        if not previous and available:
            previous = available[0]
        self.lv_m4_floors.blockSignals(True)
        self.lv_m4_floors.clear()
        for floor in available:
            item = QListWidgetItem(floor)
            self.lv_m4_floors.addItem(item)
            if floor == previous:
                item.setSelected(True)
                self.lv_m4_floors.setCurrentItem(item)
        self.lv_m4_floors.blockSignals(False)
        if self.m4_shared_mode:
            self._cur_m4_floor = None
            self._load_entries_for_current_floor()
        else:
            self._cur_m4_floor = previous
            self._load_entries_for_current_floor()
        if hasattr(self, "btn_m4_copy2all"):
            self.btn_m4_copy2all.setEnabled(bool(available))

    def _m4_set_selected_floors(self, floors: set[str]):
        self.m4_selected_floors = set(floors)
        for name, btn in self.m4_floor_buttons.items():
            btn.blockSignals(True)
            btn.setChecked(name in self.m4_selected_floors)
            btn.blockSignals(False)
        self._refresh_m4_floor_list()

    def _m4_on_floor_chip_toggled(self, name: str, checked: bool):
        self._save_current_floor_entries()
        if checked:
            self.m4_selected_floors.add(name)
        else:
            self.m4_selected_floors.discard(name)
            if self._cur_m4_floor == name:
                self._cur_m4_floor = None
        self._refresh_m4_floor_list()

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

    def _reset_m4_plan_state(self):
        self.m4_selected_floors = set()
        self.m4_entries_shared = []
        self.m4_entries_by_floor = {}
        self.m4_all_floors = []
        self._cur_m4_floor = None
        self.m4_floor_buttons = {}
        self.m4_shared_mode = True
        if hasattr(self, "tbl_m4_plan"):
            self._plan_table_set_entries(self.tbl_m4_plan, [])
        if hasattr(self, "tbl_m4_default"):
            self._plan_table_set_entries(self.tbl_m4_default, [])
        if hasattr(self, "lv_m4_floors"):
            self.lv_m4_floors.clear()
        if hasattr(self, "rb_m4_shared"):
            self.rb_m4_shared.setChecked(True)
        if hasattr(self, "btn_m4_copy2all"):
            self.btn_m4_copy2all.hide()

    def _m4_select_all_floors(self):
        self._save_current_floor_entries()
        self._m4_set_selected_floors(set(self.m4_all_floors))

    def _m4_clear_all_floors(self):
        self._save_current_floor_entries()
        self._m4_set_selected_floors(set())

    def _m4_select_basement_only(self):
        self._save_current_floor_entries()
        selected = {f for f in self.m4_all_floors if f.upper().startswith("B")}
        self._m4_set_selected_floors(selected)

    def _m4_select_standard_only(self):
        self._save_current_floor_entries()
        selected = {f for f in self.m4_all_floors if re.match(r"\d+F", f.upper())}
        self._m4_set_selected_floors(selected)

    def _on_shared_mode_changed(self):
        if not hasattr(self, "rb_m4_shared"):
            return
        self._save_current_floor_entries()
        self.m4_shared_mode = self.rb_m4_shared.isChecked()
        self.lv_m4_floors.setDisabled(self.m4_shared_mode)
        self.btn_m4_copy2all.setVisible(not self.m4_shared_mode)
        if self.m4_shared_mode:
            self._cur_m4_floor = None
        else:
            current = self._current_selected_floor()
            if current:
                self._cur_m4_floor = current
            elif self.m4_selected_floors:
                sorter = _floor_sort_key_by_label or (lambda x: x)
                ordered = sorted(self.m4_selected_floors, key=sorter)
                self._cur_m4_floor = ordered[0] if ordered else None
                if self._cur_m4_floor:
                    items = self.lv_m4_floors.findItems(self._cur_m4_floor, Qt.MatchExactly)
                    if items:
                        self.lv_m4_floors.setCurrentItem(items[0])
        self._load_entries_for_current_floor()

    def _on_floor_selected_change(self):
        if self.m4_shared_mode:
            return
        self._save_current_floor_entries()
        self._cur_m4_floor = self._current_selected_floor()
        self._load_entries_for_current_floor()

    def _on_copy_to_all(self):
        if self.m4_shared_mode:
            return
        entries = self._plan_table_collect(self.tbl_m4_plan)
        for floor in self.m4_selected_floors:
            self.m4_entries_by_floor[floor] = list(entries)

    def _collect_m4_plan_from_ui(self) -> dict:
        if not hasattr(self, "tbl_m4_plan"):
            return {}
        self._save_current_floor_entries()

        categories: list[str] = []
        if self.sw_m4_cat_gz.isVisible() and self.sw_m4_cat_gz.isEnabled() and self.sw_m4_cat_gz.isChecked():
            categories.append("钢柱")
        if self.sw_m4_cat_gl.isVisible() and self.sw_m4_cat_gl.isEnabled() and self.sw_m4_cat_gl.isChecked():
            categories.append("钢梁")
        if self.sw_m4_cat_sup.isVisible() and self.sw_m4_cat_sup.isEnabled() and self.sw_m4_cat_sup.isChecked():
            categories.append("支撑")
        if self.sw_m4_cat_net.isVisible() and self.sw_m4_cat_net.isEnabled() and self.sw_m4_cat_net.isChecked():
            categories.append("网架")

        include_support = (
            self.ck_m4_support.isVisible()
            and self.ck_m4_support.isEnabled()
            and self.ck_m4_support.isChecked()
        )

        if not include_support and "支撑" in categories:
            categories.remove("支撑")

        if not categories or not self.m4_selected_floors:
            return {}

        if self.m4_shared_mode:
            self.m4_entries_shared = self._plan_table_collect(self.tbl_m4_plan)
            by_floor = {
                floor: list(self.m4_entries_shared)
                for floor in self.m4_selected_floors
                if self.m4_entries_shared
            }
        else:
            if self._cur_m4_floor:
                self.m4_entries_by_floor[self._cur_m4_floor] = self._plan_table_collect(self.tbl_m4_plan)
            by_floor = {
                floor: list(self.m4_entries_by_floor.get(floor, []))
                for floor in self.m4_selected_floors
                if self.m4_entries_by_floor.get(floor)
            }

        plan: dict[str, dict[str, list[tuple[str, int | None]]]] = {}
        for cat in categories:
            if by_floor:
                plan[cat] = {floor: list(entries) for floor, entries in by_floor.items()}

        return plan

    def _apply_detection_to_mode4_ui(self):
        if not hasattr(self, "m4_floor_chips"):
            return
        gz_ok = self.present.get("钢柱", False)
        gl_ok = self.present.get("钢梁", False)
        sup_ok = self.present.get("支撑", False)
        net_ok = self.present.get("网架", False)

        for ok, widget in (
            (gz_ok, self.sw_m4_cat_gz),
            (gl_ok, self.sw_m4_cat_gl),
            (sup_ok, self.sw_m4_cat_sup),
            (net_ok, self.sw_m4_cat_net),
        ):
            widget.setVisible(ok)
            widget.setEnabled(ok)
            if ok and not widget.isChecked():
                widget.setChecked(True)
            if not ok:
                widget.setChecked(False)

        self.ck_m4_support.setVisible(sup_ok)
        if not sup_ok:
            self.ck_m4_support.setChecked(False)
        self.lb_m4_sup_strategy.setVisible(sup_ok)
        self.cmb_m4_sup_strategy.setVisible(sup_ok)
        sup_enabled = sup_ok and self.ck_m4_support.isChecked()
        self.lb_m4_sup_strategy.setEnabled(sup_enabled)
        self.cmb_m4_sup_strategy.setEnabled(sup_enabled)
        self.sw_m4_cat_sup.setEnabled(sup_enabled)
        if not sup_enabled:
            self.sw_m4_cat_sup.setChecked(False)

        self.lb_m4_net_strategy.setVisible(net_ok)
        self.cmb_m4_net_strategy.setVisible(net_ok)
        self.sw_m4_cat_net.setVisible(net_ok)
        if not net_ok:
            self.sw_m4_cat_net.setChecked(False)

        active_cats = gz_ok or gl_ok or sup_ok or net_ok
        self.box_m4.setDisabled(not active_cats)

        sorter = _floor_sort_key_by_label or (lambda x: x)
        floors_set: set[str] = set()
        for cat in ("钢柱", "钢梁", "支撑", "网架"):
            floors_set.update(self._floors_by_cat.get(cat, set()))
        floors = sorted(floors_set, key=sorter)
        self.m4_all_floors = floors
        self._rebuild_m4_floor_chips(floors)
        if self.m4_selected_floors:
            selected = self.m4_selected_floors & set(floors)
            if not selected and floors:
                selected = set(floors)
        else:
            selected = set(floors)
        self._m4_set_selected_floors(selected)

        has_floors = bool(floors)
        for btn in (
            self.btn_m4_floor_all,
            self.btn_m4_floor_none,
            self.btn_m4_floor_base,
            self.btn_m4_floor_std,
        ):
            btn.setEnabled(has_floors)

    def _update_m4_floor_hint(self):
        if not hasattr(self, "lb_m4_floors"):
            return
        if not self._floors_by_cat:
            self.lb_m4_floors.setText("（楼层信息将在读取后显示）")
            return
        parts = []
        sorter = _floor_sort_key_by_label or (lambda x: x)
        for cat in ("钢柱", "钢梁", "支撑", "网架"):
            floors = sorted(self._floors_by_cat.get(cat, []), key=sorter)
            if floors:
                parts.append(f"{cat}：{' '.join(floors)}")
        self.lb_m4_floors.setText(" | ".join(parts))

    def _on_m4_support_toggled(self, checked: bool):
        if not hasattr(self, "sw_m4_cat_sup"):
            return
        sup_ok = self.present.get("支撑", False)
        enabled = checked and sup_ok
        self.lb_m4_sup_strategy.setEnabled(enabled)
        self.cmb_m4_sup_strategy.setEnabled(enabled)
        self.sw_m4_cat_sup.setEnabled(enabled)
        if not enabled:
            self.sw_m4_cat_sup.setChecked(False)
        elif enabled and not self.sw_m4_cat_sup.isChecked():
            self.sw_m4_cat_sup.setChecked(True)

    def _on_m4_fallback_changed(self, idx: int):
        if not hasattr(self, "w_m4_default"):
            return
        show = idx == 1
        self.w_m4_default.setVisible(show)
        if show and hasattr(self, "tbl_m4_default") and self.tbl_m4_default.rowCount() == 0:
            self._plan_table_add_row(self.tbl_m4_default)


    # ====== 返回 Step1 重选文件 ======
    def _go_back_to_select(self):
        self.stack.setCurrentIndex(0)
        self.status.setText("准备就绪")

    # ====== Mode 1：日期分桶 ======
    def _collect_mode1_buckets(self) -> list[dict]:
        buckets: list[dict] = []
        if not self._m1_day_forms:
            return buckets

        for form in self._m1_day_forms:
            date_str = form["date"].text().strip()
            parts: dict[str, object] = {}

            if "钢柱" in form:
                parts["钢柱"] = self._to_rule(form["钢柱"].text())
            if "钢梁" in form:
                parts["钢梁"] = self._to_rule(form["钢梁"].text())
            if "支撑" in form:
                parts["支撑"] = self._to_rule(form["支撑"].text())
            if "网架_xx" in form:
                parts["网架"] = {
                    "XX": self._to_rule(form["网架_xx"].text()),
                    "FG": self._to_rule(form["网架_fg"].text()),
                    "SX": self._to_rule(form["网架_sx"].text()),
                    "GEN": self._to_rule(form["网架_gen"].text()),
                }

            buckets.append({"date": date_str, "rules": parts})

        return buckets

    def _on_run_mode1(self):
        if not export_mode1_noninteractive:
            QMessageBox.critical(self, "提示", "后端暂不支持 Mode 1 生成接口。")
            return
        if not self.doc_path:
            QMessageBox.warning(self, "提示", "请先选择 Word 源文件。")
            return

        buckets = self._collect_mode1_buckets()
        if not buckets:
            QMessageBox.warning(self, "提示", "请至少填写一天数据。")
            return

        def _has_content(bucket: dict) -> bool:
            if bucket.get("date"):
                return True
            rules = bucket.get("rules") or bucket.get("parts") or {}
            for key, value in rules.items():
                if key == "网架":
                    if any(part.get("enabled") for part in value.values()):
                        return True
                elif value.get("enabled"):
                    return True
            return False

        if not any(_has_content(b) for b in buckets):
            QMessageBox.warning(self, "提示", "请至少填写一天数据。")
            return

        support_strategy = "floor" if self.cmb_m1_sup.isVisible() and self.cmb_m1_sup.currentIndex() == 1 else "number"
        net_strategy = "floor" if self.cmb_m1_net.isVisible() and self.cmb_m1_net.currentIndex() == 1 else "number"
        later_priority = self.ck_m1_later.isChecked()
        auto_merge_rest = self.ck_m1_merge.isChecked()

        self.status.setText("⏳ 正在生成（Mode 1 / 日期分桶）…")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            xlsx, word = export_mode1_noninteractive(
                src_docx=str(self.doc_path),
                buckets=buckets,
                support_strategy=support_strategy,
                net_strategy=net_strategy,
                later_priority=later_priority,
                auto_merge_rest=auto_merge_rest,
                meta={},
            )
            QMessageBox.information(self, "完成", f"✅ 生成完成！\nExcel：{xlsx}\n汇总Word：{word}")
            self.status.setText("✅ 日期分桶完成")
        except Exception as e:
            QMessageBox.critical(self, "失败", f"生成失败：\n{e}")
            self.status.setText("❌ 生成失败")
        finally:
            QApplication.restoreOverrideCursor()



    def _on_run_mode4(self):
        if not export_mode4_noninteractive:
            QMessageBox.critical(self, "提示", "后端暂不支持 Mode 4 生成接口。")
            return
        if not self.doc_path:
            QMessageBox.warning(self, "提示", "请先选择 Word 源文件。")
            return

        plan = self._collect_m4_plan_from_ui()
        if not plan:
            QMessageBox.warning(self, "提示", "请至少为一个类别填写计划。")
            return

        sup_strategy = "number"
        if self.lb_m4_sup_strategy.isVisible() and self.cmb_m4_sup_strategy.currentIndex() == 1:
            sup_strategy = "floor"
        net_strategy = "number"
        if self.lb_m4_net_strategy.isVisible() and self.cmb_m4_net_strategy.currentIndex() == 1:
            net_strategy = "floor"

        fb_map = {0: "append_last", 1: "default", 2: "error"}
        fallback = fb_map.get(self.cmb_m4_fallback.currentIndex(), "append_last")

        default_entries = None
        if fallback == "default":
            default_entries = self._plan_table_collect(self.tbl_m4_default)
            if not default_entries:
                QMessageBox.warning(self, "提示", "请填写默认计划的日期与上限。")
                return

        include_support = (
                self.ck_m4_support.isVisible() and self.ck_m4_support.isEnabled() and self.ck_m4_support.isChecked()
        )

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
            )
            QMessageBox.information(self, "完成", f"✅ 生成完成！\nExcel：{xlsx}\n汇总Word：{word}")
            self.status.setText("✅ Mode 4 完成")
        except Exception as e:
            QMessageBox.critical(self, "失败", f"生成失败：\n{e}")
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
