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

from PySide6.QtCore import Qt, QSize, QThread, Signal
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QGroupBox, QFileDialog, QRadioButton, QButtonGroup,
    QCheckBox, QMessageBox, QSpacerItem, QSizePolicy, QStackedWidget, QFrame,
    QComboBox, QScrollArea, QSpinBox,
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

        self.doc_path: Path | None = None
        self.present = {k: False for k in CANON_KEYS}
        self.counts  = {k: 0 for k in CANON_KEYS}
        self._m1_day_forms: list[dict] = []

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

        self.lb_status1 = QLabel("就绪"); self.lb_status1.setStyleSheet("color:#777;")
        b.addWidget(self.lb_status1)
        lay.addWidget(box)

        tip = QLabel(f"后端模块：{ORF_LOADED_FROM or '未知'}"); tip.setStyleSheet("color:#999;")
        lay.addWidget(tip); lay.addStretch(1)

        self.btn_browse.clicked.connect(self._on_browse_and_probe)
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
        self.rb_m4.setEnabled(False)
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

        # 容器：只显示当前模式对应的表单
        self.panel_wrap = QVBoxLayout()
        self.panel_wrap.addWidget(self.box_m1)
        self.panel_wrap.addWidget(self.box_m2)  # 默认显示 M2
        self.panel_wrap.addWidget(self.box_m3)
        self.box_m1.setVisible(False)
        self.box_m3.setVisible(False)

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

        self._apply_detection_to_mode1_ui()

        return w

    # ====== 样式（增加 QCheckBox 的蓝色勾） ======
    def _apply_styles(self):
        self.setStyleSheet("""
            QWidget { background:#ffffff; color:#333; font-size:14px; }
            QGroupBox {
                border:1px solid #e7e7e7; border-radius:12px; margin-top:12px; padding:12px;
                font-weight:600;
            }
            QGroupBox::title { subcontrol-origin: margin; left:12px; padding:0 6px; background:transparent; }
            QLineEdit {
                height:34px; border:1px solid #d9d9d9; border-radius:8px; padding:4px 10px; background:#fafafa;
            }
            QPushButton {
                height:34px; border:1px solid #d9d9d9; border-radius:10px; background:#f6f6f6; padding:0 12px;
            }
            QPushButton:hover { background:#efefef; }
            /* —— 小蓝点单选 —— */
            QRadioButton { spacing:8px; }
            QRadioButton::indicator {
                width:14px; height:14px; border-radius:7px;
                border:2px solid #9aa0a6; background:#fff; margin-right:6px;
            }
            QRadioButton::indicator:hover { border-color:#6f8ccf; }
            QRadioButton::indicator:checked {
                background:#2d89ef; border:2px solid #2d89ef;
            }
            QRadioButton:checked { color:#2d89ef; font-weight:700; }
            /* —— 复选框明显可见 —— */
            QCheckBox::indicator {
                width:16px; height:16px; border-radius:4px;
                border:2px solid #9aa0a6; background:#fff; margin-right:6px;
            }
            QCheckBox::indicator:hover { border-color:#6f8ccf; }
            QCheckBox::indicator:checked {
                image: none; background:#2d89ef; border:2px solid #2d89ef;
            }
        """)

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
        text = (value or "").strip()
        if not text:
            return {"enabled": False, "ranges": []}
        if text in {"*", "＊", "全部", "所有"}:
            return {"enabled": True, "ranges": []}
        return {"enabled": True, "ranges": text}

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

            buckets.append({"date": date_str, "parts": parts})

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
            for key, value in bucket.get("parts", {}).items():
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
