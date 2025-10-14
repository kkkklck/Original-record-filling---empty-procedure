# graphic.py — 双页面向导式 GUI（PySide6）
# Step 1: 仅路径 -> 自动静默检索 -> 进入 Step 2
# Step 2: 显示“识别结果（带数量）”、选择 Mode，并只展开对应表单
# 改动要点：
#   - 新增：类别规范化映射，兼容“斜撑/桁架/Truss”等写法
#   - 新增：顶部“识别结果”标签条（有什么就展示什么）
#   - 改进：Mode2 的“可包含”行带数量，复选框采用蓝色勾选样式，更显眼

from __future__ import annotations
import os, sys, importlib.util
from pathlib import Path
from dataclasses import dataclass

from PySide6.QtCore import Qt, QSize, QThread, Signal
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QGroupBox, QFileDialog, QRadioButton, QButtonGroup,
    QCheckBox, QMessageBox, QSpacerItem, QSizePolicy, QStackedWidget, QFrame
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
        self.rb_m1 = QRadioButton("Mode 1"); self.rb_m2 = QRadioButton("Mode 2"); self.rb_m3 = QRadioButton("Mode 3"); self.rb_m4 = QRadioButton("Mode 4")
        self.rb_m2.setChecked(True); self.rb_m1.setEnabled(False); self.rb_m4.setEnabled(False)
        self.grp_mode = QButtonGroup(self)
        for i, rb in enumerate([self.rb_m1, self.rb_m2, self.rb_m3, self.rb_m4], start=1):
            self.grp_mode.addButton(rb, i); lm.addWidget(rb)
        lm.addStretch(1)
        lay.addWidget(mode_box)

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

        row_sub = QHBoxLayout()
        row_sub.addWidget(QLabel("子模式"))
        self.rb_both = QRadioButton("两者"); self.rb_gz = QRadioButton("仅钢柱"); self.rb_gl = QRadioButton("仅钢梁")
        self.grp_sub = QButtonGroup(self)
        for i, b in enumerate([self.rb_both, self.rb_gz, self.rb_gl], start=1):
            self.grp_sub.addButton(b, i); row_sub.addWidget(b)
        lm2.addLayout(row_sub)

        row_bp1 = QHBoxLayout()
        self.lb_bp_gz = QLabel("钢柱断点")
        self.ed_bp_gz = QLineEdit(); self.ed_bp_gz.setPlaceholderText("例：3 6 10（空=不分段）")
        row_bp1.addWidget(self.lb_bp_gz); row_bp1.addWidget(self.ed_bp_gz, 1)

        row_bp2 = QHBoxLayout()
        self.lb_bp_gl = QLabel("钢梁断点")
        self.ed_bp_gl = QLineEdit(); self.ed_bp_gl.setPlaceholderText("例：3 6 10（空=不分段）")
        row_bp2.addWidget(self.lb_bp_gl); row_bp2.addWidget(self.ed_bp_gl, 1)

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
        self.ck_net     = QCheckBox("网架")
        self.ck_other   = QCheckBox("其他")
        row_inc.addWidget(self.ck_support); row_inc.addWidget(self.ck_net); row_inc.addWidget(self.ck_other)
        row_inc.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))

        row_go = QHBoxLayout()
        row_go.addSpacerItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.btn_run_m2 = QPushButton("生成（楼层断点）")
        self.btn_run_m2.setFixedSize(QSize(160, 36))
        row_go.addWidget(self.btn_run_m2)

        for r in (row_bp1, row_bp2, row_dt, row_inc, row_go):
            lm2.addLayout(r)

        # 容器：只显示当前模式对应的表单
        self.panel_wrap = QVBoxLayout()
        self.panel_wrap.addWidget(self.box_m2)  # 默认显示 M2
        self.panel_wrap.addWidget(self.box_m3)
        self.box_m3.setVisible(False)

        lay.addLayout(self.panel_wrap)
        lay.addStretch(1)

        lay.addWidget(hline())
        self.status = QLabel("准备就绪"); self.status.setStyleSheet("color:#555;")
        lay.addWidget(self.status)

        # 事件
        self.btn_back.clicked.connect(self._go_back_to_select)
        self.grp_mode.idToggled.connect(self._on_mode_switched)
        self.btn_run_m2.clicked.connect(self._on_run_mode2)
        self.btn_run_m3.clicked.connect(self._on_run_mode3)

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
        self._apply_detection_to_mode2_ui()
        self._refresh_found_bar()
        self.lb_file_short.setText(f"文件：{self.doc_path.name}")
        self.status.setText("✅ 已分析完成，可选择模式继续")
        self.stack.setCurrentIndex(1)

    # ====== Step2：模式切换 & 表单显隐 ======
    def _on_mode_switched(self, _id: int, checked: bool):
        if not checked:
            return
        m2 = (self.grp_mode.checkedButton() is self.rb_m2)
        self.box_m2.setVisible(m2)
        self.box_m3.setVisible(not m2)

    # 顶部“识别结果”标签条
    def _refresh_found_bar(self):
        parts = []
        for k in CANON_KEYS:
            if self.present.get(k, False):
                num = self.counts.get(k, 0)
                parts.append(f"{k}（{num}）" if num else f"{k}")
        self.lb_found.setText("、".join(parts) if parts else "未识别到有效构件")

    def _apply_detection_to_mode2_ui(self):
        # 子模式（仅针对钢柱/钢梁）
        gz_ok = self.present.get("钢柱", False)
        gl_ok = self.present.get("钢梁", False)
        both_ok = gz_ok and gl_ok

        self.rb_both.setVisible(both_ok)
        self.rb_gz.setVisible(gz_ok)
        self.rb_gl.setVisible(gl_ok)

        if both_ok:
            self.rb_both.setChecked(True)
        elif gz_ok:
            self.rb_gz.setChecked(True)
        elif gl_ok:
            self.rb_gl.setChecked(True)
        else:
            self.box_m2.setDisabled(True)
            self.status.setText("⚠ 未识别到钢柱/钢梁，Mode 2 可能不可用。")

        # 断点输入显隐
        self.lb_bp_gz.setVisible(gz_ok); self.ed_bp_gz.setVisible(gz_ok)
        self.lb_bp_gl.setVisible(gl_ok); self.ed_bp_gl.setVisible(gl_ok)

        # 其他构件（带数量）
        def set_ck(ck: QCheckBox, key: str):
            vis = self.present.get(key, False)
            ck.setVisible(vis); ck.setEnabled(vis); ck.setChecked(vis)
            num = self.counts.get(key, 0)
            base = key if num == 0 else f"{key}（{num}）"
            ck.setText(base)

        set_ck(self.ck_support, "支撑")
        set_ck(self.ck_net, "网架")
        set_ck(self.ck_other, "其他")

    # ====== 返回 Step1 重选文件 ======
    def _go_back_to_select(self):
        self.stack.setCurrentIndex(0)
        self.status.setText("准备就绪")

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
            QMessageBox.warning(self, "提示", "请先选择 Word 源文件。"); return

        # 子模式（钢柱/钢梁）
        if self.rb_both.isVisible() and self.rb_both.isChecked():
            choose = "both"
        elif self.rb_gz.isVisible() and self.rb_gz.isChecked():
            choose = "gz"
        else:
            choose = "gl"

        bp_gz = (self.ed_bp_gz.text() or "").strip() if self.ed_bp_gz.isVisible() else ""
        bp_gl = (self.ed_bp_gl.text() or "").strip() if self.ed_bp_gl.isVisible() else ""
        dt_first  = (self.ed_dt_first.text() or "").strip()
        dt_second = (self.ed_dt_second.text() or "").strip()

        inc_support = self.ck_support.isVisible() and self.ck_support.isChecked()
        inc_net     = self.ck_net.isVisible() and self.ck_net.isChecked()
        inc_other   = self.ck_other.isVisible() and self.ck_other.isChecked()

        meta = {}

        self.status.setText("⏳ 正在生成（楼层断点）…")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            out = export_mode2_noninteractive(
                src=str(self.doc_path),
                meta=meta,
                choose=choose,
                breaks_gz=bp_gz,
                breaks_gl=bp_gl,
                date_first=dt_first,
                date_second=dt_second,
                include_support=inc_support,
                include_net=inc_net,
                include_other=inc_other,
                support_strategy="number",
                net_strategy="number",
            )
            xlsx = out.get("excel"); word = out.get("word")
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
