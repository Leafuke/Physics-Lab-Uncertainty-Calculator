from __future__ import annotations

import re
from pathlib import Path

from PySide6.QtCore import QEvent, Qt, QTimer, Signal
from PySide6.QtGui import QAction, QCloseEvent, QIcon
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStyledItemDelegate,
    QStyleFactory,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolBar,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from . import APP_DISPLAY_NAME, APP_VERSION
from .calculations import CalculationResult, calculate_project, format_number, normalize_decimal_places, rounded_measurement
from .excel_io import export_data_to_excel, export_project_to_excel, import_measurements_from_excel
from .models import BSource, BSourceType, ProjectData, b_source_display_name, default_divisor_for
from .persistence import (
    clear_recent_files,
    last_project_path,
    load_autosave,
    load_project_file,
    push_recent_file,
    recent_files,
    save_project_file,
    set_last_project_path,
    settings,
    write_autosave,
)


APP_STYLE_SHEET = """
QMainWindow {
    background: #f3f6fb;
}

QToolBar {
    spacing: 6px;
    padding: 5px 8px;
    background: #ffffff;
    border: none;
    border-bottom: 1px solid #d8e0ea;
}

QToolBar::separator {
    width: 1px;
    margin: 4px 6px;
    background: #d8e0ea;
}

QToolButton {
    padding: 6px 10px;
    border-radius: 10px;
}

QToolButton:hover {
    background: #eef4ff;
}

QMenu {
    background: #ffffff;
    border: 1px solid #d8e0ea;
    border-radius: 12px;
    padding: 6px;
}

QMenu::item {
    padding: 8px 14px;
    border-radius: 8px;
}

QMenu::item:selected {
    background: #eef4ff;
    color: #12345a;
}

QGroupBox {
    font-size: 14px;
    font-weight: 600;
    color: #243b53;
    background: #ffffff;
    border: 1px solid #d8e0ea;
    border-radius: 16px;
    margin-top: 12px;
    padding: 16px 14px 14px 14px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 14px;
    padding: 0 4px;
}

QLabel {
    color: #243b53;
}

QLineEdit,
QPlainTextEdit,
QTextEdit,
QTableWidget,
QListWidget,
QComboBox,
QDoubleSpinBox {
    background: #fbfcfe;
    border: 1px solid #c7d2e2;
    border-radius: 12px;
    padding: 6px 8px;
    selection-background-color: #cfe1ff;
}

QLineEdit:focus,
QPlainTextEdit:focus,
QTextEdit:focus,
QTableWidget:focus,
QListWidget:focus,
QComboBox:focus,
QDoubleSpinBox:focus {
    border: 1px solid #4c8bf5;
}

QPushButton {
    background: #e8eef9;
    color: #16324f;
    border: 1px solid #c4d4ec;
    border-radius: 12px;
    padding: 8px 12px;
}

QPushButton:hover {
    background: #dce8fb;
}

QPushButton#primaryButton {
    background: #1f5fbf;
    color: #ffffff;
    border: none;
}

QPushButton#primaryButton:hover {
    background: #184d9a;
}

QHeaderView::section {
    background: #eff4fb;
    color: #243b53;
    border: none;
    border-bottom: 1px solid #d8e0ea;
    padding: 8px;
    font-weight: 600;
}

QFrame#metricCard {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1, stop: 0 #ffffff, stop: 1 #edf5ff);
    border: 1px solid #d1ddef;
    border-radius: 18px;
}

QLabel#metricTitle {
    color: #5b7083;
    font-size: 12px;
    font-weight: 600;
}

QLabel#metricValue {
    color: #12345a;
    font-size: 24px;
    font-weight: 700;
}

QLabel#resultHeadline {
    color: #12345a;
    font-size: 20px;
    font-weight: 700;
}

QLabel#resultSubline {
    color: #486581;
    font-size: 13px;
}

QLabel#warningLabel {
    background: #fff4e5;
    color: #9c5b00;
    border: 1px solid #ffd59c;
    border-radius: 12px;
    padding: 10px 12px;
}

QLabel#infoPill {
    color: #486581;
    background: #edf3fb;
    border-radius: 12px;
    padding: 8px 10px;
}
"""


class MeasurementItemDelegate(QStyledItemDelegate):
    enterPressed = Signal(int)

    def createEditor(self, parent, option, index):
        editor = super().createEditor(parent, option, index)
        if editor is not None:
            editor.setProperty("measurementRow", index.row())
            editor.installEventFilter(self)
        return editor

    def eventFilter(self, editor, event):
        if event.type() == QEvent.Type.KeyPress and event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            row_value = editor.property("measurementRow")
            if row_value is not None:
                QTimer.singleShot(0, lambda current_row=int(row_value): self.enterPressed.emit(current_row))
        return super().eventFilter(editor, event)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.project = ProjectData()
        self.project.ensure_defaults()
        self.latest_result = CalculationResult()
        self.current_project_path: str | None = None
        self.is_dirty = False
        self._loading = False

        self.autosave_timer = QTimer(self)
        self.autosave_timer.setSingleShot(True)
        self.autosave_timer.timeout.connect(self._write_autosave_snapshot)

        self.setWindowTitle(APP_DISPLAY_NAME)
        self.setMinimumSize(1420, 880)

        self._build_actions()
        self._build_ui()
        self._restore_window_state()
        self._load_initial_project()

    def _build_actions(self) -> None:
        self.new_action = QAction("新建项目", self)
        self.new_action.setShortcut("Ctrl+N")
        self.new_action.triggered.connect(self.new_project)

        self.open_action = QAction("打开项目", self)
        self.open_action.setShortcut("Ctrl+O")
        self.open_action.triggered.connect(self.open_project_dialog)

        self.save_action = QAction("保存项目", self)
        self.save_action.setShortcut("Ctrl+S")
        self.save_action.triggered.connect(self.save_project)

        self.save_as_action = QAction("另存为", self)
        self.save_as_action.setShortcut("Ctrl+Shift+S")
        self.save_as_action.triggered.connect(self.save_project_as)

        self.import_excel_action = QAction("导入 Excel", self)
        self.import_excel_action.triggered.connect(self.import_excel_measurements)

        self.export_data_excel_action = QAction("导出数据 Excel", self)
        self.export_data_excel_action.triggered.connect(self.export_data_excel)

        self.export_excel_action = QAction("导出结果 Excel", self)
        self.export_excel_action.triggered.connect(self.export_excel)

        self.export_txt_action = QAction("导出 TXT", self)
        self.export_txt_action.triggered.connect(self.export_txt)

        self.export_image_action = QAction("导出结果图片", self)
        self.export_image_action.triggered.connect(self.export_image)

        self.clear_recent_action = QAction("清空最近项目", self)
        self.clear_recent_action.triggered.connect(self.clear_recent_project_history)

    def _build_ui(self) -> None:
        self._build_top_toolbar()

        central_widget = QWidget(self)
        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(10, 10, 10, 10)

        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_splitter.setChildrenCollapsible(False)

        main_splitter.addWidget(self._build_left_panel())
        main_splitter.addWidget(self._build_center_panel())
        main_splitter.addWidget(self._build_right_panel())
        main_splitter.setStretchFactor(0, 4)
        main_splitter.setStretchFactor(1, 5)
        main_splitter.setStretchFactor(2, 6)

        central_layout.addWidget(main_splitter)
        self.setCentralWidget(central_widget)
        self.setStatusBar(self.statusBar())
        self.statusBar().showMessage("准备就绪")

    def _build_top_toolbar(self) -> None:
        toolbar = QToolBar("操作栏", self)
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        toolbar.addAction(self.new_action)
        toolbar.addAction(self.open_action)
        toolbar.addAction(self.save_action)
        toolbar.addAction(self.save_as_action)
        toolbar.addSeparator()
        toolbar.addAction(self.import_excel_action)
        toolbar.addAction(self.export_data_excel_action)
        toolbar.addAction(self.export_excel_action)
        toolbar.addAction(self.export_txt_action)
        toolbar.addAction(self.export_image_action)
        toolbar.addSeparator()

        spacer = QWidget(self)
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(spacer)

        self.recent_projects_menu = QMenu(self)
        self.recent_projects_button = QToolButton(self)
        self.recent_projects_button.setText("最近项目")
        self.recent_projects_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.recent_projects_button.setMenu(self.recent_projects_menu)
        toolbar.addWidget(self.recent_projects_button)

        about_button = QToolButton(self)
        about_button.setText("关于本程序")
        about_button.clicked.connect(self.show_about_dialog)
        toolbar.addWidget(about_button)

        self.addToolBar(toolbar)

    def _build_left_panel(self) -> QWidget:
        panel = QWidget(self)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        project_group = QGroupBox("项目与原始数据")
        project_layout = QVBoxLayout(project_group)

        project_form = QFormLayout()
        self.quantity_name_edit = QLineEdit()
        self.quantity_name_edit.setPlaceholderText("例如：单摆周期、钢丝长度、电压")
        self.quantity_name_edit.textChanged.connect(self._handle_user_edit)

        self.unit_edit = QLineEdit()
        self.unit_edit.setPlaceholderText("例如：s、cm、V")
        self.unit_edit.textChanged.connect(self._handle_user_edit)

        project_form.addRow("测量量名称", self.quantity_name_edit)
        project_form.addRow("单位", self.unit_edit)
        project_layout.addLayout(project_form)

        values_button_row = QHBoxLayout()
        add_value_button = QPushButton("新增行")
        add_value_button.clicked.connect(lambda: self.add_measurement_row(focus_and_edit=True))
        remove_value_button = QPushButton("删除选中")
        remove_value_button.clicked.connect(self.remove_selected_measurement_rows)
        clear_value_button = QPushButton("清空数据")
        clear_value_button.clicked.connect(self.clear_measurements)
        values_button_row.addWidget(add_value_button)
        values_button_row.addWidget(remove_value_button)
        values_button_row.addWidget(clear_value_button)
        project_layout.addLayout(values_button_row)

        import_row = QHBoxLayout()
        paste_button = QPushButton("粘贴多项数据")
        paste_button.clicked.connect(self.paste_measurements_from_clipboard)
        import_excel_button = QPushButton("从 Excel 导入")
        import_excel_button.setObjectName("primaryButton")
        import_excel_button.clicked.connect(self.import_excel_measurements)
        import_row.addWidget(paste_button)
        import_row.addWidget(import_excel_button)
        project_layout.addLayout(import_row)

        self.measurement_table = QTableWidget(0, 1)
        self.measurement_table.setHorizontalHeaderLabels(["测量值"])
        self.measurement_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.measurement_table.verticalHeader().setVisible(False)
        self.measurement_table.setAlternatingRowColors(True)
        self.measurement_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.measurement_table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
            | QAbstractItemView.EditTrigger.SelectedClicked
        )
        self.measurement_delegate = MeasurementItemDelegate(self.measurement_table)
        self.measurement_delegate.enterPressed.connect(self._move_to_next_measurement_input)
        self.measurement_table.setItemDelegate(self.measurement_delegate)
        self.measurement_table.itemChanged.connect(self._on_measurement_table_item_changed)
        project_layout.addWidget(self.measurement_table, 1)

        self.import_info_label = QLabel("尚未导入 Excel 数据")
        self.import_info_label.setObjectName("infoPill")
        self.import_info_label.setWordWrap(True)
        project_layout.addWidget(self.import_info_label)

        layout.addWidget(project_group, 1)

        return panel

    def _build_center_panel(self) -> QWidget:
        panel = QWidget(self)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        b_group = QGroupBox("B类评定与计算设置")
        b_layout = QVBoxLayout(b_group)

        b_button_row = QHBoxLayout()
        add_b_button = QPushButton("新增 B 类分量")
        add_b_button.clicked.connect(lambda: self.add_b_source_row())
        remove_b_button = QPushButton("删除选中分量")
        remove_b_button.clicked.connect(self.remove_selected_b_rows)
        b_button_row.addWidget(add_b_button)
        b_button_row.addWidget(remove_b_button)
        b_layout.addLayout(b_button_row)

        self.b_table = QTableWidget(0, 6)
        self.b_table.setHorizontalHeaderLabels(["类型", "名称", "输入值", "分布因子", "备注", "标准不确定度"])
        self.b_table.verticalHeader().setVisible(False)
        self.b_table.setAlternatingRowColors(True)
        self.b_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.b_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.b_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.b_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.b_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.b_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.b_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self.b_table.itemChanged.connect(self._on_b_table_item_changed)
        b_layout.addWidget(self.b_table, 1)

        settings_group = QGroupBox("报告设置")
        settings_layout = QVBoxLayout(settings_group)
        settings_form = QFormLayout()

        self.coverage_factor_spin = QDoubleSpinBox()
        self.coverage_factor_spin.setRange(0.1, 99.0)
        self.coverage_factor_spin.setDecimals(3)
        self.coverage_factor_spin.setSingleStep(0.1)
        self.coverage_factor_spin.valueChanged.connect(self._handle_user_edit)
        settings_form.addRow("覆盖因子 k", self.coverage_factor_spin)

        self.result_decimal_combo = QComboBox()
        self.result_decimal_combo.addItem("自动修约", None)
        for digits in range(0, 11):
            self.result_decimal_combo.addItem(f"固定 {digits} 位", digits)
        self.result_decimal_combo.currentIndexChanged.connect(self._handle_user_edit)
        settings_form.addRow("结果保留小数位", self.result_decimal_combo)

        settings_layout.addLayout(settings_form)
        settings_layout.addWidget(QLabel("项目备注"))

        self.notes_edit = QTextEdit()
        self.notes_edit.setPlaceholderText("可记录实验条件、仪器型号、B 类来源说明等")
        self.notes_edit.textChanged.connect(self._handle_user_edit)
        settings_layout.addWidget(self.notes_edit)

        refresh_button = QPushButton("刷新计算")
        refresh_button.setObjectName("primaryButton")
        refresh_button.clicked.connect(self.refresh_calculation)
        settings_layout.addWidget(refresh_button)

        layout.addWidget(b_group, 4)
        layout.addWidget(settings_group, 2)
        return panel

    def _build_right_panel(self) -> QWidget:
        panel = QWidget(self)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        self.summary_group = QGroupBox("结果总览")
        summary_layout = QVBoxLayout(self.summary_group)

        self.result_headline = QLabel("等待输入测量数据")
        self.result_headline.setObjectName("resultHeadline")
        self.result_headline.setWordWrap(True)
        self.result_headline.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.result_subline = QLabel("当前结果会随着输入自动更新。")
        self.result_subline.setObjectName("resultSubline")
        self.result_subline.setWordWrap(True)
        self.result_subline.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.warning_label = QLabel("当前没有警告。")
        self.warning_label.setObjectName("warningLabel")
        self.warning_label.setWordWrap(True)

        summary_layout.addWidget(self.result_headline)
        summary_layout.addWidget(self.result_subline)
        summary_layout.addWidget(self.warning_label)

        cards_layout = QGridLayout()
        cards_layout.setHorizontalSpacing(12)
        cards_layout.setVerticalSpacing(12)
        self.metric_labels: dict[str, QLabel] = {}
        card_titles = [
            ("mean", "平均值 x̄"),
            ("ua", "A类标准不确定度 uA"),
            ("ub", "B类合成标准不确定度 uB"),
            ("uc", "合成标准不确定度 uc"),
            ("U", "扩展不确定度 U"),
        ]

        for index, (key, title) in enumerate(card_titles):
            card, value_label = self._create_metric_card(title)
            row = index // 2
            column = index % 2
            cards_layout.addWidget(card, row, column)
            self.metric_labels[key] = value_label

        summary_layout.addLayout(cards_layout)
        layout.addWidget(self.summary_group, 2)

        lower_splitter = QSplitter(Qt.Orientation.Vertical)
        lower_splitter.setChildrenCollapsible(False)

        process_group = QGroupBox("计算过程")
        process_layout = QVBoxLayout(process_group)
        self.process_table = QTableWidget(0, 4)
        self.process_table.setHorizontalHeaderLabels(["类别", "项目", "数值", "说明"])
        self.process_table.verticalHeader().setVisible(False)
        self.process_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.process_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.process_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.process_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.process_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.process_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        process_layout.addWidget(self.process_table)
        lower_splitter.addWidget(process_group)

        report_group = QGroupBox("文本结论")
        report_layout = QVBoxLayout(report_group)
        self.report_text = QPlainTextEdit()
        self.report_text.setReadOnly(True)
        report_layout.addWidget(self.report_text)
        report_button_row = QHBoxLayout()
        copy_button = QPushButton("复制文本结论")
        copy_button.clicked.connect(self.copy_report_text)
        report_button_row.addStretch(1)
        report_button_row.addWidget(copy_button)
        report_layout.addLayout(report_button_row)
        lower_splitter.addWidget(report_group)
        lower_splitter.setStretchFactor(0, 3)
        lower_splitter.setStretchFactor(1, 2)

        layout.addWidget(lower_splitter, 5)
        return panel

    def _create_metric_card(self, title: str) -> tuple[QFrame, QLabel]:
        card = QFrame()
        card.setObjectName("metricCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 16, 16, 16)
        title_label = QLabel(title)
        title_label.setObjectName("metricTitle")
        value_label = QLabel("--")
        value_label.setObjectName("metricValue")
        value_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(title_label)
        layout.addWidget(value_label)
        return card, value_label

    def _load_initial_project(self) -> None:
        autosave_project = load_autosave()
        if autosave_project is not None:
            self.load_project_into_ui(autosave_project)
            self.statusBar().showMessage("已从自动保存恢复上次工作状态", 5000)
        else:
            path = last_project_path()
            if path and Path(path).exists():
                try:
                    self.open_project_file(path, prompt_on_error=False)
                except Exception:
                    self.load_project_into_ui(ProjectData())
            else:
                self.load_project_into_ui(ProjectData())
        self._refresh_recent_project_views()

    def load_project_into_ui(self, project: ProjectData) -> None:
        project.ensure_defaults()
        self.project = project
        self.current_project_path = project.project_path
        self._loading = True

        self.quantity_name_edit.setText(project.quantity_name)
        self.unit_edit.setText(project.unit)
        self.coverage_factor_spin.setValue(project.coverage_factor or 2.0)
        self._set_result_decimal_places(project.result_decimal_places)
        self.notes_edit.setPlainText(project.notes)
        self._populate_measurement_table(project.measured_values)
        self._populate_b_source_table(project.b_sources)
        self._set_import_info(project.last_import_path)

        self._loading = False
        self.refresh_calculation(mark_dirty=False)
        self.is_dirty = False
        self._update_window_title()

    def collect_project_from_ui(self) -> ProjectData:
        values = []
        for row in range(self.measurement_table.rowCount()):
            item = self.measurement_table.item(row, 0)
            if item is None:
                continue
            parsed = self._parse_float(item.text(), default=None)
            if parsed is not None:
                values.append(parsed)

        b_sources: list[BSource] = []
        for row in range(self.b_table.rowCount()):
            combo = self.b_table.cellWidget(row, 0)
            if not isinstance(combo, QComboBox):
                continue
            source_type = str(combo.currentData())
            name_item = self.b_table.item(row, 1)
            value_item = self.b_table.item(row, 2)
            divisor_item = self.b_table.item(row, 3)
            notes_item = self.b_table.item(row, 4)

            name = name_item.text().strip() if name_item else b_source_display_name(source_type)
            value = self._parse_float(value_item.text() if value_item else "", 0.0)
            divisor = self._parse_float(divisor_item.text() if divisor_item else "", default_divisor_for(source_type))
            notes = notes_item.text().strip() if notes_item else ""
            b_sources.append(
                BSource(
                    source_type=source_type,
                    name=name or b_source_display_name(source_type),
                    value=value,
                    divisor=divisor,
                    notes=notes,
                )
            )

        project = ProjectData(
            quantity_name=self.quantity_name_edit.text().strip(),
            unit=self.unit_edit.text().strip(),
            measured_values=values,
            b_sources=b_sources,
            coverage_factor=self.coverage_factor_spin.value(),
            result_decimal_places=self._selected_result_decimal_places(),
            notes=self.notes_edit.toPlainText().strip(),
            project_path=self.current_project_path,
            last_import_path=self.project.last_import_path,
            created_at=self.project.created_at,
            updated_at=self.project.updated_at,
        )
        project.ensure_defaults()
        return project

    def new_project(self) -> None:
        if not self._ensure_safe_to_continue():
            return
        project = ProjectData(coverage_factor=2.0)
        project.ensure_defaults()
        self.load_project_into_ui(project)
        self.current_project_path = None
        self.is_dirty = False
        self._update_window_title()
        self.statusBar().showMessage("已新建项目", 3000)

    def open_project_dialog(self) -> None:
        if not self._ensure_safe_to_continue():
            return
        path, _ = QFileDialog.getOpenFileName(
            self,
            "打开项目文件",
            str(Path.cwd()),
            "项目文件 (*.uncx);;JSON 文件 (*.json);;所有文件 (*)",
        )
        if path:
            self.open_project_file(path)

    def open_project_file(self, path: str, prompt_on_error: bool = True) -> None:
        try:
            project = load_project_file(path)
        except Exception as exc:
            if prompt_on_error:
                QMessageBox.critical(self, "打开失败", f"无法打开项目文件。\n\n{exc}")
            raise

        self.load_project_into_ui(project)
        self.current_project_path = path
        push_recent_file(path)
        set_last_project_path(path)
        self._refresh_recent_project_views()
        self.statusBar().showMessage(f"已打开项目: {Path(path).name}", 5000)

    def save_project(self) -> bool:
        if not self.current_project_path:
            return self.save_project_as()
        return self._save_project_to(self.current_project_path)

    def save_project_as(self) -> bool:
        suggested_name = self._default_file_stem() + ".uncx"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "保存项目文件",
            str(Path.cwd() / suggested_name),
            "项目文件 (*.uncx)",
        )
        if not path:
            return False
        if not path.lower().endswith(".uncx"):
            path += ".uncx"
        return self._save_project_to(path)

    def _save_project_to(self, path: str) -> bool:
        try:
            project = self.collect_project_from_ui()
            project = save_project_file(project, path)
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", f"无法保存项目文件。\n\n{exc}")
            return False

        self.project = project
        self.current_project_path = path
        self.is_dirty = False
        push_recent_file(path)
        set_last_project_path(path)
        self._refresh_recent_project_views()
        self._update_window_title()
        self._write_autosave_snapshot()
        self.statusBar().showMessage(f"已保存项目: {Path(path).name}", 5000)
        return True

    def import_excel_measurements(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "导入 Excel 测量数据",
            str(Path.cwd()),
            "Excel 文件 (*.xlsx *.xls)",
        )
        if not path:
            return

        try:
            values, sheet_name = import_measurements_from_excel(path)
        except Exception as exc:
            QMessageBox.warning(self, "导入失败", str(exc))
            return

        existing_values = self._measurement_values_from_table()
        if existing_values:
            message_box = QMessageBox(self)
            message_box.setWindowTitle("导入方式")
            message_box.setText("已有原始数据。是否用 Excel 数据替换当前测量值？")
            message_box.addButton("替换", QMessageBox.ButtonRole.YesRole)
            append_button = message_box.addButton("追加", QMessageBox.ButtonRole.NoRole)
            cancel_button = message_box.addButton("取消", QMessageBox.ButtonRole.RejectRole)
            message_box.exec()
            clicked_button = message_box.clickedButton()

            if clicked_button == cancel_button:
                return
            if clicked_button == append_button:
                values = existing_values + values

        self._populate_measurement_table(values)
        self.project.last_import_path = path
        self._set_import_info(path, sheet_name)
        self._handle_user_edit()
        self.statusBar().showMessage(f"已从 {Path(path).name} 导入 {len(values)} 个测量值", 5000)

    def export_excel(self) -> None:
        project, result = self._project_and_result_for_export()
        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出结果 Excel",
            str(Path.cwd() / f"{self._default_file_stem()}_结果.xlsx"),
            "Excel 文件 (*.xlsx)",
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        try:
            export_project_to_excel(project, result, path)
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", f"无法导出结果 Excel。\n\n{exc}")
            return

        self.statusBar().showMessage(f"已导出结果 Excel: {Path(path).name}", 5000)

    def export_data_excel(self) -> None:
        project = self.collect_project_from_ui()
        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出数据 Excel",
            str(Path.cwd() / f"{self._default_file_stem()}_数据.xlsx"),
            "Excel 文件 (*.xlsx)",
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        try:
            export_data_to_excel(project, path)
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", f"无法导出数据 Excel。\n\n{exc}")
            return

        self.statusBar().showMessage(f"已导出数据 Excel: {Path(path).name}", 5000)

    def export_txt(self) -> None:
        project, result = self._project_and_result_for_export()
        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出 TXT",
            str(Path.cwd() / f"{self._default_file_stem()}_结果.txt"),
            "文本文件 (*.txt)",
        )
        if not path:
            return
        if not path.lower().endswith(".txt"):
            path += ".txt"

        Path(path).write_text(result.summary_text, encoding="utf-8")
        self.statusBar().showMessage(f"已导出 TXT: {Path(path).name}", 5000)

    def export_image(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出结果摘要图",
            str(Path.cwd() / f"{self._default_file_stem()}_结果.png"),
            "PNG 图片 (*.png)",
        )
        if not path:
            return
        if not path.lower().endswith(".png"):
            path += ".png"

        pixmap = self.summary_group.grab()
        if not pixmap.save(path, "PNG"):
            QMessageBox.warning(self, "导出失败", "无法保存结果图片。")
            return
        self.statusBar().showMessage(f"已导出图片: {Path(path).name}", 5000)

    def copy_report_text(self) -> None:
        QApplication.clipboard().setText(self.report_text.toPlainText())
        self.statusBar().showMessage("已复制文本结论到剪贴板", 3000)

    def add_measurement_row(self, value: str = "", focus_and_edit: bool = False) -> int:
        row = self.measurement_table.rowCount()
        self.measurement_table.insertRow(row)
        item = QTableWidgetItem(value)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.measurement_table.setItem(row, 0, item)
        if focus_and_edit and not self._loading:
            QTimer.singleShot(0, lambda target_row=row: self._focus_measurement_row(target_row))
        if not self._loading:
            self._handle_user_edit()
        return row

    def remove_selected_measurement_rows(self) -> None:
        selected_rows = sorted({index.row() for index in self.measurement_table.selectionModel().selectedRows()}, reverse=True)
        if not selected_rows:
            return
        for row in selected_rows:
            self.measurement_table.removeRow(row)
        if self.measurement_table.rowCount() == 0:
            self.add_measurement_row()
        self._handle_user_edit()

    def clear_measurements(self) -> None:
        self._populate_measurement_table([])
        self._handle_user_edit()

    def paste_measurements_from_clipboard(self) -> None:
        text = QApplication.clipboard().text().strip()
        if not text:
            QMessageBox.information(self, "剪贴板为空", "当前剪贴板中没有可用的数据。")
            return

        tokens = [token for token in re.split(r"[\s,;，；]+", text) if token]
        values: list[float] = []
        for token in tokens:
            parsed = self._parse_float(token, default=None)
            if parsed is not None:
                values.append(parsed)

        if not values:
            QMessageBox.warning(self, "粘贴失败", "未在剪贴板中解析出有效数值。")
            return

        self._populate_measurement_table(values)
        self._handle_user_edit()
        self.statusBar().showMessage(f"已粘贴 {len(values)} 个测量值", 3000)

    def add_b_source_row(self, source: BSource | None = None) -> None:
        source = source or BSource(
            source_type=BSourceType.RESOLUTION.value,
            name=b_source_display_name(BSourceType.RESOLUTION.value),
            value=0.0,
            divisor=default_divisor_for(BSourceType.RESOLUTION.value),
        )

        row = self.b_table.rowCount()
        self.b_table.insertRow(row)

        combo = QComboBox()
        for source_type in BSourceType:
            combo.addItem(b_source_display_name(source_type.value), source_type.value)
        combo.currentIndexChanged.connect(self._on_b_source_type_changed)
        combo.setCurrentIndex(max(combo.findData(source.source_type), 0))
        self.b_table.setCellWidget(row, 0, combo)

        self._set_b_table_item(row, 1, source.name or b_source_display_name(source.source_type))
        self._set_b_table_item(row, 2, format_number(source.value))
        self._set_b_table_item(row, 3, format_number(source.divisor))
        self._set_b_table_item(row, 4, source.notes)
        self._set_b_table_item(row, 5, "0", editable=False)
        self._apply_b_row_rules(row)

        if not self._loading:
            self._handle_user_edit()

    def remove_selected_b_rows(self) -> None:
        selected_rows = sorted({index.row() for index in self.b_table.selectionModel().selectedRows()}, reverse=True)
        if not selected_rows:
            return
        for row in selected_rows:
            self.b_table.removeRow(row)
        if self.b_table.rowCount() == 0:
            self.add_b_source_row()
        self._handle_user_edit()

    def refresh_calculation(self, mark_dirty: bool = True) -> None:
        project = self.collect_project_from_ui()
        result = calculate_project(project)
        self.project = project
        self.latest_result = result
        self._update_b_standard_uncertainties(result)
        self._update_summary_panel(project, result)
        self._update_process_table(result)
        self.report_text.setPlainText(result.summary_text)

        if mark_dirty and not self._loading:
            self.is_dirty = True
            self._update_window_title()
            self.autosave_timer.start(800)

    def clear_recent_project_history(self) -> None:
        clear_recent_files()
        self._refresh_recent_project_views()
        self.statusBar().showMessage("已清空最近项目历史", 3000)

    def show_about_dialog(self) -> None:
        QMessageBox.about(
            self,
            "关于本程序",
            f"{APP_DISPLAY_NAME}\n版本 {APP_VERSION}\n作者: Leafuke\n\n"
            "用于物理实验课中单一物理量的不确定度计算。\n"
            "当前支持 A 类评定、B 类评定、合成标准不确定度、扩展不确定度，\n"
            "并支持项目保存、Excel 导入、数据/结果 Excel 导出、TXT 导出和结果图片导出。",
        )

    def closeEvent(self, event: QCloseEvent) -> None:
        if not self._ensure_safe_to_continue():
            event.ignore()
            return

        app_settings = settings()
        app_settings.setValue("geometry", self.saveGeometry())
        app_settings.setValue("windowState", self.saveState())
        self._write_autosave_snapshot()
        super().closeEvent(event)

    def _restore_window_state(self) -> None:
        app_settings = settings()
        geometry = app_settings.value("geometry")
        window_state = app_settings.value("windowState")
        if geometry:
            self.restoreGeometry(geometry)
        if window_state:
            self.restoreState(window_state)

    def _refresh_recent_project_views(self) -> None:
        items = recent_files()
        self.recent_projects_menu.clear()
        self.clear_recent_action.setEnabled(bool(items))

        if items:
            for path in items:
                action = QAction(Path(path).name, self)
                action.setToolTip(path)
                action.triggered.connect(lambda checked=False, target=path: self._open_recent_project_by_path(target))
                self.recent_projects_menu.addAction(action)
        else:
            empty_action = QAction("暂无最近项目", self)
            empty_action.setEnabled(False)
            self.recent_projects_menu.addAction(empty_action)

        self.recent_projects_menu.addSeparator()
        self.recent_projects_menu.addAction(self.clear_recent_action)
        button_text = f"最近项目 ({len(items)})" if items else "最近项目"
        self.recent_projects_button.setText(button_text)

    def _open_recent_project_by_path(self, path: str) -> None:
        if not self._ensure_safe_to_continue():
            return
        if not Path(path).exists():
            QMessageBox.warning(self, "文件不存在", "该项目文件已不存在，已从历史记录中移除。")
            self._refresh_recent_project_views()
            return
        self.open_project_file(path)

    def _on_measurement_table_item_changed(self, item: QTableWidgetItem) -> None:
        if self._loading:
            return
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self._handle_user_edit()

    def _on_b_table_item_changed(self, item: QTableWidgetItem) -> None:
        if self._loading:
            return
        if item.column() == 5:
            return
        self._handle_user_edit()

    def _on_b_source_type_changed(self) -> None:
        if self._loading:
            return
        combo = self.sender()
        if not isinstance(combo, QComboBox):
            return
        row = self._row_for_cell_widget(self.b_table, combo)
        if row < 0:
            return

        self._loading = True
        try:
            self._apply_b_row_rules(row)
        finally:
            self._loading = False
        self._handle_user_edit()

    def _apply_b_row_rules(self, row: int) -> None:
        combo = self.b_table.cellWidget(row, 0)
        if not isinstance(combo, QComboBox):
            return
        source_type = str(combo.currentData())

        name_item = self.b_table.item(row, 1)
        divisor_item = self.b_table.item(row, 3)
        current_name = name_item.text().strip() if name_item else ""
        known_display_names = {b_source_display_name(item.value) for item in BSourceType}
        if not current_name or current_name in known_display_names:
            self.b_table.item(row, 1).setText(b_source_display_name(source_type))

        if source_type == BSourceType.GIVEN_STDDEV.value:
            divisor_item.setText("1")
            divisor_item.setFlags((divisor_item.flags() | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled) & ~Qt.ItemFlag.ItemIsEditable)
        else:
            if self._parse_float(divisor_item.text(), default=None) in (None, 0.0):
                divisor_item.setText(format_number(default_divisor_for(source_type)))
            divisor_item.setFlags(divisor_item.flags() | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)

    def _populate_measurement_table(self, values: list[float]) -> None:
        self._loading = True
        self.measurement_table.setRowCount(0)
        if values:
            for value in values:
                self.add_measurement_row(format_number(value))
        self.add_measurement_row("")
        self._loading = False

    def _move_to_next_measurement_input(self, row: int) -> None:
        if self._loading:
            return

        current_item = self.measurement_table.item(row, 0)
        current_text = current_item.text().strip() if current_item else ""
        if not current_text:
            self._focus_measurement_row(row)
            return

        next_row = row + 1
        if next_row >= self.measurement_table.rowCount():
            self.add_measurement_row("")

        self._focus_measurement_row(next_row)

    def _focus_measurement_row(self, row: int) -> None:
        if row >= self.measurement_table.rowCount():
            row = self.add_measurement_row("")

        item = self.measurement_table.item(row, 0)
        if item is None:
            item = QTableWidgetItem("")
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.measurement_table.setItem(row, 0, item)

        self.measurement_table.setCurrentCell(row, 0)
        self.measurement_table.scrollToItem(item, QAbstractItemView.ScrollHint.PositionAtCenter)
        self.measurement_table.editItem(item)

    def _populate_b_source_table(self, sources: list[BSource]) -> None:
        self._loading = True
        self.b_table.setRowCount(0)
        if sources:
            for source in sources:
                self.add_b_source_row(source)
        else:
            self.add_b_source_row()
        self._loading = False

    def _set_b_table_item(self, row: int, column: int, text: str, editable: bool = True) -> None:
        item = QTableWidgetItem(text)
        if column in (2, 3, 5):
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        if not editable:
            item.setFlags((item.flags() | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled) & ~Qt.ItemFlag.ItemIsEditable)
        self.b_table.setItem(row, column, item)

    def _update_b_standard_uncertainties(self, result: CalculationResult) -> None:
        decimal_places = normalize_decimal_places(self.project.result_decimal_places)
        self._loading = True
        try:
            for row, component in enumerate(result.b_components):
                item = self.b_table.item(row, 5)
                if item is None:
                    self._set_b_table_item(
                        row,
                        5,
                        format_number(component.standard_uncertainty, decimal_places),
                        editable=False,
                    )
                else:
                    item.setText(format_number(component.standard_uncertainty, decimal_places))
        finally:
            self._loading = False

    def _update_summary_panel(self, project: ProjectData, result: CalculationResult) -> None:
        decimal_places = normalize_decimal_places(project.result_decimal_places)
        unit_suffix = f" {project.unit}" if project.unit else ""
        rounded_value, rounded_uncertainty = rounded_measurement(
            result.mean,
            result.expanded_uncertainty,
            decimal_places,
        )
        quantity_name = project.quantity_name or "测量量"
        self.result_headline.setText(
            f"{quantity_name} = ({rounded_value} ± {rounded_uncertainty}){unit_suffix}"
        )
        decimal_text = "自动修约" if decimal_places is None else f"固定 {decimal_places} 位小数"
        self.result_subline.setText(
            f"覆盖因子 k = {format_number(result.coverage_factor)}，{decimal_text}，当前按单一物理量的独立量模型计算。"
        )
        warning_text = "\n".join(result.warnings) if result.warnings else "当前输入有效，可直接导出为 Excel、TXT 或结果图片。"
        self.warning_label.setText(warning_text)

        self.metric_labels["mean"].setText(f"{format_number(result.mean, decimal_places)}{unit_suffix}")
        self.metric_labels["ua"].setText(f"{format_number(result.type_a_uncertainty, decimal_places)}{unit_suffix}")
        self.metric_labels["ub"].setText(f"{format_number(result.type_b_uncertainty, decimal_places)}{unit_suffix}")
        self.metric_labels["uc"].setText(f"{format_number(result.combined_uncertainty, decimal_places)}{unit_suffix}")
        self.metric_labels["U"].setText(f"{format_number(result.expanded_uncertainty, decimal_places)}{unit_suffix}")

    def _update_process_table(self, result: CalculationResult) -> None:
        self.process_table.setRowCount(len(result.process_rows))
        for row, row_data in enumerate(result.process_rows):
            for column, key in enumerate(["类别", "项目", "数值", "说明"]):
                item = QTableWidgetItem(row_data.get(key, ""))
                if column == 2:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.process_table.setItem(row, column, item)

    def _project_and_result_for_export(self) -> tuple[ProjectData, CalculationResult]:
        project = self.collect_project_from_ui()
        result = calculate_project(project)
        return project, result

    def _measurement_values_from_table(self) -> list[float]:
        values: list[float] = []
        for row in range(self.measurement_table.rowCount()):
            item = self.measurement_table.item(row, 0)
            if item is None:
                continue
            value = self._parse_float(item.text(), default=None)
            if value is not None:
                values.append(value)
        return values

    def _selected_result_decimal_places(self) -> int | None:
        return normalize_decimal_places(self.result_decimal_combo.currentData())

    def _set_result_decimal_places(self, decimal_places: int | None) -> None:
        normalized = normalize_decimal_places(decimal_places)
        target_index = self.result_decimal_combo.findData(normalized)
        if target_index < 0:
            target_index = 0
        self.result_decimal_combo.setCurrentIndex(target_index)

    def _set_import_info(self, path: str | None, sheet_name: str | None = None) -> None:
        if not path:
            self.import_info_label.setText("尚未导入 Excel 数据")
            return
        path_text = Path(path).name
        if sheet_name:
            self.import_info_label.setText(f"最近导入: {path_text} / 工作表: {sheet_name}")
        else:
            self.import_info_label.setText(f"最近导入: {path_text}")

    def _write_autosave_snapshot(self) -> None:
        try:
            project = self.collect_project_from_ui()
            write_autosave(project)
        except Exception:
            return

    def _ensure_safe_to_continue(self) -> bool:
        if not self.is_dirty:
            return True

        message_box = QMessageBox(self)
        message_box.setWindowTitle("保存当前更改")
        message_box.setText("当前项目有未保存的更改。是否先保存？")
        save_button = message_box.addButton("保存", QMessageBox.ButtonRole.AcceptRole)
        discard_button = message_box.addButton("不保存", QMessageBox.ButtonRole.DestructiveRole)
        cancel_button = message_box.addButton("取消", QMessageBox.ButtonRole.RejectRole)
        message_box.exec()

        clicked_button = message_box.clickedButton()
        if clicked_button == cancel_button:
            return False
        if clicked_button == save_button:
            return self.save_project()
        return clicked_button == discard_button

    def _default_file_stem(self) -> str:
        base = self.quantity_name_edit.text().strip() or "不确定度项目"
        cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", base)
        return cleaned

    def _parse_float(self, text: str, default: float | None = 0.0) -> float | None:
        cleaned = text.strip().replace("，", ".")
        if not cleaned:
            return default
        if cleaned.count(",") == 1 and "." not in cleaned:
            cleaned = cleaned.replace(",", ".")
        cleaned = cleaned.replace(" ", "")
        try:
            return float(cleaned)
        except ValueError:
            return default

    def _row_for_cell_widget(self, table: QTableWidget, widget: QWidget) -> int:
        for row in range(table.rowCount()):
            if table.cellWidget(row, 0) is widget:
                return row
        return -1

    def _handle_user_edit(self) -> None:
        if self._loading:
            return
        self.refresh_calculation(mark_dirty=True)

    def _update_window_title(self) -> None:
        project_name = Path(self.current_project_path).name if self.current_project_path else "未保存项目"
        dirty_mark = " *" if self.is_dirty else ""
        self.setWindowTitle(f"{APP_DISPLAY_NAME} {dirty_mark}- {project_name}")


def run() -> None:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setApplicationDisplayName(APP_DISPLAY_NAME)
    app.setApplicationName(APP_DISPLAY_NAME)
    app.setStyleSheet(APP_STYLE_SHEET)

    icon = _load_app_icon()
    if icon is not None:
        app.setWindowIcon(icon)

    window = MainWindow()
    if icon is not None:
        window.setWindowIcon(icon)
    window.show()
    app.exec()


def _load_app_icon() -> QIcon | None:
    project_root = Path(__file__).resolve().parent.parent
    candidate_paths = [
        project_root / "app.ico",
        project_root / "assets" / "app.ico",
    ]
    for path in candidate_paths:
        if path.exists():
            return QIcon(str(path))
    return None