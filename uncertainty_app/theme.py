from __future__ import annotations

from dataclasses import asdict, dataclass

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication


@dataclass(frozen=True)
class ThemeColors:
    window_bg: str
    toolbar_bg: str
    status_bg: str
    border: str
    subtle_border: str
    label_text: str
    selection_bg: str
    selection_text: str
    hover_bg: str
    menu_bg: str
    menu_selected_bg: str
    menu_selected_text: str
    group_text: str
    group_bg: str
    input_bg: str
    input_border: str
    input_focus: str
    button_bg: str
    button_text: str
    button_border: str
    button_hover: str
    primary_button_bg: str
    primary_button_hover: str
    header_bg: str
    header_text: str
    card_start: str
    card_end: str
    card_border: str
    metric_title: str
    metric_value: str
    result_headline: str
    result_subline: str
    warning_bg: str
    warning_text: str
    warning_border: str
    info_bg: str
    info_text: str


STYLE_TEMPLATE = """
QMainWindow {
    background: {window_bg};
    color: {label_text};
}

QToolBar {
    spacing: 6px;
    padding: 5px 8px;
    background: {toolbar_bg};
    color: {label_text};
    border: none;
    border-bottom: 1px solid {border};
}

QStatusBar {
    background: {status_bg};
    color: {label_text};
    border-top: 1px solid {border};
}

QToolBar::separator {
    width: 1px;
    margin: 4px 6px;
    background: {border};
}

QToolButton {
    padding: 6px 10px;
    border-radius: 10px;
    color: {button_text};
}

QToolButton:hover {
    background: {hover_bg};
}

QMenu {
    background: {menu_bg};
    color: {label_text};
    border: 1px solid {border};
    border-radius: 12px;
    padding: 6px;
}

QMenu::item {
    padding: 8px 14px;
    border-radius: 8px;
}

QMenu::item:selected {
    background: {menu_selected_bg};
    color: {menu_selected_text};
}

QGroupBox {
    font-size: 14px;
    font-weight: 600;
    color: {group_text};
    background: {group_bg};
    border: 1px solid {border};
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
    color: {label_text};
}

QLineEdit,
QPlainTextEdit,
QTextEdit,
QTableWidget,
QListWidget,
QComboBox,
QDoubleSpinBox,
QTextBrowser {
    background: {input_bg};
    color: {label_text};
    border: 1px solid {input_border};
    border-radius: 12px;
    padding: 6px 8px;
    selection-background-color: {selection_bg};
    selection-color: {selection_text};
}

QLineEdit:focus,
QPlainTextEdit:focus,
QTextEdit:focus,
QTableWidget:focus,
QListWidget:focus,
QComboBox:focus,
QDoubleSpinBox:focus,
QTextBrowser:focus {
    border: 1px solid {input_focus};
}

QPushButton {
    background: {button_bg};
    color: {button_text};
    border: 1px solid {button_border};
    border-radius: 12px;
    padding: 8px 12px;
}

QPushButton:hover {
    background: {button_hover};
}

QPushButton#primaryButton {
    background: {primary_button_bg};
    color: #ffffff;
    border: none;
}

QPushButton#primaryButton:hover {
    background: {primary_button_hover};
}

QHeaderView::section {
    background: {header_bg};
    color: {header_text};
    border: none;
    border-bottom: 1px solid {border};
    padding: 8px;
    font-weight: 600;
}

QFrame#metricCard {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1, stop: 0 {card_start}, stop: 1 {card_end});
    border: 1px solid {card_border};
    border-radius: 18px;
}

QLabel#metricTitle {
    color: {metric_title};
    font-size: 12px;
    font-weight: 600;
}

QLabel#metricValue {
    color: {metric_value};
    font-size: 24px;
    font-weight: 700;
}

QLabel#resultHeadline {
    color: {result_headline};
    font-size: 20px;
    font-weight: 700;
}

QLabel#resultSubline {
    color: {result_subline};
    font-size: 13px;
}

QLabel#warningLabel {
    background: {warning_bg};
    color: {warning_text};
    border: 1px solid {warning_border};
    border-radius: 12px;
    padding: 10px 12px;
}

QLabel#infoPill {
    color: {info_text};
    background: {info_bg};
    border-radius: 12px;
    padding: 8px 10px;
}

QCheckBox {
    color: {label_text};
    spacing: 8px;
}
"""


LIGHT_THEME = ThemeColors(
    window_bg="#f3f6fb",
    toolbar_bg="#ffffff",
    status_bg="#ffffff",
    border="#d8e0ea",
    subtle_border="#c7d2e2",
    label_text="#243b53",
    selection_bg="#cfe1ff",
    selection_text="#102a43",
    hover_bg="#eef4ff",
    menu_bg="#ffffff",
    menu_selected_bg="#eef4ff",
    menu_selected_text="#12345a",
    group_text="#243b53",
    group_bg="#ffffff",
    input_bg="#fbfcfe",
    input_border="#c7d2e2",
    input_focus="#4c8bf5",
    button_bg="#e8eef9",
    button_text="#16324f",
    button_border="#c4d4ec",
    button_hover="#dce8fb",
    primary_button_bg="#1f5fbf",
    primary_button_hover="#184d9a",
    header_bg="#eff4fb",
    header_text="#243b53",
    card_start="#ffffff",
    card_end="#edf5ff",
    card_border="#d1ddef",
    metric_title="#5b7083",
    metric_value="#12345a",
    result_headline="#12345a",
    result_subline="#486581",
    warning_bg="#fff4e5",
    warning_text="#9c5b00",
    warning_border="#ffd59c",
    info_bg="#edf3fb",
    info_text="#486581",
)


DARK_THEME = ThemeColors(
    window_bg="#161b22",
    toolbar_bg="#1d2430",
    status_bg="#1d2430",
    border="#304055",
    subtle_border="#3f516b",
    label_text="#e6edf5",
    selection_bg="#365e9d",
    selection_text="#ffffff",
    hover_bg="#29374a",
    menu_bg="#1d2430",
    menu_selected_bg="#314462",
    menu_selected_text="#f7fbff",
    group_text="#f3f7fc",
    group_bg="#202833",
    input_bg="#111722",
    input_border="#41516b",
    input_focus="#6ea8ff",
    button_bg="#243245",
    button_text="#eaf2ff",
    button_border="#41516b",
    button_hover="#30415b",
    primary_button_bg="#3b82f6",
    primary_button_hover="#2f6fd1",
    header_bg="#202938",
    header_text="#e6edf5",
    card_start="#222c39",
    card_end="#18212c",
    card_border="#314155",
    metric_title="#9aaabc",
    metric_value="#f7fbff",
    result_headline="#f7fbff",
    result_subline="#b9c7d6",
    warning_bg="#3f2d12",
    warning_text="#ffd79d",
    warning_border="#8f6a2d",
    info_bg="#202938",
    info_text="#b9c7d6",
)


def _palette(colors: ThemeColors) -> QPalette:
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(colors.window_bg))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(colors.label_text))
    palette.setColor(QPalette.ColorRole.Base, QColor(colors.input_bg))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(colors.group_bg))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(colors.group_bg))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(colors.label_text))
    palette.setColor(QPalette.ColorRole.Text, QColor(colors.label_text))
    palette.setColor(QPalette.ColorRole.Button, QColor(colors.button_bg))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(colors.button_text))
    palette.setColor(QPalette.ColorRole.BrightText, QColor("#ffffff"))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(colors.selection_bg))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(colors.selection_text))
    palette.setColor(QPalette.ColorRole.Link, QColor(colors.primary_button_bg))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(colors.metric_title))
    return palette


def _render_stylesheet(colors: ThemeColors) -> str:
    stylesheet = STYLE_TEMPLATE
    for key, value in asdict(colors).items():
        stylesheet = stylesheet.replace(f"{{{key}}}", value)
    return stylesheet


def apply_application_theme(app: QApplication, color_scheme: Qt.ColorScheme | None = None) -> Qt.ColorScheme:
    scheme = color_scheme if color_scheme is not None else app.styleHints().colorScheme()
    colors = DARK_THEME if scheme == Qt.ColorScheme.Dark else LIGHT_THEME
    app.setPalette(_palette(colors))
    app.setStyleSheet(_render_stylesheet(colors))
    return scheme