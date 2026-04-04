import io
import os
import subprocess
import sys
import time
import traceback
from collections import deque
from contextlib import redirect_stdout

from PySide6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QObject,
    QRunnable,
    QSortFilterProxyModel,
    Qt,
    QThreadPool,
    QTimer,
    Signal,
)
from PySide6.QtGui import QIcon, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QInputDialog,
    QPushButton,
    QPlainTextEdit,
    QDialog,
    QDialogButtonBox,
    QStyle,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

import app_path_migration_core as core
from app_path_manager import PathConfig
from app_logger import setup_logger
from . import __version__

logger = setup_logger(__name__)

# 从 PathConfig 获取 UI 状态文件路径
UI_STATE_FILE = PathConfig.UI_STATE_FILE


def preferred_migration_roots():
    return r"D:\Program Files", r"D:\Program Files (x86)"


class WorkerSignals(QObject):
    finished = Signal(object, str)
    failed = Signal(str, str)


class Worker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    def run(self):
        buf = io.StringIO()
        try:
            with redirect_stdout(buf):
                result = self.fn(*self.args, **self.kwargs)
            self.signals.finished.emit(result, buf.getvalue())
        except Exception:
            self.signals.failed.emit(traceback.format_exc(), buf.getvalue())


class AppTableModel(QAbstractTableModel):
    checkedCountChanged = Signal(int)

    def __init__(self, icon_provider, parent=None):
        super().__init__(parent)
        self._icon_provider = icon_provider
        self._apps = []
        self._checked_ids = set()
        self._headers = ["选择", "图标", "应用名", "位数", "安装日期", "安装路径", "发布者"]
        self._icon_rows = {}

    def _app_id(self, app):
        return (
            os.path.normcase(str(app.get("install_dir", "") or "")),
            str(app.get("display_name", "") or "").strip().lower(),
        )

    def app_id(self, app):
        return self._app_id(app)

    def rowCount(self, parent=QModelIndex()):
        if parent.isValid():
            return 0
        return len(self._apps)

    def columnCount(self, parent=QModelIndex()):
        if parent.isValid():
            return 0
        return len(self._headers)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal and 0 <= section < len(self._headers):
            return self._headers[section]
        return str(section + 1)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        row = index.row()
        col = index.column()
        if row < 0 or row >= len(self._apps):
            return None

        app = self._apps[row]
        app_id = self._app_id(app)

        if col == 0:
            if role == Qt.CheckStateRole:
                return Qt.Checked if app_id in self._checked_ids else Qt.Unchecked
            if role == Qt.DisplayRole:
                return ""
            return None

        if col == 1:
            if role == Qt.DecorationRole:
                return self._icon_provider(app)
            if role == Qt.DisplayRole:
                return ""
            return None

        if role == Qt.DisplayRole:
            if col == 2:
                return app.get("display_name", "")
            if col == 3:
                return app.get("arch", "") or "未知"
            if col == 4:
                return app.get("install_date", "") or ""
            if col == 5:
                return app.get("install_dir", "")
            if col == 6:
                return app.get("publisher", "")

        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if index.column() == 0:
            flags |= Qt.ItemIsUserCheckable
        return flags

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid() or index.column() != 0 or role != Qt.CheckStateRole:
            return False

        row = index.row()
        if row < 0 or row >= len(self._apps):
            return False

        app_id = self._app_id(self._apps[row])
        checked = int(value) == int(Qt.Checked)
        if checked:
            self._checked_ids.add(app_id)
        else:
            self._checked_ids.discard(app_id)

        self.dataChanged.emit(index, index, [Qt.CheckStateRole])
        self.checkedCountChanged.emit(len(self._checked_ids))
        return True

    def sort(self, column, order=Qt.AscendingOrder):
        if not self._apps:
            return

        reverse = order == Qt.DescendingOrder

        def key_func(app):
            if column == 0:
                return self._app_id(app) in self._checked_ids
            if column == 3:
                return str(app.get("arch", "") or "")
            if column == 4:
                return str(app.get("install_date", "") or "")
            if column == 5:
                return str(app.get("install_dir", "") or "").lower()
            if column == 6:
                return str(app.get("publisher", "") or "").lower()
            return str(app.get("display_name", "") or "").lower()

        self.layoutAboutToBeChanged.emit()
        self._apps.sort(key=key_func, reverse=reverse)
        self.layoutChanged.emit()

    def set_apps(self, apps, preserve_checks=False):
        new_apps = list(apps or [])
        if preserve_checks:
            new_ids = {self._app_id(app) for app in new_apps}
            self._checked_ids &= new_ids
        else:
            self._checked_ids.clear()

        icon_rows = {}
        for i, app in enumerate(new_apps):
            icon_key = str(app.get("_icon_key", "") or "")
            if not icon_key:
                continue
            icon_rows.setdefault(icon_key, []).append(i)

        self.beginResetModel()
        self._apps = new_apps
        self._icon_rows = icon_rows
        self.endResetModel()
        self.checkedCountChanged.emit(len(self._checked_ids))

    def set_checked_all(self, checked):
        if not self._apps:
            self._checked_ids.clear()
            self.checkedCountChanged.emit(0)
            return

        if checked:
            self._checked_ids = {self._app_id(app) for app in self._apps}
        else:
            self._checked_ids.clear()

        top_left = self.index(0, 0)
        bottom_right = self.index(len(self._apps) - 1, 0)
        self.dataChanged.emit(top_left, bottom_right, [Qt.CheckStateRole])
        self.checkedCountChanged.emit(len(self._checked_ids))

    def clear_checks(self):
        if not self._apps and not self._checked_ids:
            return
        self._checked_ids.clear()
        if self._apps:
            top_left = self.index(0, 0)
            bottom_right = self.index(len(self._apps) - 1, 0)
            self.dataChanged.emit(top_left, bottom_right, [Qt.CheckStateRole])
        self.checkedCountChanged.emit(0)

    def set_checked_ids(self, checked_ids):
        self._checked_ids = set(checked_ids or set())
        if self._apps:
            top_left = self.index(0, 0)
            bottom_right = self.index(len(self._apps) - 1, 0)
            self.dataChanged.emit(top_left, bottom_right, [Qt.CheckStateRole])
        self.checkedCountChanged.emit(len(self._checked_ids))

    def notify_icon_loaded(self, icon_key):
        rows = self._icon_rows.get(icon_key, [])
        for row in rows:
            idx = self.index(row, 1)
            if idx.isValid():
                self.dataChanged.emit(idx, idx, [Qt.DecorationRole])

    def checked_count(self):
        return len(self._checked_ids)

    def selected_apps(self):
        return [app for app in self._apps if self._app_id(app) in self._checked_ids]

    def app_at_row(self, row):
        if row < 0 or row >= len(self._apps):
            return None
        return self._apps[row]


class AppFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.exclude_standard = False
        self.exclude_users = True
        self.keywords = tuple()
        self.search_text = ""

    def set_filter_options(self, exclude_standard, exclude_users, keywords, search_text):
        new_keywords = tuple(sorted([k.lower() for k in keywords if k])) if keywords else tuple()
        changed = (
            self.exclude_standard != bool(exclude_standard)
            or self.exclude_users != bool(exclude_users)
            or self.keywords != new_keywords
            or self.search_text != (search_text or "").strip().lower()
        )
        self.exclude_standard = bool(exclude_standard)
        self.exclude_users = bool(exclude_users)
        self.keywords = new_keywords
        self.search_text = (search_text or "").strip().lower()

        if changed:
            self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()
        if model is None:
            return False
        app = model.app_at_row(source_row)
        if not app:
            return False

        if self.exclude_standard and app.get("_is_standard_path", False):
            return False

        if self.exclude_users and app.get("_is_users_path", False):
            return False

        if self.keywords:
            keyword_blob = app.get("_keyword_blob", "")
            if any(k in keyword_blob for k in self.keywords):
                return False

        if self.search_text:
            search_blob = app.get("_search_blob", "")
            if self.search_text not in search_blob:
                return False

        return True


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"应用安装路径迁移工具 - GUI v{__version__}")
        self.resize(1200, 760)

        self.thread_pool = QThreadPool.globalInstance()
        self.raw_apps = []
        self.active_workers = set()
        self.is_busy = False
        self.icon_cache = {}
        self.missing_icon_paths = set()
        self.install_dir_icon_cache = {}
        self.default_file_icon = self.style().standardIcon(QStyle.SP_FileIcon)
        self.pending_icon_keys = deque()
        self.pending_icon_key_set = set()
        self.icon_loader_timer = QTimer(self)
        self.icon_loader_timer.setSingleShot(True)
        self.icon_loader_timer.timeout.connect(self._process_pending_icons)
        self.icon_load_batch_size = 12
        self.batch_rows = []
        # 区分复选框和搜索框的防抖延迟
        self.checkbox_filter_timer = QTimer(self)
        self.checkbox_filter_timer.setSingleShot(True)
        self.checkbox_filter_timer.setInterval(50)  # 复选框立即预览
        self.checkbox_filter_timer.timeout.connect(self._execute_pending_checkbox_filter)

        self.search_filter_timer = QTimer(self)
        self.search_filter_timer.setSingleShot(True)
        self.search_filter_timer.setInterval(300)  # 搜索输入延迟
        self.search_filter_timer.timeout.connect(self._execute_pending_search_filter)

        self.batch_table_updating = False
        self._pending_checkbox_filter = None
        self._pending_search_filter = None

        root = QWidget(self)
        self.setCentralWidget(root)
        main_layout = QVBoxLayout(root)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_app_migration_page(), "应用迁移")
        self.tabs.addTab(self._build_batch_center_page(), "操作记录")
        self.log_tab = self._build_log_page()
        self.tabs.addTab(self.log_tab, "日志")
        self.program_info_tab = self._build_program_info_page()
        self.tabs.addTab(self.program_info_tab, "程序说明")
        main_layout.addWidget(self.tabs)

        self.drive_fix_dialog = self._create_drive_fix_dialog()

        tools_menu = self.menuBar().addMenu("工具")
        self.open_drive_fix_action = tools_menu.addAction("盘符修复窗口")
        self.open_drive_fix_action.triggered.connect(self.open_drive_fix_dialog)
        self.drive_restore_action = tools_menu.addAction("恢复盘符修复记录")
        self.drive_restore_action.triggered.connect(self.restore_drive_fix_batch)

        settings_menu = self.menuBar().addMenu("设置")
        self.restore_defaults_action = settings_menu.addAction("恢复默认设置")
        self.restore_defaults_action.triggered.connect(self.restore_ui_defaults)

        self.apply_ui_state(self.load_ui_state())
        self.bind_ui_state_events()

        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        self.copy_shortcut.activated.connect(self.copy_selected_cells)
        self.find_shortcut = QShortcut(QKeySequence.Find, self)
        self.find_shortcut.activated.connect(self.focus_app_search)
        self.drive_fix_shortcut = QShortcut(QKeySequence("Ctrl+Shift+D"), self)
        self.drive_fix_shortcut.activated.connect(self.open_drive_fix_dialog)

        self._set_busy(False)
        self.refresh_batch_center()
        # Delay initial scan until the event loop starts so the window becomes responsive first.
        QTimer.singleShot(0, self.scan_apps)

    def _create_drive_fix_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("盘符修复")
        dlg.resize(840, 360)
        dlg.setModal(False)

        layout = QVBoxLayout(dlg)
        self.drive_status_label = QLabel("最近盘符修复记录: (无)")
        self.drive_status_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        layout.addWidget(self.drive_status_label)
        layout.addWidget(self._build_drive_fix_box())
        return dlg

    def open_drive_fix_dialog(self):
        self.drive_fix_dialog.show()
        self.drive_fix_dialog.raise_()
        self.drive_fix_dialog.activateWindow()

    def _build_app_migration_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        self.migration_status_label = QLabel("最近应用迁移记录: (无)")
        self.migration_status_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        workbench = QGroupBox("迁移工作台")
        workbench_layout = QVBoxLayout(workbench)
        workbench_header = QHBoxLayout()
        self.workbench_details_cb = QCheckBox("设置明细")
        self.workbench_details_cb.setChecked(True)
        self.workbench_details_cb.toggled.connect(
            lambda checked: self.workbench_details_panel.setVisible(checked)
        )
        workbench_header.addWidget(self.workbench_details_cb)
        workbench_header.addStretch(1)
        workbench_layout.addLayout(workbench_header)

        self.workbench_details_panel = QWidget()
        workbench_top = QHBoxLayout(self.workbench_details_panel)
        workbench_top.setContentsMargins(0, 0, 0, 0)
        workbench_top.addWidget(self._build_filter_box(), 3)
        workbench_top.addWidget(self._build_target_box(), 4)
        workbench_layout.addWidget(self.workbench_details_panel)
        workbench_layout.addWidget(self._build_action_bar())

        app_list_box = QGroupBox("应用列表")
        app_list_layout = QVBoxLayout(app_list_box)
        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("快速搜索:"))
        self.app_search_edit = QLineEdit()
        self.app_search_edit.setPlaceholderText("按应用名/路径/发布者过滤")
        self.clear_search_btn = QPushButton("清空")
        self.select_all_btn = QPushButton("全选")
        self.clear_select_btn = QPushButton("清空选择")
        self.selection_status_label = QLabel("已选 0 项")
        self.selection_status_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.app_search_edit.textChanged.connect(self.request_search_filter)
        self.clear_search_btn.clicked.connect(self.clear_app_search)
        self.select_all_btn.clicked.connect(self.select_all)
        self.clear_select_btn.clicked.connect(self.clear_selection)
        search_row.addWidget(self.app_search_edit, 1)
        search_row.addWidget(self.clear_search_btn)
        search_row.addWidget(self.select_all_btn)
        search_row.addWidget(self.clear_select_btn)
        search_row.addWidget(self.selection_status_label)
        app_list_layout.addLayout(search_row)
        app_list_layout.addWidget(self._build_table())

        layout.addWidget(self.migration_status_label)
        layout.addWidget(workbench)
        layout.addWidget(app_list_box, 1)
        return page

    def _build_batch_center_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        action_row = QHBoxLayout()
        self.batch_type_filter = QComboBox()
        self.batch_type_filter.addItems(["全部", "应用迁移", "盘符修复"])
        self.batch_status_filter = QComboBox()
        self.batch_status_filter.addItems(["全部", "applied", "restored"])
        self.batch_refresh_btn = QPushButton("刷新记录列表")
        self.batch_detail_btn = QPushButton("查看记录详情")
        self.batch_copy_btn = QPushButton("复制记录摘要")
        self.batch_delete_btn = QPushButton("删除所选记录")
        self.batch_restore_migration_btn = QPushButton("恢复应用迁移记录")
        self.batch_restore_drive_btn = QPushButton("恢复盘符修复记录")
        self.batch_bulk_restore_btn = QPushButton("恢复勾选记录")
        self.batch_bulk_delete_btn = QPushButton("删除勾选记录")
        self.batch_select_all_cb = QCheckBox("全选")
        self.batch_checked_status_label = QLabel("已勾选 0 项")

        self.batch_type_filter.currentIndexChanged.connect(self.refresh_batch_center)
        self.batch_status_filter.currentIndexChanged.connect(self.refresh_batch_center)
        self.batch_refresh_btn.clicked.connect(self.refresh_batch_center)
        self.batch_detail_btn.clicked.connect(self.show_selected_batch_details)
        self.batch_copy_btn.clicked.connect(self.copy_selected_batch_summary)
        self.batch_delete_btn.clicked.connect(self.delete_selected_record)
        self.batch_restore_migration_btn.clicked.connect(self.restore_batch)
        self.batch_restore_drive_btn.clicked.connect(self.restore_drive_fix_batch)
        self.batch_bulk_restore_btn.clicked.connect(self.restore_checked_records)
        self.batch_bulk_delete_btn.clicked.connect(self.delete_checked_records)
        self.batch_select_all_cb.toggled.connect(self.toggle_check_all_batches)

        action_row.addWidget(QLabel("类型筛选:"))
        action_row.addWidget(self.batch_type_filter)
        action_row.addWidget(QLabel("状态筛选:"))
        action_row.addWidget(self.batch_status_filter)
        action_row.addWidget(self.batch_refresh_btn)
        action_row.addWidget(self.batch_detail_btn)
        action_row.addWidget(self.batch_copy_btn)
        action_row.addWidget(self.batch_delete_btn)
        action_row.addWidget(self.batch_restore_migration_btn)
        action_row.addWidget(self.batch_restore_drive_btn)
        action_row.addWidget(self.batch_bulk_restore_btn)
        action_row.addWidget(self.batch_bulk_delete_btn)
        action_row.addWidget(self.batch_select_all_cb)
        action_row.addWidget(self.batch_checked_status_label)
        action_row.addStretch(1)
        layout.addLayout(action_row)

        self.batch_table = QTableWidget(0, 7)
        self.batch_table.setHorizontalHeaderLabels(
            ["选", "类型", "记录ID", "状态", "创建时间", "恢复时间", "描述"]
        )
        self.batch_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.batch_table.horizontalHeader().setStretchLastSection(True)
        self.batch_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.batch_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.batch_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.batch_table.setColumnWidth(0, 54)
        self.batch_table.setColumnWidth(1, 120)
        self.batch_table.setColumnWidth(2, 180)
        self.batch_table.setColumnWidth(3, 90)
        self.batch_table.setColumnWidth(4, 180)
        self.batch_table.setColumnWidth(5, 180)
        self.batch_table.itemDoubleClicked.connect(lambda _item: self.show_selected_batch_details())
        self.batch_table.itemChanged.connect(self.on_batch_table_item_changed)
        self.batch_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.batch_table.customContextMenuRequested.connect(self.show_batch_context_menu)
        layout.addWidget(self.batch_table, 1)

        return page

    def refresh_batch_center(self):
        checked_ids = set()
        if self.batch_rows and self.batch_table.rowCount() > 0:
            for i, old_row in enumerate(self.batch_rows):
                item = self.batch_table.item(i, 0)
                if item is not None and item.checkState() == Qt.Checked:
                    checked_ids.add(old_row.get("id", ""))

        rows = []
        migration_batches = core.list_batches(applied_only=False)
        drive_fix_batches = core.list_drive_fix_batches(applied_only=False)

        for b in migration_batches:
            rows.append(
                {
                    "kind": "migration",
                    "type": "应用迁移",
                    "id": b.get("id", ""),
                    "status": b.get("status", ""),
                    "created_at": b.get("created_at", ""),
                    "restored_at": b.get("restored_at", ""),
                    "desc": f"应用数 {len(b.get('apps', []))}",
                    "record": b,
                }
            )

        for b in drive_fix_batches:
            rows.append(
                {
                    "kind": "drive_fix",
                    "type": "盘符修复",
                    "id": b.get("id", ""),
                    "status": b.get("status", ""),
                    "created_at": b.get("created_at", ""),
                    "restored_at": b.get("restored_at", ""),
                    "desc": f"{b.get('old_drive', '')}->{b.get('new_drive', '')}",
                    "record": b,
                }
            )

        rows.sort(key=lambda x: x.get("created_at", ""), reverse=True)

        selected_filter = (
            self.batch_type_filter.currentText() if hasattr(self, "batch_type_filter") else "全部"
        )
        if selected_filter != "全部":
            rows = [r for r in rows if r.get("type") == selected_filter]

        status_filter = (
            self.batch_status_filter.currentText()
            if hasattr(self, "batch_status_filter")
            else "全部"
        )
        if status_filter != "全部":
            rows = [r for r in rows if r.get("status") == status_filter]

        self.batch_rows = rows
        self.batch_table_updating = True
        self.batch_table.setUpdatesEnabled(False)
        self.batch_table.setRowCount(len(self.batch_rows))
        for i, row in enumerate(self.batch_rows):
            check_item = QTableWidgetItem()
            check_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            check_item.setCheckState(
                Qt.Checked if row.get("id", "") in checked_ids else Qt.Unchecked
            )
            self.batch_table.setItem(i, 0, check_item)
            self.batch_table.setItem(i, 1, QTableWidgetItem(row.get("type", "")))
            self.batch_table.setItem(i, 2, QTableWidgetItem(row.get("id", "")))
            self.batch_table.setItem(i, 3, QTableWidgetItem(row.get("status", "")))
            self.batch_table.setItem(i, 4, QTableWidgetItem(row.get("created_at", "")))
            self.batch_table.setItem(i, 5, QTableWidgetItem(row.get("restored_at", "")))
            self.batch_table.setItem(i, 6, QTableWidgetItem(row.get("desc", "")))

        self.batch_table.setUpdatesEnabled(True)
        self.batch_table_updating = False
        self.update_batch_checked_count()

        self.update_latest_status_labels(migration_batches, drive_fix_batches)

    def update_latest_status_labels(self, migration_batches=None, drive_fix_batches=None):
        migration_batches = (
            migration_batches
            if migration_batches is not None
            else core.list_batches(applied_only=False)
        )
        drive_fix_batches = (
            drive_fix_batches
            if drive_fix_batches is not None
            else core.list_drive_fix_batches(applied_only=False)
        )

        latest_migration = None
        for b in migration_batches:
            if latest_migration is None or b.get("created_at", "") > latest_migration.get(
                "created_at", ""
            ):
                latest_migration = b

        latest_drive = None
        for b in drive_fix_batches:
            if latest_drive is None or b.get("created_at", "") > latest_drive.get("created_at", ""):
                latest_drive = b

        if latest_migration:
            self.migration_status_label.setText(
                f"最近应用迁移记录: {latest_migration.get('id', '')} | 状态 {latest_migration.get('status', '')}"
            )
        else:
            self.migration_status_label.setText("最近应用迁移记录: (无)")

        if latest_drive:
            self.drive_status_label.setText(
                f"最近盘符修复记录: {latest_drive.get('id', '')} | 状态 {latest_drive.get('status', '')}"
            )
        else:
            self.drive_status_label.setText("最近盘符修复记录: (无)")

    def get_selected_batch_row(self):
        row = self.batch_table.currentRow()
        if row < 0 or row >= len(self.batch_rows):
            return None
        return self.batch_rows[row]

    def get_checked_batch_rows(self):
        checked = []
        for i, row in enumerate(self.batch_rows):
            item = self.batch_table.item(i, 0)
            if item is not None and item.checkState() == Qt.Checked:
                checked.append(row)
        return checked

    def update_batch_checked_count(self):
        checked = len(self.get_checked_batch_rows())
        self.batch_checked_status_label.setText(f"已勾选 {checked} 项")
        self.batch_select_all_cb.blockSignals(True)
        self.batch_select_all_cb.setChecked(checked > 0 and checked == len(self.batch_rows))
        self.batch_select_all_cb.blockSignals(False)

    def on_batch_table_item_changed(self, item):
        if self.batch_table_updating:
            return
        if item is not None and item.column() == 0:
            self.update_batch_checked_count()

    def toggle_check_all_batches(self, checked):
        self.batch_table_updating = True
        for i in range(self.batch_table.rowCount()):
            item = self.batch_table.item(i, 0)
            if item is not None:
                item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
        self.batch_table_updating = False
        self.update_batch_checked_count()

    def show_selected_batch_details(self):
        row = self.get_selected_batch_row()
        if not row:
            QMessageBox.information(self, "提示", "请先在操作记录中选择一条记录。")
            return

        text = self.build_batch_summary_text(row)

        dlg = QDialog(self)
        dlg.setWindowTitle("记录详情")
        dlg.resize(760, 420)
        layout = QVBoxLayout(dlg)
        box = QPlainTextEdit()
        box.setReadOnly(True)
        box.setPlainText(text)
        layout.addWidget(box)
        btn = QDialogButtonBox(QDialogButtonBox.Ok, parent=dlg)
        btn.accepted.connect(dlg.accept)
        layout.addWidget(btn)
        dlg.exec()

    def build_batch_summary_text(self, row):
        rec = row.get("record", {}) or {}
        text = (
            f"类型: {row.get('type', '')}\n"
            f"记录ID: {row.get('id', '')}\n"
            f"状态: {row.get('status', '')}\n"
            f"创建时间: {row.get('created_at', '')}\n"
            f"恢复时间: {row.get('restored_at', '')}\n"
            f"描述: {row.get('desc', '')}\n"
            f"备份目录: {rec.get('backup_base', '')}\n"
        )

        if row.get("kind") == "migration":
            apps = rec.get("apps", [])
            text += f"应用数量: {len(apps)}\n"
            text += f"目标根目录: {rec.get('target_root', '')}\n"
        elif row.get("kind") == "drive_fix":
            text += f"盘符变更: {rec.get('old_drive', '')} -> {rec.get('new_drive', '')}\n"
            text += f"快捷方式变更: {rec.get('shortcuts', {}).get('changed', 0)}\n"
            text += f"注册表匹配: {rec.get('registry', {}).get('matched', 0)}\n"
            text += f"环境变量变更: {rec.get('environment', {}).get('changed', 0)}\n"

        return text

    def copy_selected_batch_summary(self):
        row = self.get_selected_batch_row()
        if not row:
            QMessageBox.information(self, "提示", "请先在操作记录中选择一条记录。")
            return

        text = self.build_batch_summary_text(row)
        QApplication.clipboard().setText(text)
        self.append_log(f"[操作记录] 已复制摘要: {row.get('id', '')}")

    def show_batch_context_menu(self, pos):
        item = self.batch_table.itemAt(pos)
        if item is None:
            return

        self.batch_table.selectRow(item.row())
        row = self.get_selected_batch_row()
        if not row:
            return

        menu = QMenu(self)
        act_detail = menu.addAction("查看详情")
        act_copy = menu.addAction("复制摘要")
        act_restore = menu.addAction("执行恢复")
        act_delete = menu.addAction("删除记录")
        act_open_backup = menu.addAction("打开备份目录")
        act_restore.setEnabled(row.get("status") == "applied")

        chosen = menu.exec(self.batch_table.viewport().mapToGlobal(pos))
        if chosen == act_detail:
            self.show_selected_batch_details()
        elif chosen == act_copy:
            self.copy_selected_batch_summary()
        elif chosen == act_restore:
            if row.get("kind") == "migration":
                self.restore_batch()
            elif row.get("kind") == "drive_fix":
                self.restore_drive_fix_batch()
        elif chosen == act_delete:
            self.delete_selected_record()
        elif chosen == act_open_backup:
            backup_base = (row.get("record", {}) or {}).get("backup_base", "")
            if not backup_base or not os.path.isdir(backup_base):
                QMessageBox.warning(self, "路径不存在", "该记录备份目录不存在。")
                return
            try:
                os.startfile(backup_base)
            except Exception as e:
                QMessageBox.warning(self, "打开失败", f"无法打开目录: {e}")

    def delete_selected_record(self):
        row = self.get_selected_batch_row()
        if not row:
            QMessageBox.information(self, "提示", "请先在操作记录中选择一条记录。")
            return

        record_id = row.get("id", "")
        kind = row.get("kind", "")

        confirm = QMessageBox.question(
            self,
            "删除记录",
            f"将删除记录 {record_id}。\n选择“是”将同时删除备份目录；选择“否”仅删除记录；选择“取消”终止。",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.No,
        )
        if confirm == QMessageBox.Cancel:
            return

        delete_backup = confirm == QMessageBox.Yes
        if kind == "migration":
            ok, status = core.delete_migration_record(record_id, delete_backup=delete_backup)
        elif kind == "drive_fix":
            ok, status = core.delete_drive_fix_record(record_id, delete_backup=delete_backup)
        else:
            QMessageBox.warning(self, "删除失败", "未知记录类型。")
            return

        if not ok:
            QMessageBox.warning(self, "删除失败", f"删除失败: {status}")
            return

        self.refresh_batch_center()
        if status == "record_deleted_backup_failed":
            QMessageBox.warning(self, "删除完成", "记录已删除，但备份目录删除失败。")
        else:
            QMessageBox.information(self, "删除完成", "记录已删除。")
        self.append_log(
            f"[操作记录] 已删除记录: {record_id} | 删除备份={'是' if delete_backup else '否'}"
        )

    def delete_checked_records(self):
        rows = self.get_checked_batch_rows()
        if not rows:
            QMessageBox.information(self, "提示", "请先勾选要删除的记录。")
            return

        confirm = QMessageBox.question(
            self,
            "删除勾选记录",
            f"将删除已勾选的 {len(rows)} 条记录。\n选择“是”将同时删除备份目录；选择“否”仅删除记录；选择“取消”终止。",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.No,
        )
        if confirm == QMessageBox.Cancel:
            return

        delete_backup = confirm == QMessageBox.Yes
        ok_count = 0
        fail_count = 0
        for row in rows:
            record_id = row.get("id", "")
            kind = row.get("kind", "")
            if kind == "migration":
                ok, _ = core.delete_migration_record(record_id, delete_backup=delete_backup)
            elif kind == "drive_fix":
                ok, _ = core.delete_drive_fix_record(record_id, delete_backup=delete_backup)
            else:
                ok = False

            if ok:
                ok_count += 1
            else:
                fail_count += 1

        self.refresh_batch_center()
        QMessageBox.information(self, "删除完成", f"删除成功 {ok_count} 条，失败 {fail_count} 条。")
        self.append_log(
            f"[操作记录] 批量删除完成: 成功 {ok_count} | 失败 {fail_count} | 删除备份={'是' if delete_backup else '否'}"
        )

    def restore_checked_records(self):
        rows = self.get_checked_batch_rows()
        if not rows:
            QMessageBox.information(self, "提示", "请先勾选要恢复的记录。")
            return

        targets = [r for r in rows if r.get("status") == "applied"]
        if not targets:
            QMessageBox.information(self, "提示", "勾选记录中没有可恢复的 applied 状态项。")
            return

        confirm = QMessageBox.question(
            self,
            "恢复勾选记录",
            f"将恢复 {len(targets)} 条记录（仅 applied 状态）。是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm != QMessageBox.Yes:
            return

        payload = [{"kind": r.get("kind", ""), "id": r.get("id", "")} for r in targets]
        self.append_log(f"[操作记录] 批量恢复开始: {len(payload)} 条")
        self.run_async(self._restore_records_task, self._on_restore_checked_done, payload)

    def _restore_records_task(self, payload):
        results = []
        for item in payload:
            kind = item.get("kind", "")
            batch_id = item.get("id", "")
            if kind == "migration":
                batch, status = core.restore_migration_batch(batch_id)
            elif kind == "drive_fix":
                batch, status = core.restore_drive_fix_batch(batch_id)
            else:
                batch, status = None, "unknown_kind"
            results.append({"kind": kind, "id": batch_id, "status": status, "batch": batch})
        return results

    def _on_restore_checked_done(self, result):
        rows = result if isinstance(result, list) else []
        ok_count = sum(1 for r in rows if r.get("status") == "ok")
        fail_rows = [r for r in rows if r.get("status") != "ok"]

        self.refresh_batch_center()
        if fail_rows:
            failed_text = "\n".join(
                [f"{r.get('id', '')}: {r.get('status', '')}" for r in fail_rows[:8]]
            )
            if len(fail_rows) > 8:
                failed_text += f"\n... 另有 {len(fail_rows) - 8} 条失败"
            QMessageBox.warning(
                self,
                "批量恢复完成",
                f"恢复成功 {ok_count} 条，失败 {len(fail_rows)} 条。\n\n失败详情:\n{failed_text}",
            )
        else:
            QMessageBox.information(self, "批量恢复完成", f"恢复成功 {ok_count} 条。")

        self.append_log(f"[操作记录] 批量恢复完成: 成功 {ok_count} | 失败 {len(fail_rows)}")

    def _build_filter_box(self):
        box = QWidget()
        layout = QGridLayout(box)
        layout.setContentsMargins(0, 0, 0, 0)

        self.exclude_standard_paths_cb = QCheckBox("排除 Program Files / Program Files (x86)")
        self.exclude_standard_paths_cb.setChecked(False)
        self.exclude_standard_paths_cb.toggled.connect(self.request_checkbox_filter)

        self.exclude_users_path_cb = QCheckBox("排除 C:\\Users 下的应用")
        self.exclude_users_path_cb.setChecked(True)
        self.exclude_users_path_cb.toggled.connect(self.request_checkbox_filter)

        self.keywords_edit = QLineEdit(",".join(core.DEFAULT_EXCLUDE_VENDOR_KEYWORDS))
        self.keywords_edit.setPlaceholderText("关键词逗号分隔；留空表示不按关键词过滤")
        self.keywords_edit.textChanged.connect(self.request_apply_filters)

        layout.addWidget(QLabel("关键词黑名单:"), 0, 0)
        layout.addWidget(self.keywords_edit, 0, 1)
        layout.addWidget(self.exclude_standard_paths_cb, 1, 0, 1, 2)
        layout.addWidget(self.exclude_users_path_cb, 2, 0, 1, 2)

        return box

    def _build_action_bar(self):
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)

        self.scan_btn = QPushButton("重新扫描系统")
        self.open_batch_center_btn = QPushButton("打开操作记录")
        self.open_drive_fix_btn = QPushButton("盘符修复")
        self.cleanup_btn = QPushButton("清理待删除旧路径")
        self.migrate_btn = QPushButton("迁移已选应用")

        self.scan_btn.clicked.connect(self.scan_apps)
        self.open_batch_center_btn.clicked.connect(lambda: self.tabs.setCurrentIndex(1))
        self.open_drive_fix_btn.clicked.connect(self.open_drive_fix_dialog)
        self.cleanup_btn.clicked.connect(self.cleanup_pending)
        self.migrate_btn.clicked.connect(self.migrate_selected)

        layout.addWidget(self.scan_btn)
        layout.addWidget(self.open_batch_center_btn)
        layout.addWidget(self.open_drive_fix_btn)
        layout.addWidget(self.cleanup_btn)
        layout.addWidget(self.migrate_btn)
        layout.addStretch(1)

        return container

    def _build_table(self):
        self.app_table_model = AppTableModel(self.app_icon, self)
        self.app_proxy_model = AppFilterProxyModel(self)
        self.app_proxy_model.setSourceModel(self.app_table_model)
        self.table = QTableView()
        self.table.setModel(self.app_proxy_model)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSortIndicatorShown(True)
        self.table.horizontalHeader().setSectionsClickable(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setDefaultSectionSize(22)
        self.table.setSortingEnabled(True)
        self.table.verticalScrollBar().valueChanged.connect(self._on_table_scroll_changed)
        self.app_table_model.checkedCountChanged.connect(
            lambda _count: self.update_selected_count()
        )
        self.apply_app_table_column_layout()
        return self.table

    def apply_app_table_column_layout(self):
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        # 避免大表格下 ResizeToContents 的全量测量开销。
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        header.setSectionResizeMode(3, QHeaderView.Fixed)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.setSectionResizeMode(5, QHeaderView.Stretch)
        header.setSectionResizeMode(6, QHeaderView.Interactive)
        self.table.setColumnWidth(0, 52)
        self.table.setColumnWidth(1, 44)
        self.table.setColumnWidth(2, 300)
        self.table.setColumnWidth(3, 66)
        self.table.setColumnWidth(4, 110)
        self.table.setColumnWidth(6, 220)

    def update_selected_count(self):
        if not hasattr(self, "selection_status_label"):
            return
        selected = self.app_table_model.checked_count() if hasattr(self, "app_table_model") else 0
        self.selection_status_label.setText(f"已选 {selected} 项")

    def _table_cell_text(self, row, col):
        model = self.table.model()
        idx = model.index(row, col)
        if not idx.isValid():
            return ""
        if col == 0:
            return "√" if model.data(idx, Qt.CheckStateRole) == Qt.Checked else ""
        value = model.data(idx, Qt.DisplayRole)
        return str(value) if value is not None else ""

    def copy_selected_cells(self):
        selection_model = self.table.selectionModel()
        if selection_model is None:
            return

        selected_ranges = list(selection_model.selection())
        if not selected_ranges:
            current = self.table.currentIndex()
            if current.isValid():
                QApplication.clipboard().setText(
                    self._table_cell_text(current.row(), current.column())
                )
            return

        r = selected_ranges[0]
        rows = []
        for row in range(r.top(), r.bottom() + 1):
            cols = []
            for col in range(r.left(), r.right() + 1):
                cols.append(self._table_cell_text(row, col))
            rows.append("\t".join(cols))

        QApplication.clipboard().setText("\n".join(rows))

    def _build_target_box(self):
        box = QWidget()
        layout = QGridLayout(box)
        layout.setContentsMargins(0, 0, 0, 0)

        pf_x64, pf_x86 = preferred_migration_roots()
        self.auto_arch_cb = QCheckBox("按应用位数自动分流到 Program Files / Program Files (x86)")
        self.auto_arch_cb.setChecked(True)
        self.auto_arch_cb.toggled.connect(self.sync_target_inputs)
        self.preserve_layout_cb = QCheckBox("智能保留子目录结构（非 Program Files 路径）")
        self.preserve_layout_cb.setChecked(True)

        self.target_root_edit = QLineEdit(PathConfig.DEFAULT_TARGET_ROOT)
        self.browse_btn = QPushButton("浏览")
        self.target_root_x64_edit = QLineEdit(pf_x64)
        self.target_root_x86_edit = QLineEdit(pf_x86)
        self.browse_x64_btn = QPushButton("浏览")
        self.browse_x86_btn = QPushButton("浏览")

        self.browse_btn.clicked.connect(self.choose_target_root)
        self.browse_x64_btn.clicked.connect(self.choose_target_root_x64)
        self.browse_x86_btn.clicked.connect(self.choose_target_root_x86)

        layout.addWidget(self.auto_arch_cb, 0, 0, 1, 4)
        layout.addWidget(self.preserve_layout_cb, 1, 0, 1, 4)

        layout.addWidget(QLabel("手动目标根目录:"), 2, 0)
        layout.addWidget(self.target_root_edit, 2, 1, 1, 2)
        layout.addWidget(self.browse_btn, 2, 3)

        layout.addWidget(QLabel("x64 目标根目录:"), 3, 0)
        layout.addWidget(self.target_root_x64_edit, 3, 1, 1, 2)
        layout.addWidget(self.browse_x64_btn, 3, 3)

        layout.addWidget(QLabel("x86 目标根目录:"), 4, 0)
        layout.addWidget(self.target_root_x86_edit, 4, 1, 1, 2)
        layout.addWidget(self.browse_x86_btn, 4, 3)

        self.sync_target_inputs()

        return box

    def _build_drive_fix_box(self):
        box = QGroupBox("盘符修复")
        layout = QGridLayout(box)

        self.drive_old_edit = QLineEdit()
        self.drive_old_edit.setPlaceholderText("旧盘符，例如 E:")
        self.drive_new_edit = QLineEdit()
        self.drive_new_edit.setPlaceholderText("新盘符，例如 D:")

        self.drive_scope_label = QLabel("执行范围: 注册表 + 环境变量 + 快捷方式（固定）")

        self.drive_advanced_box = QGroupBox("高级选项")
        self.drive_advanced_box.setCheckable(True)
        self.drive_advanced_box.setChecked(False)
        adv_outer = QVBoxLayout(self.drive_advanced_box)

        self.drive_advanced_content = QWidget()
        adv_layout = QGridLayout(self.drive_advanced_content)

        self.drive_shortcut_roots_edit = QLineEdit()
        self.drive_shortcut_roots_edit.setPlaceholderText(
            "快捷方式目录(可选，分号分隔，留空使用系统默认目录)"
        )
        adv_layout.addWidget(QLabel("快捷方式目录:"), 0, 0)
        adv_layout.addWidget(self.drive_shortcut_roots_edit, 0, 1)
        adv_outer.addWidget(self.drive_advanced_content)
        self.drive_advanced_box.toggled.connect(self.drive_advanced_content.setVisible)
        self.drive_advanced_box.toggled.connect(
            lambda checked: self.drive_shortcut_roots_edit.setEnabled(
                (not self.is_busy) and checked
            )
        )
        self.drive_advanced_content.setVisible(False)

        self.drive_fix_btn = QPushButton("执行盘符修复")
        self.drive_restore_btn = QPushButton("恢复盘符修复记录")
        self.drive_fix_btn.clicked.connect(self.run_drive_fix)
        self.drive_restore_btn.clicked.connect(self.restore_drive_fix_batch)

        layout.addWidget(QLabel("旧盘符:"), 0, 0)
        layout.addWidget(self.drive_old_edit, 0, 1)
        layout.addWidget(QLabel("新盘符:"), 0, 2)
        layout.addWidget(self.drive_new_edit, 0, 3)

        layout.addWidget(self.drive_scope_label, 1, 0, 1, 4)

        layout.addWidget(self.drive_advanced_box, 2, 0, 1, 4)

        layout.addWidget(self.drive_fix_btn, 3, 2)
        layout.addWidget(self.drive_restore_btn, 3, 3)

        return box

    def _build_log_box(self):
        box = QGroupBox("日志")
        layout = QVBoxLayout(box)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)
        return box

    def _build_log_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        log_box = self._build_log_box()

        tools = QHBoxLayout()
        self.log_copy_btn = QPushButton("复制日志")
        self.log_clear_btn = QPushButton("清空日志")
        self.log_copy_btn.clicked.connect(
            lambda: QApplication.clipboard().setText(self.log.toPlainText())
        )
        self.log_clear_btn.clicked.connect(self.log.clear)
        tools.addWidget(self.log_copy_btn)
        tools.addWidget(self.log_clear_btn)
        tools.addStretch(1)

        layout.addLayout(tools)
        layout.addWidget(log_box)
        return page

    def _build_program_info_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        info_box = QPlainTextEdit()
        info_box.setReadOnly(True)
        info_box.setPlainText(
            "程序简介\n"
            "- 该工具用于迁移应用安装路径，也可以用于修复盘符变更导致的路径失效。\n\n"
            "页面与控件说明\n"
            "1) 应用迁移页\n"
            "- 关键词黑名单: 逗号分隔，按应用名/发布者/路径进行排除过滤。\n"
            "- 排除 Program Files / 排除 C:\\Users: 控制是否隐藏常见系统或用户目录下应用。\n"
            "- 快速搜索: 对当前扫描结果做即时检索（应用名、路径、发布者等）。\n"
            "- 全选 / 清空选择: 批量勾选或取消勾选应用。\n"
            "- 重新扫描系统: 重新枚举已安装应用并刷新列表。\n"
            "- 迁移已选应用: 对勾选应用执行迁移，执行前会弹出预览确认。\n"
            "- 按位数自动分流: x64 与 x86 应用分别迁移到两套目标根目录。\n"
            "- 智能保留子目录结构: 对非 Program Files 路径尽量保留原目录层级。\n"
            "- 手动目标根目录 / x64目标根目录 / x86目标根目录: 配置迁移目标路径。\n"
            "- 清理待删除旧路径: 对迁移后暂未删除成功的旧目录进行再次清理。\n\n"
            "2) 操作记录页\n"
            "- 类型筛选 / 状态筛选: 按记录类型和状态过滤批次。\n"
            "- 查看记录详情: 打开当前选中记录的完整摘要。\n"
            "- 复制记录摘要: 将记录关键信息复制到剪贴板。\n"
            "- 恢复应用迁移记录 / 恢复盘符修复记录: 回滚对应批次变更。\n"
            "- 恢复勾选记录 / 删除勾选记录: 对多条记录执行批量操作。\n"
            "- 全选: 快速勾选当前表格中的全部记录。\n\n"
            "3) 日志页\n"
            "- 日志窗口: 实时显示扫描、过滤、迁移、恢复、清理等过程输出。\n"
            "- 复制日志: 一键复制全部日志，便于问题排查与反馈。\n"
            "- 清空日志: 清空当前日志显示（不影响已执行结果和记录）。\n\n"
            "4) 菜单与盘符修复窗口\n"
            "- 工具 -> 盘符修复窗口: 打开独立盘符修复面板。\n"
            "- 旧盘符 / 新盘符: 输入要替换的盘符，如 E: -> D:。\n"
            "- 高级选项 -> 快捷方式目录: 自定义快捷方式扫描目录，留空使用默认目录。\n"
            "- 执行盘符修复: 执行注册表、环境变量、快捷方式中的盘符替换。\n"
            "- 恢复盘符修复记录: 回滚最近或指定盘符修复批次。\n"
            "- 设置 -> 恢复默认设置: 恢复界面配置，不影响历史记录和备份。\n\n"
            "安全特性\n"
            "- 执行前会生成备份，支持恢复。\n"
            "- 支持失败项清理与重启后清理计划。\n"
            "- 默认保守筛选，减少对系统组件的误操作。\n\n"
            "风险提示\n"
            "- 本程序会修改注册表、快捷方式、环境变量和系统路径信息，请确认目标路径与筛选规则。\n"
            "- 建议在执行前关闭相关应用，避免文件占用导致迁移或清理失败。\n"
            "- 部分应用包含自定义启动器/服务，迁移后仍可能需要人工检查。\n"
            "- 建议在关键变更前创建系统还原点或完整备份。\n\n"
            "建议流程\n"
            "1. 先扫描并检查过滤结果。\n"
            "2. 明确目标路径，确认迁移预览。\n"
            "3. 执行迁移并查看摘要。\n"
            "4. 在操作记录中留存/导出关键信息。\n"
            "5. 如异常，优先使用恢复功能回滚。\n"
        )
        layout.addWidget(info_box)
        return page

    def append_log(self, text):
        if not text:
            return
        self.log.appendPlainText(text.rstrip())

    def show_runtime_error(self, title, message):
        self.tabs.setCurrentWidget(self.log_tab)
        QMessageBox.critical(self, title, message + "\n\n已自动切换到日志页，请查看详细日志。")

    def focus_app_search(self):
        self.tabs.setCurrentIndex(0)
        self.app_search_edit.setFocus()
        self.app_search_edit.selectAll()

    def clear_app_search(self):
        self.app_search_edit.clear()

    def default_ui_state(self):
        pf_x64, pf_x86 = preferred_migration_roots()
        return {
            "ui_layout_version": 2,
            "current_tab": 0,
            "exclude_standard_paths": False,
            "exclude_users_path": True,
            "keywords": ",".join(core.DEFAULT_EXCLUDE_VENDOR_KEYWORDS),
            "app_search_text": "",
            "auto_arch": True,
            "preserve_layout": True,
            "workbench_details_checked": True,
            "target_root": PathConfig.DEFAULT_TARGET_ROOT,
            "target_root_x64": pf_x64,
            "target_root_x86": pf_x86,
            "drive_old": "",
            "drive_new": "",
            "drive_advanced_checked": False,
            "drive_shortcut_roots": "",
            "batch_type_filter": "全部",
            "batch_status_filter": "全部",
        }

    def load_ui_state(self):
        defaults = self.default_ui_state()
        data = core.load_json(UI_STATE_FILE, {})
        if not isinstance(data, dict):
            return defaults
        defaults.update(data)
        return defaults

    def collect_ui_state(self):
        return {
            "ui_layout_version": 2,
            "current_tab": self.tabs.currentIndex(),
            "exclude_standard_paths": self.exclude_standard_paths_cb.isChecked(),
            "exclude_users_path": self.exclude_users_path_cb.isChecked(),
            "keywords": self.keywords_edit.text(),
            "app_search_text": self.app_search_edit.text(),
            "auto_arch": self.auto_arch_cb.isChecked(),
            "preserve_layout": self.preserve_layout_cb.isChecked(),
            "workbench_details_checked": self.workbench_details_cb.isChecked(),
            "target_root": self.target_root_edit.text(),
            "target_root_x64": self.target_root_x64_edit.text(),
            "target_root_x86": self.target_root_x86_edit.text(),
            "drive_old": self.drive_old_edit.text(),
            "drive_new": self.drive_new_edit.text(),
            "drive_advanced_checked": self.drive_advanced_box.isChecked(),
            "drive_shortcut_roots": self.drive_shortcut_roots_edit.text(),
            "batch_type_filter": self.batch_type_filter.currentText(),
            "batch_status_filter": self.batch_status_filter.currentText(),
        }

    def save_ui_state(self):
        core.save_json(UI_STATE_FILE, self.collect_ui_state())

    def bind_ui_state_events(self):
        self.tabs.currentChanged.connect(lambda _=None: self.save_ui_state())
        self.exclude_standard_paths_cb.toggled.connect(lambda _=None: self.save_ui_state())
        self.exclude_users_path_cb.toggled.connect(lambda _=None: self.save_ui_state())
        self.keywords_edit.editingFinished.connect(self.save_ui_state)
        self.app_search_edit.editingFinished.connect(self.save_ui_state)

        self.auto_arch_cb.toggled.connect(lambda _=None: self.save_ui_state())
        self.preserve_layout_cb.toggled.connect(lambda _=None: self.save_ui_state())
        self.workbench_details_cb.toggled.connect(lambda _=None: self.save_ui_state())
        self.target_root_edit.editingFinished.connect(self.save_ui_state)
        self.target_root_x64_edit.editingFinished.connect(self.save_ui_state)
        self.target_root_x86_edit.editingFinished.connect(self.save_ui_state)

        self.drive_old_edit.editingFinished.connect(self.save_ui_state)
        self.drive_new_edit.editingFinished.connect(self.save_ui_state)
        self.drive_advanced_box.toggled.connect(lambda _=None: self.save_ui_state())
        self.drive_shortcut_roots_edit.editingFinished.connect(self.save_ui_state)

        self.batch_type_filter.currentIndexChanged.connect(lambda _=None: self.save_ui_state())
        self.batch_status_filter.currentIndexChanged.connect(lambda _=None: self.save_ui_state())

    def apply_ui_state(self, state):
        self.exclude_standard_paths_cb.setChecked(bool(state.get("exclude_standard_paths", False)))
        self.exclude_users_path_cb.setChecked(bool(state.get("exclude_users_path", True)))
        self.keywords_edit.setText(
            str(state.get("keywords", ",".join(core.DEFAULT_EXCLUDE_VENDOR_KEYWORDS)))
        )
        self.app_search_edit.setText(str(state.get("app_search_text", "")))

        self.auto_arch_cb.setChecked(bool(state.get("auto_arch", True)))
        self.preserve_layout_cb.setChecked(bool(state.get("preserve_layout", True)))
        details_checked = bool(
            state.get("workbench_details_checked", state.get("target_box_checked", True))
        )
        self.workbench_details_cb.setChecked(details_checked)
        self.workbench_details_panel.setVisible(details_checked)
        self.target_root_edit.setText(str(state.get("target_root", PathConfig.DEFAULT_TARGET_ROOT)))
        preferred_x64, preferred_x86 = preferred_migration_roots()
        self.target_root_x64_edit.setText(str(state.get("target_root_x64", preferred_x64)))
        self.target_root_x86_edit.setText(str(state.get("target_root_x86", preferred_x86)))

        self.drive_old_edit.setText(str(state.get("drive_old", "")))
        self.drive_new_edit.setText(str(state.get("drive_new", "")))
        self.drive_advanced_box.setChecked(bool(state.get("drive_advanced_checked", False)))
        self.drive_shortcut_roots_edit.setEnabled(self.drive_advanced_box.isChecked())
        self.drive_shortcut_roots_edit.setText(str(state.get("drive_shortcut_roots", "")))

        self.batch_type_filter.setCurrentText(str(state.get("batch_type_filter", "全部")))
        self.batch_status_filter.setCurrentText(str(state.get("batch_status_filter", "全部")))

        tab_idx = int(state.get("current_tab", 0))
        layout_version = int(state.get("ui_layout_version", 1))
        # Compatibility with old 4-tab layout (应用迁移/盘符修复/操作记录/日志)
        if layout_version < 2 and self.tabs.count() == 4:
            if tab_idx == 1:
                tab_idx = 0
            elif tab_idx == 2:
                tab_idx = 1
            elif tab_idx >= 3:
                tab_idx = 2
        if 0 <= tab_idx < self.tabs.count():
            self.tabs.setCurrentIndex(tab_idx)

    def restore_ui_defaults(self):
        ok = QMessageBox.question(
            self,
            "恢复默认设置",
            "将恢复界面配置为默认值（不影响已执行记录和备份文件）。是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ok != QMessageBox.Yes:
            return

        self.apply_ui_state(self.default_ui_state())
        self.refresh_batch_center()
        if self.raw_apps:
            self.apply_filters(log_result=True)
        self.save_ui_state()
        self.append_log("[设置] 已恢复默认配置。")

    def _set_busy(self, busy):
        self.is_busy = busy
        self.scan_btn.setEnabled(not busy)
        self.open_batch_center_btn.setEnabled(not busy)
        self.open_drive_fix_btn.setEnabled(not busy)
        self.workbench_details_cb.setEnabled(not busy)
        self.select_all_btn.setEnabled(not busy)
        self.clear_select_btn.setEnabled(not busy)
        self.app_search_edit.setEnabled(not busy)
        self.clear_search_btn.setEnabled(not busy)
        self.cleanup_btn.setEnabled(not busy)
        self.auto_arch_cb.setEnabled(not busy)
        self.preserve_layout_cb.setEnabled(not busy)
        self.migrate_btn.setEnabled(not busy)
        self.drive_old_edit.setEnabled(not busy)
        self.drive_new_edit.setEnabled(not busy)
        self.drive_scope_label.setEnabled(not busy)
        self.drive_advanced_box.setEnabled(not busy)
        self.drive_shortcut_roots_edit.setEnabled(
            (not busy) and self.drive_advanced_box.isChecked()
        )
        self.drive_fix_btn.setEnabled(not busy)
        self.drive_restore_btn.setEnabled(not busy)
        self.batch_refresh_btn.setEnabled(not busy)
        self.batch_type_filter.setEnabled(not busy)
        self.batch_status_filter.setEnabled(not busy)
        self.batch_detail_btn.setEnabled(not busy)
        self.batch_copy_btn.setEnabled(not busy)
        self.batch_delete_btn.setEnabled(not busy)
        self.batch_bulk_restore_btn.setEnabled(not busy)
        self.batch_bulk_delete_btn.setEnabled(not busy)
        self.batch_select_all_cb.setEnabled(not busy)
        self.restore_defaults_action.setEnabled(not busy)
        self.batch_restore_migration_btn.setEnabled(not busy)
        self.batch_restore_drive_btn.setEnabled(not busy)
        self.open_drive_fix_action.setEnabled(not busy)
        self.drive_restore_action.setEnabled(not busy)
        self.log_copy_btn.setEnabled(not busy)
        self.log_clear_btn.setEnabled(not busy)
        self.sync_target_inputs()

    def _keywords(self):
        raw = self.keywords_edit.text().strip()
        if raw == "":
            return []
        return [x.strip().lower() for x in raw.split(",") if x.strip()]

    def app_icon(self, app):
        icon_key = str(app.get("_icon_key", "") or "")
        if icon_key in self.icon_cache:
            return self.icon_cache[icon_key]
        if icon_key and icon_key not in self.missing_icon_paths:
            self._queue_icon_load(icon_key)
        return self.default_file_icon

    def _queue_icon_load(self, icon_key):
        if (
            not icon_key
            or icon_key in self.icon_cache
            or icon_key in self.missing_icon_paths
            or icon_key in self.pending_icon_key_set
        ):
            return
        self.pending_icon_keys.append(icon_key)
        self.pending_icon_key_set.add(icon_key)
        if not self.icon_loader_timer.isActive():
            self.icon_loader_timer.start(0)

    def _process_pending_icons(self):
        processed = 0
        while self.pending_icon_keys and processed < self.icon_load_batch_size:
            icon_key = self.pending_icon_keys.popleft()
            self.pending_icon_key_set.discard(icon_key)

            icon = QIcon(icon_key) if os.path.exists(icon_key) else QIcon()
            if icon.isNull():
                self.missing_icon_paths.add(icon_key)
            else:
                self.icon_cache[icon_key] = icon

            if hasattr(self, "app_table_model"):
                self.app_table_model.notify_icon_loaded(icon_key)
            processed += 1

        if self.pending_icon_keys:
            self.icon_loader_timer.start(0)

    def _prime_visible_icon_loads(self):
        rows = min(160, self.app_proxy_model.rowCount())
        for row in range(rows):
            proxy_idx = self.app_proxy_model.index(row, 1)
            source_idx = self.app_proxy_model.mapToSource(proxy_idx)
            app = self.app_table_model.app_at_row(source_idx.row())
            if app:
                self._queue_icon_load(str(app.get("_icon_key", "") or ""))

    def _on_table_scroll_changed(self, _value):
        # 仅在数据量较大时进行滚动预热，避免小列表下额外开销。
        total = self.app_proxy_model.rowCount()
        if total <= 120:
            return

        top_row = self.table.rowAt(0)
        if top_row < 0:
            top_row = 0
        bottom_row = self.table.rowAt(self.table.viewport().height() - 1)
        if bottom_row < 0:
            bottom_row = min(top_row + 40, total - 1)

        preload_start = max(0, top_row - 40)
        preload_end = min(total - 1, bottom_row + 120)

        for row in range(preload_start, preload_end + 1):
            proxy_idx = self.app_proxy_model.index(row, 1)
            source_idx = self.app_proxy_model.mapToSource(proxy_idx)
            app = self.app_table_model.app_at_row(source_idx.row())
            if app:
                self._queue_icon_load(str(app.get("_icon_key", "") or ""))

    def find_install_dir_icon_path(self, app):
        install_dir = str(app.get("install_dir", "") or "").strip()
        if not install_dir:
            return ""

        if install_dir in self.install_dir_icon_cache:
            return self.install_dir_icon_cache[install_dir]

        if not os.path.isdir(install_dir):
            self.install_dir_icon_cache[install_dir] = ""
            return ""

        try:
            icon_files = []
            for name in os.listdir(install_dir):
                if not name.lower().endswith(".ico"):
                    continue
                full = os.path.join(install_dir, name)
                if not os.path.isfile(full):
                    continue
                icon_files.append(full)
        except OSError:
            self.install_dir_icon_cache[install_dir] = ""
            return ""

        if not icon_files:
            self.install_dir_icon_cache[install_dir] = ""
            return ""

        icon_files.sort(
            key=lambda p: (
                "icon" not in os.path.basename(p).lower(),
                os.path.basename(p).lower(),
            )
        )
        best = icon_files[0]

        self.install_dir_icon_cache[install_dir] = best
        return best

    def request_apply_filters(self, _=None):
        """关键词输入过滤请求。"""
        if not self.raw_apps:
            return
        self.request_search_filter()

    def _current_filter_payload(self, log_result=True):
        return {
            "exclude_standard": self.exclude_standard_paths_cb.isChecked(),
            "exclude_users_path": self.exclude_users_path_cb.isChecked(),
            "keywords": self._keywords(),
            "search_text": self.app_search_edit.text().strip().lower(),
            "log_result": bool(log_result),
        }

    def _dispatch_filter_request(self, payload):
        if not payload:
            return

        t0 = time.perf_counter()

        self.app_table_model.clear_checks()
        self.app_proxy_model.set_filter_options(
            payload["exclude_standard"],
            payload["exclude_users_path"],
            payload["keywords"],
            payload["search_text"],
        )

        visible_count = self.app_proxy_model.rowCount()
        elapsed_ms = (time.perf_counter() - t0) * 1000

        self.update_selected_count()

        if payload.get("log_result"):
            self.append_log(
                f"[过滤] 关键词={payload['keywords'] if payload['keywords'] else '(无)'} | 搜索={payload['search_text'] if payload['search_text'] else '(无)'} | 排除ProgramFiles={'开' if payload['exclude_standard'] else '关'} | 排除Users={'开' if payload['exclude_users_path'] else '关'} | 结果={visible_count}/{len(self.raw_apps)} | 【性能】代理过滤+刷新={elapsed_ms:.1f}ms"
            )

    def _execute_pending_checkbox_filter(self):
        payload = self._pending_checkbox_filter
        self._pending_checkbox_filter = None
        self._dispatch_filter_request(payload)

    def _execute_pending_search_filter(self):
        payload = self._pending_search_filter
        self._pending_search_filter = None
        self._dispatch_filter_request(payload)

    def request_checkbox_filter(self, _=None):
        """处理勾选类过滤项，使用短防抖。"""
        if not self.raw_apps:
            return

        # 短防抖合并快速连击，避免重复渲染。
        self.checkbox_filter_timer.stop()
        self._pending_checkbox_filter = self._current_filter_payload(log_result=True)
        self.checkbox_filter_timer.start()

    def request_search_filter(self, _=None):
        """处理搜索框输入，使用长防抖。"""
        if not self.raw_apps:
            return

        # 重置搜索防抖（300ms）
        self.search_filter_timer.stop()
        self._pending_search_filter = self._current_filter_payload(log_result=True)
        self.search_filter_timer.start()

    def run_async(self, fn, on_ok, *args, **kwargs):
        worker = Worker(fn, *args, **kwargs)
        self.active_workers.add(worker)

        def _release_worker():
            if worker in self.active_workers:
                self.active_workers.remove(worker)

        def _ok(result, stdout_text):
            _release_worker()
            self._set_busy(False)
            self.append_log(stdout_text)
            try:
                on_ok(result)
            except Exception:
                self.append_log(traceback.format_exc())
                self.show_runtime_error("回调失败", "任务已完成，但界面更新失败。")

        def _fail(err_text, stdout_text):
            _release_worker()
            self._set_busy(False)
            self.append_log(stdout_text)
            self.append_log(err_text)
            self.show_runtime_error("执行失败", "任务执行失败。")

        worker.signals.finished.connect(_ok)
        worker.signals.failed.connect(_fail)
        self._set_busy(True)
        self.thread_pool.start(worker)

    def scan_apps(self):
        if self.is_busy:
            return
        self.append_log("[扫描] 开始扫描系统应用列表...")

        self.run_async(
            core.enum_installed_apps,
            self._on_scan_done,
            exclude_standard_paths=False,
            exclude_keywords=[],
        )

    def _on_scan_done(self, apps):
        self.raw_apps = apps or []

        users_root_norm = os.path.normcase("C:\\Users\\")
        for app in self.raw_apps:
            install_dir = str(app.get("install_dir", "") or "")
            install_dir_norm = os.path.normcase(install_dir)
            reg_subkey = str(app.get("reg_subkey", "") or "")
            subkey_leaf = reg_subkey.rsplit("\\", 1)[-1] if reg_subkey else ""

            app["_is_standard_path"] = core.is_in_standard_install_path(install_dir)
            app["_is_users_path"] = install_dir_norm.startswith(users_root_norm)
            app["_search_blob"] = " ".join(
                [
                    str(app.get("display_name", "")),
                    install_dir,
                    str(app.get("publisher", "")),
                    str(app.get("arch", "")),
                    str(app.get("install_date", "")),
                ]
            ).lower()
            app["_keyword_blob"] = " ".join(
                [
                    str(app.get("display_name", "")),
                    str(app.get("publisher", "")),
                    install_dir,
                    subkey_leaf,
                ]
            ).lower()
            icon_path = str(app.get("display_icon", "") or "").strip()
            app["_icon_key"] = os.path.normcase(os.path.normpath(icon_path)) if icon_path else ""

        self.app_table_model.set_apps(self.raw_apps, preserve_checks=False)
        self.app_proxy_model.invalidateFilter()

        self.append_log(f"[扫描] 完成，基线应用 {len(self.raw_apps)} 项。")
        self.apply_filters(log_result=True)

    def apply_filters(self, _=None, log_result=False):
        """兼容入口：应用当前过滤条件到代理模型。"""
        if not self.raw_apps:
            self.app_table_model.set_apps([])
            self.update_selected_count()
            return
        self._dispatch_filter_request(self._current_filter_payload(log_result=log_result))
        self._prime_visible_icon_loads()

    def select_all(self):
        checked_ids = set()
        for row in range(self.app_proxy_model.rowCount()):
            source_idx = self.app_proxy_model.mapToSource(self.app_proxy_model.index(row, 0))
            app = self.app_table_model.app_at_row(source_idx.row())
            if app:
                checked_ids.add(self.app_table_model.app_id(app))
        self.app_table_model.set_checked_ids(checked_ids)
        self.update_selected_count()

    def clear_selection(self):
        self.app_table_model.clear_checks()
        self.update_selected_count()

    def choose_target_root(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择目标根目录",
            self.target_root_edit.text().strip() or PathConfig.DEFAULT_TARGET_ROOT,
        )
        if folder:
            self.target_root_edit.setText(folder)

    def choose_target_root_x64(self):
        preferred_x64, _ = preferred_migration_roots()
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择 x64 目标根目录",
            self.target_root_x64_edit.text().strip() or preferred_x64,
        )
        if folder:
            self.target_root_x64_edit.setText(folder)

    def choose_target_root_x86(self):
        _, preferred_x86 = preferred_migration_roots()
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择 x86 目标根目录",
            self.target_root_x86_edit.text().strip() or preferred_x86,
        )
        if folder:
            self.target_root_x86_edit.setText(folder)

    def sync_target_inputs(self):
        auto_arch = self.auto_arch_cb.isChecked()
        can_edit = not self.is_busy
        self.target_root_edit.setEnabled((not auto_arch) and can_edit)
        self.browse_btn.setEnabled((not auto_arch) and can_edit)
        self.target_root_x64_edit.setEnabled(auto_arch and can_edit)
        self.target_root_x86_edit.setEnabled(auto_arch and can_edit)
        self.browse_x64_btn.setEnabled(auto_arch and can_edit)
        self.browse_x86_btn.setEnabled(auto_arch and can_edit)

    def normalize_root_input(self, raw, fallback):
        text = (raw or "").strip()
        if not text:
            text = fallback
        text = os.path.expanduser(os.path.expandvars(text))
        if len(text) == 2 and text[1] == ":" and text[0].isalpha():
            text = text + "\\"
        if not os.path.isabs(text):
            return ""
        return os.path.abspath(text)

    def show_migration_preview_dialog(self, plan_items):
        dlg = QDialog(self)
        dlg.setWindowTitle("迁移预览与路径确认")
        dlg.resize(1050, 620)

        layout = QVBoxLayout(dlg)
        tip = QLabel("请确认每个应用的目标路径；可直接编辑“目标路径”列。")
        layout.addWidget(tip)

        table = QTableWidget(len(plan_items), 4)
        table.setHorizontalHeaderLabels(["应用", "位数", "源路径", "目标路径(可编辑)"])
        preview_header = table.horizontalHeader()
        preview_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        preview_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        preview_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        preview_header.setSectionResizeMode(3, QHeaderView.Stretch)

        for i, item in enumerate(plan_items):
            table.setItem(i, 0, QTableWidgetItem(item.get("name", "")))
            table.setItem(i, 1, QTableWidgetItem(item.get("detected_arch", "")))
            table.setItem(i, 2, QTableWidgetItem(item.get("src", "")))

            dst_item = QTableWidgetItem(item.get("dst", ""))
            dst_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            table.setItem(i, 3, dst_item)

        layout.addWidget(table, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=dlg)
        layout.addWidget(buttons)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)

        if dlg.exec() != QDialog.Accepted:
            return None

        edited = []
        for i, item in enumerate(plan_items):
            dst_item = table.item(i, 3)
            edited_dst = (dst_item.text() if dst_item is not None else item.get("dst", "")).strip()
            if len(edited_dst) == 2 and edited_dst[1] == ":" and edited_dst[0].isalpha():
                edited_dst = edited_dst + "\\"
            if not edited_dst:
                QMessageBox.warning(
                    self, "路径无效", f"应用 {item.get('name', '')} 的目标路径为空。"
                )
                return None
            if not os.path.isabs(edited_dst):
                QMessageBox.warning(
                    self,
                    "路径无效",
                    f"应用 {item.get('name', '')} 的目标路径必须是绝对路径。",
                )
                return None
            new_item = dict(item)
            new_item["dst"] = os.path.abspath(edited_dst)
            edited.append(new_item)

        return edited

    def show_migration_summary_dialog(self, batch):
        apps = batch.get("apps", []) if isinstance(batch, dict) else []
        copy_ok = sum(1 for a in apps if a.get("copy") == "ok")
        copy_fail = sum(1 for a in apps if a.get("copy") == "failed")
        reg_changed = sum(int(a.get("registry", {}).get("changed", 0) or 0) for a in apps)
        env_changed = sum(int(a.get("environment", {}).get("changed", 0) or 0) for a in apps)
        svc_changed = sum(int(a.get("services", {}).get("changed", 0) or 0) for a in apps)
        task_changed = sum(int(a.get("tasks", {}).get("changed", 0) or 0) for a in apps)
        sc_changed = sum(int(a.get("shortcuts", {}).get("changed", 0) or 0) for a in apps)
        del_ok = sum(1 for a in apps if a.get("delete_old", {}).get("success"))

        text = (
            f"记录ID: {batch.get('id', '')}\n"
            f"目标根目录: {batch.get('target_root', '')}\n"
            f"应用总数: {len(apps)}\n"
            f"复制成功/失败: {copy_ok}/{copy_fail}\n"
            f"注册表变更: {reg_changed}\n"
            f"环境变量变更: {env_changed}\n"
            f"服务项变更: {svc_changed}\n"
            f"计划任务变更: {task_changed}\n"
            f"快捷方式变更: {sc_changed}\n"
            f"旧路径删除成功: {del_ok}\n"
            f"备份目录: {batch.get('backup_base', '')}"
        )

        dlg = QDialog(self)
        dlg.setWindowTitle("迁移摘要")
        dlg.resize(760, 420)
        layout = QVBoxLayout(dlg)
        box = QPlainTextEdit()
        box.setReadOnly(True)
        box.setPlainText(text)
        layout.addWidget(box)
        btn = QDialogButtonBox(QDialogButtonBox.Ok, parent=dlg)
        btn.accepted.connect(dlg.accept)
        layout.addWidget(btn)
        dlg.exec()

    def show_restore_summary_dialog(self, batch):
        rr = (batch or {}).get("restore_result", {})
        text = (
            f"记录ID: {(batch or {}).get('id', '')}\n"
            f"程序路径恢复 成功/失败: {rr.get('program_path_ok', 0)}/{rr.get('program_path_fail', 0)}\n"
            f"注册表恢复 成功/失败: {rr.get('registry_ok', 0)}/{rr.get('registry_fail', 0)}\n"
            f"环境变量恢复 成功/失败: {rr.get('environment_ok', 0)}/{rr.get('environment_fail', 0)}\n"
            f"服务项恢复 成功/失败: {rr.get('services_ok', 0)}/{rr.get('services_fail', 0)}\n"
            f"计划任务恢复 成功/失败: {rr.get('tasks_ok', 0)}/{rr.get('tasks_fail', 0)}\n"
            f"快捷方式恢复 成功/失败: {rr.get('shortcuts_ok', 0)}/{rr.get('shortcuts_fail', 0)}\n"
            f"恢复时间: {(batch or {}).get('restored_at', '')}"
        )
        QMessageBox.information(self, "恢复摘要", text)

    def selected_apps(self):
        if not hasattr(self, "app_table_model"):
            return []
        return self.app_table_model.selected_apps()

    def migrate_selected(self):
        selected = self.selected_apps()
        if not selected:
            QMessageBox.information(self, "提示", "请先在列表中勾选要迁移的应用。")
            return

        auto_arch = self.auto_arch_cb.isChecked()
        preserve_layout = self.preserve_layout_cb.isChecked()
        pf_x64, pf_x86 = core.get_program_files_roots()

        if auto_arch:
            target_root_x64 = self.normalize_root_input(self.target_root_x64_edit.text(), pf_x64)
            target_root_x86 = self.normalize_root_input(self.target_root_x86_edit.text(), pf_x86)
            if not target_root_x64 or not target_root_x86:
                QMessageBox.warning(
                    self,
                    "路径无效",
                    "x64/x86 目标根目录必须是绝对路径，例如 D:\\Program Files",
                )
                return
            core.ensure_dir(target_root_x64)
            core.ensure_dir(target_root_x86)
            target_root = target_root_x64
        else:
            target_root = self.normalize_root_input(
                self.target_root_edit.text(), PathConfig.DEFAULT_TARGET_ROOT
            )
            if not target_root:
                QMessageBox.warning(
                    self,
                    "路径无效",
                    "手动目标根目录必须是绝对路径，例如 D:\\Program Files",
                )
                return
            core.ensure_dir(target_root)
            target_root_x64 = None
            target_root_x86 = None

        plan = core.build_migration_preview(
            selected,
            target_root,
            auto_arch=auto_arch,
            target_root_x64=target_root_x64,
            target_root_x86=target_root_x86,
            preserve_relative_layout=preserve_layout,
        )
        edited_plan = self.show_migration_preview_dialog(plan)
        if edited_plan is None:
            return

        destination_overrides = {
            core.path_norm(item.get("src", "")): item.get("dst", "") for item in edited_plan
        }

        if auto_arch:
            self.append_log(f"[迁移] 开始，自动分流：x64={target_root_x64} | x86={target_root_x86}")
        else:
            self.append_log(f"[迁移] 开始，目标目录: {target_root}")

        self.run_async(
            core.migrate_selected_apps,
            self._on_migrate_done,
            selected,
            target_root,
            auto_arch=auto_arch,
            target_root_x64=target_root_x64,
            target_root_x86=target_root_x86,
            preserve_relative_layout=preserve_layout,
            destination_overrides=destination_overrides,
        )

    def _on_migrate_done(self, _):
        self.append_log("[迁移] 完成。")
        self.refresh_batch_center()
        if isinstance(_, dict):
            self.show_migration_summary_dialog(_)
            failed_delete_items = []
            for app in _.get("apps", []):
                d = app.get("delete_old", {})
                if not d.get("success"):
                    failed_delete_items.append(
                        {
                            "path": app.get("src", ""),
                            "reason": d.get("reason", ""),
                            "batch_id": _.get("id", ""),
                            "app_name": app.get("name", ""),
                        }
                    )
            if failed_delete_items:
                self.offer_reboot_cleanup_for_failed_items(
                    failed_delete_items,
                    prompt_title="迁移后清理失败",
                )
        else:
            QMessageBox.information(self, "完成", "迁移执行完成，请查看日志。")

    def restore_batch(self):
        selected_row = self.get_selected_batch_row()
        if selected_row and selected_row.get("kind") == "migration":
            if selected_row.get("status") != "applied":
                QMessageBox.information(self, "提示", "所选应用迁移记录不可恢复（可能已恢复）。")
                return

            batch_id = selected_row.get("id", "")
            confirm = QMessageBox.question(
                self,
                "确认恢复",
                f"将恢复记录 {batch_id}。\n此操作会尝试回退程序目录/注册表/环境变量/快捷方式。\n是否继续？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if confirm != QMessageBox.Yes:
                return

            self.append_log(f"[恢复] 开始恢复记录: {batch_id}")
            self.run_async(core.restore_migration_batch, self._on_restore_done, batch_id)
            return

        batches = core.list_batches(applied_only=True)
        if not batches:
            QMessageBox.information(self, "提示", "没有可恢复的已执行记录。")
            return

        options = [f"{b.get('id', '')} | {b.get('created_at', '')}" for b in batches]
        picked, ok = QInputDialog.getItem(
            self, "选择恢复记录", "请选择要恢复的记录：", options, 0, False
        )
        if not ok or not picked:
            return

        batch_id = picked.split("|", 1)[0].strip()
        confirm = QMessageBox.question(
            self,
            "确认恢复",
            f"将恢复记录 {batch_id}。\n此操作会尝试回退程序目录/注册表/环境变量/快捷方式。\n是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm != QMessageBox.Yes:
            return

        self.append_log(f"[恢复] 开始恢复记录: {batch_id}")
        self.run_async(core.restore_migration_batch, self._on_restore_done, batch_id)

    def _on_restore_done(self, result):
        batch, status = result if isinstance(result, tuple) else (None, "unknown")
        if status != "ok" or not batch:
            QMessageBox.warning(self, "恢复失败", f"恢复失败: {status}")
            return
        self.append_log(f"[恢复] 完成: {batch.get('id', '')}")
        self.refresh_batch_center()
        self.show_restore_summary_dialog(batch)

    def cleanup_pending(self):
        ok = QMessageBox.question(
            self,
            "确认清理",
            "将尝试删除待清理列表中的旧路径。是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ok != QMessageBox.Yes:
            return

        self.append_log("[清理] 开始处理待删除旧路径...")
        self.run_async(core.perform_cleanup_pending, self._on_cleanup_done)

    def _on_cleanup_done(self, result):
        self.append_log("[清理] 完成。")
        if isinstance(result, dict) and int(result.get("fail", 0)) > 0:
            failed_items = result.get("failed_items", [])
            self.offer_reboot_cleanup_for_failed_items(
                failed_items,
                prompt_title="清理存在失败",
            )
        self.save_ui_state()

    def offer_reboot_cleanup_for_failed_items(self, failed_items, prompt_title):
        failed_items = failed_items or []
        if not failed_items:
            return

        lines = []
        for i, item in enumerate(failed_items[:5], 1):
            lines.append(f"{i}. {item.get('path', '')}\n原因: {item.get('reason', '')}")
        more = ""
        if len(failed_items) > 5:
            more = f"\n... 另有 {len(failed_items) - 5} 条失败记录，请查看日志页。"

        QMessageBox.warning(
            self,
            prompt_title,
            "以下路径清理失败：\n\n" + "\n\n".join(lines) + more,
        )

        ask_schedule = QMessageBox.question(
            self,
            "计划重启后清理",
            "是否注册一次“开机自动清理”任务，并在重启后再次尝试删除这些路径？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if ask_schedule != QMessageBox.Yes:
            return

        paths = [item.get("path", "") for item in failed_items if item.get("path")]
        ok, status, detail = core.schedule_cleanup_on_reboot(paths)
        if not ok:
            QMessageBox.warning(self, "注册失败", f"注册开机清理任务失败: {status}")
            return

        task_name = detail.get("task_name", "")
        self.append_log(f"[清理] 已注册开机自动清理任务: {task_name}")
        ask_reboot = QMessageBox.question(
            self,
            "立即重启",
            "清理任务已注册。是否现在重启系统以执行清理？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ask_reboot == QMessageBox.Yes:
            try:
                subprocess.run(
                    ["shutdown", "/r", "/t", "5"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=True,
                )
                QMessageBox.information(self, "即将重启", "系统将在 5 秒后重启。")
            except subprocess.CalledProcessError:
                QMessageBox.warning(self, "重启失败", "无法发起重启，请手动重启系统。")

    def run_drive_fix(self):
        old_drive = core.normalize_drive(self.drive_old_edit.text().strip())
        new_drive = core.normalize_drive(self.drive_new_edit.text().strip())
        if not old_drive or not new_drive:
            QMessageBox.warning(self, "输入无效", "请输入合法盘符，例如 E: 和 D:")
            return
        if old_drive == new_drive:
            QMessageBox.warning(self, "输入无效", "原盘符和新盘符不能相同。")
            return

        shortcut_roots = None
        raw_roots = self.drive_shortcut_roots_edit.text().strip()
        if raw_roots:
            shortcut_roots = [os.path.abspath(x.strip()) for x in raw_roots.split(";") if x.strip()]

        confirm = QMessageBox.question(
            self,
            "确认盘符修复",
            f"将执行盘符修复: {old_drive} -> {new_drive}\n"
            "执行范围: 注册表 + 环境变量 + 快捷方式\n"
            "是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm != QMessageBox.Yes:
            return

        self.append_log(f"[盘符修复] 开始: {old_drive} -> {new_drive}")
        self.run_async(
            core.run_drive_letter_fix,
            self._on_drive_fix_done,
            old_drive,
            new_drive,
            include_registry=True,
            include_environment=True,
            include_shortcuts=True,
            shortcut_roots=shortcut_roots,
        )

    def _on_drive_fix_done(self, batch):
        if not isinstance(batch, dict):
            QMessageBox.information(self, "完成", "盘符修复执行完成，请查看日志。")
            return

        self.append_log(f"[盘符修复] 完成: {batch.get('id', '')}")
        self.refresh_batch_center()
        sc = batch.get("shortcuts", {})
        rg = batch.get("registry", {})
        env = batch.get("environment", {})
        text = (
            f"记录ID: {batch.get('id', '')}\n"
            f"快捷方式: 扫描 {sc.get('scanned', 0)} | 变更 {sc.get('changed', 0)} | 失败 {sc.get('failed', 0)}\n"
            f"注册表: 匹配 {rg.get('matched', 0)} | 成功 {rg.get('applied_success', 0)} | 失败 {rg.get('applied_failed', 0)}\n"
            f"环境变量: 变更 {env.get('changed', 0)} | 失败 {env.get('failed', 0)}\n"
            f"备份目录: {batch.get('backup_base', '')}"
        )
        QMessageBox.information(self, "盘符修复摘要", text)
        self.save_ui_state()

    def restore_drive_fix_batch(self):
        selected_row = self.get_selected_batch_row()
        if selected_row and selected_row.get("kind") == "drive_fix":
            if selected_row.get("status") != "applied":
                QMessageBox.information(self, "提示", "所选盘符修复记录不可恢复（可能已恢复）。")
                return

            batch_id = selected_row.get("id", "")
            confirm = QMessageBox.question(
                self,
                "确认恢复",
                f"将恢复盘符修复记录 {batch_id}。是否继续？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if confirm != QMessageBox.Yes:
                return

            self.append_log(f"[盘符恢复] 开始恢复记录: {batch_id}")
            self.run_async(core.restore_drive_fix_batch, self._on_drive_restore_done, batch_id)
            return

        batches = core.list_drive_fix_batches(applied_only=True)
        if not batches:
            QMessageBox.information(self, "提示", "没有可恢复的盘符修复记录。")
            return

        options = [
            f"{b.get('id', '')} | {b.get('created_at', '')} | {b.get('old_drive', '')}->{b.get('new_drive', '')}"
            for b in batches
        ]
        picked, ok = QInputDialog.getItem(
            self, "选择恢复记录", "请选择要恢复的盘符修复记录：", options, 0, False
        )
        if not ok or not picked:
            return

        batch_id = picked.split("|", 1)[0].strip()
        confirm = QMessageBox.question(
            self,
            "确认恢复",
            f"将恢复盘符修复记录 {batch_id}。是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm != QMessageBox.Yes:
            return

        self.append_log(f"[盘符恢复] 开始恢复记录: {batch_id}")
        self.run_async(core.restore_drive_fix_batch, self._on_drive_restore_done, batch_id)

    def _on_drive_restore_done(self, result):
        batch, status = result if isinstance(result, tuple) else (None, "unknown")
        if status != "ok" or not batch:
            QMessageBox.warning(self, "恢复失败", f"恢复失败: {status}")
            return

        rr = batch.get("restore_result", {})
        self.append_log(f"[盘符恢复] 完成: {batch.get('id', '')}")
        self.refresh_batch_center()
        QMessageBox.information(
            self,
            "盘符恢复摘要",
            f"记录ID: {batch.get('id', '')}\n"
            f"快捷方式恢复: 成功 {rr.get('shortcut_success', 0)} | 失败 {rr.get('shortcut_failed', 0)}\n"
            f"注册表恢复: 成功 {rr.get('registry_success', 0)} | 失败 {rr.get('registry_failed', 0)}\n"
            f"环境变量恢复: 成功 {rr.get('environment_success', 0)} | 失败 {rr.get('environment_failed', 0)}",
        )
        self.save_ui_state()

    def closeEvent(self, event):
        self.save_ui_state()
        super().closeEvent(event)


def main():
    core.ensure_admin_or_exit()

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
