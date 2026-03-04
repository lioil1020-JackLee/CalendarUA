from __future__ import annotations

import csv
from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt, QEvent
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QMenu,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from database.sqlite_manager import SQLiteManager
from ui.combo_wheel_helper import attach_combo_wheel_behavior


class HolidaySettingsDialog(QDialog):
    """假日設定對話框。"""

    WEEKDAY_LABELS = [
        (1, "週一"),
        (2, "週二"),
        (3, "週三"),
        (4, "週四"),
        (5, "週五"),
        (6, "週六"),
        (7, "週日"),
    ]

    def __init__(self, db_manager: SQLiteManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.selected_rule_id: Optional[int] = None
        self._loading_data = False

        self.setWindowTitle("假日設定")
        self.setModal(True)
        self.resize(520, 460)

        self._build_ui()
        self._load_data()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        import_export_layout = QHBoxLayout()
        self.btn_import = QPushButton("匯入假日 (csv)")
        self.btn_export = QPushButton("匯出假日 (csv)")
        self.btn_import.setStyleSheet("border: 1px solid #7a7a7a;")
        self.btn_export.setStyleSheet("border: 1px solid #7a7a7a;")
        import_export_layout.addWidget(self.btn_import)
        import_export_layout.addWidget(self.btn_export)
        import_export_layout.addStretch()
        root.addLayout(import_export_layout)

        weekday_group = QGroupBox("每週假日")
        weekday_layout = QHBoxLayout(weekday_group)
        weekday_layout.setSpacing(16)
        self.weekday_checks: Dict[int, QCheckBox] = {}
        for weekday, label in self.WEEKDAY_LABELS:
            cb = QCheckBox(label)
            self.weekday_checks[weekday] = cb
            weekday_layout.addWidget(cb)
        weekday_layout.addStretch()
        root.addWidget(weekday_group)

        date_group = QGroupBox("日期假日（可新增 / 編輯 / 刪除）")
        date_layout = QVBoxLayout(date_group)

        self.table_rules = QTableWidget(0, 3)
        self.table_rules.setHorizontalHeaderLabels(["ID", "曆法", "日期"])
        self.table_rules.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_rules.setSelectionMode(QTableWidget.SingleSelection)
        self.table_rules.verticalHeader().setVisible(False)
        self.table_rules.setColumnHidden(0, True)
        self.table_rules.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_rules.horizontalHeader().setStretchLastSection(True)
        self.table_rules.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_rules.setToolTip("右鍵可新增、編輯、刪除日期假日")
        date_layout.addWidget(self.table_rules)

        root.addWidget(date_group)

        bottom = QHBoxLayout()
        bottom.addStretch()
        self.btn_close = QPushButton("關閉")
        bottom.addWidget(self.btn_close)
        root.addLayout(bottom)

        for cb in self.weekday_checks.values():
            cb.toggled.connect(self._save_weekdays_auto)
        self.table_rules.itemSelectionChanged.connect(self._on_rule_selected)
        self.table_rules.customContextMenuRequested.connect(self._show_rule_context_menu)
        self.table_rules.cellDoubleClicked.connect(self._on_rule_double_clicked)
        self.btn_import.clicked.connect(self._import_csv)
        self.btn_export.clicked.connect(self._export_csv)
        self.btn_close.clicked.connect(self.accept)

    def _load_data(self) -> None:
        self._loading_data = True
        payload = self.db_manager.get_holiday_rules_payload()

        weekdays = set(int(x) for x in payload.get("weekdays", []))
        for weekday, cb in self.weekday_checks.items():
            cb.blockSignals(True)
            cb.setChecked(weekday in weekdays)
            cb.blockSignals(False)

        self.table_rules.setRowCount(0)
        sorted_dates = sorted(
            payload.get("dates", []),
            key=lambda rule: (
                0 if str(rule.get("calendar_type", "")).strip().lower() == "solar" else 1,
                int(rule.get("month", 0) or 0),
                int(rule.get("day", 0) or 0),
            ),
        )
        for rule in sorted_dates:
            self._append_rule_to_table(rule)

        self.selected_rule_id = None
        self._loading_data = False

    def _append_rule_to_table(self, rule: Dict[str, Any]) -> None:
        row = self.table_rules.rowCount()
        self.table_rules.insertRow(row)

        rule_id = int(rule.get("id", 0) or 0)
        calendar_type = str(rule.get("calendar_type", "solar") or "solar")
        month = int(rule.get("month", 1) or 1)
        day = int(rule.get("day", 1) or 1)

        cal_text = "國曆" if calendar_type == "solar" else "農曆"
        date_text = f"{month}/{day}"

        self.table_rules.setItem(row, 0, QTableWidgetItem(str(rule_id)))
        self.table_rules.setItem(row, 1, QTableWidgetItem(cal_text))
        self.table_rules.setItem(row, 2, QTableWidgetItem(date_text))

    def _save_weekdays_auto(self) -> None:
        if self._loading_data:
            return
        weekdays = [weekday for weekday, cb in self.weekday_checks.items() if cb.isChecked()]
        ok = self.db_manager.set_weekday_holidays(weekdays)
        if not ok:
            QMessageBox.warning(self, "失敗", "儲存週別假日失敗。")
        self._load_data()

    def _add_rule_from_popup(self) -> None:
        dialog = HolidayRuleEditDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return

        values = dialog.get_values()
        new_id = self.db_manager.add_holiday_rule(
            values["calendar_type"],
            values["month"],
            values["day"],
        )
        if not new_id:
            QMessageBox.warning(self, "失敗", "新增失敗（可能為重複日期）。")
            return
        self._load_data()

    def _update_rule_from_popup(self) -> None:
        if not self.selected_rule_id:
            QMessageBox.information(self, "提示", "請先選擇要編輯的日期。")
            return

        selected = self._get_selected_rule()
        if not selected:
            QMessageBox.information(self, "提示", "請先選擇要編輯的日期。")
            return

        dialog = HolidayRuleEditDialog(
            self,
            calendar_type=selected["calendar_type"],
            month=selected["month"],
            day=selected["day"],
        )
        if dialog.exec() != QDialog.Accepted:
            return

        values = dialog.get_values()
        ok = self.db_manager.update_holiday_rule(
            self.selected_rule_id,
            values["calendar_type"],
            values["month"],
            values["day"],
        )
        if not ok:
            QMessageBox.warning(self, "失敗", "編輯失敗（可能為重複日期）。")
            return
        self._load_data()

    def _delete_rule(self) -> None:
        if not self.selected_rule_id:
            QMessageBox.information(self, "提示", "請先選擇要刪除的日期。")
            return

        ok = self.db_manager.delete_holiday_entry(self.selected_rule_id)
        if not ok:
            QMessageBox.warning(self, "失敗", "刪除失敗。")
            return

        self._load_data()

    def _show_rule_context_menu(self, position) -> None:
        menu = QMenu(self)
        action_add = QAction("新增", self)
        action_edit = QAction("編輯", self)
        action_delete = QAction("刪除", self)

        if self.selected_rule_id is None:
            action_edit.setEnabled(False)
            action_delete.setEnabled(False)

        menu.addAction(action_add)
        menu.addAction(action_edit)
        menu.addAction(action_delete)

        selected_action = menu.exec(self.table_rules.viewport().mapToGlobal(position))
        if selected_action == action_add:
            self._add_rule_from_popup()
        elif selected_action == action_edit:
            self._update_rule_from_popup()
        elif selected_action == action_delete:
            self._delete_rule()

    def _on_rule_double_clicked(self, row: int, _column: int) -> None:
        if row < 0:
            return
        self.table_rules.selectRow(row)
        self._on_rule_selected()
        self._update_rule_from_popup()

    def _get_selected_rule(self) -> Optional[Dict[str, Any]]:
        if self.selected_rule_id is None:
            return None
        row = self.table_rules.currentRow()
        if row < 0:
            return None

        item_calendar = self.table_rules.item(row, 1)
        item_date = self.table_rules.item(row, 2)
        if not item_calendar or not item_date:
            return None

        try:
            month_str, day_str = item_date.text().strip().split("/")
            month = int(month_str)
            day = int(day_str)
        except ValueError:
            return None

        calendar_type = "solar" if item_calendar.text().strip() == "國曆" else "lunar"
        return {
            "id": self.selected_rule_id,
            "calendar_type": calendar_type,
            "month": month,
            "day": day,
        }

    def _on_rule_selected(self) -> None:
        selected = self.table_rules.selectedItems()
        if not selected:
            self.selected_rule_id = None
            return

        row = selected[0].row()
        item_id = self.table_rules.item(row, 0)
        item_calendar = self.table_rules.item(row, 1)
        item_date = self.table_rules.item(row, 2)

        if not item_id or not item_calendar or not item_date:
            self.selected_rule_id = None
            return

        try:
            self.selected_rule_id = int(item_id.text())
        except ValueError:
            self.selected_rule_id = None
            return

        # 僅同步 selected_rule_id；實際編輯由右鍵 popup 完成
        return

    def _import_csv(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "匯入假日設定",
            "",
            "CSV 檔案 (*.csv)",
        )
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                weekdays: List[int] = []
                date_rules: List[Dict[str, Any]] = []
                for row in reader:
                    entry_type = str(row.get("entry_type", "") or "").strip().lower()
                    if entry_type == "weekday":
                        try:
                            weekday = int(str(row.get("weekday", "") or "0").strip())
                        except ValueError:
                            continue
                        if 1 <= weekday <= 7:
                            weekdays.append(weekday)
                        continue

                    if entry_type == "date":
                        calendar_type = str(row.get("calendar_type", "") or "").strip().lower()
                        if calendar_type not in {"solar", "lunar"}:
                            continue
                        try:
                            month = int(str(row.get("month", "") or "0").strip())
                            day = int(str(row.get("day", "") or "0").strip())
                        except ValueError:
                            continue
                        if not (1 <= month <= 12 and 1 <= day <= 31):
                            continue
                        date_rules.append(
                            {
                                "calendar_type": calendar_type,
                                "month": month,
                                "day": day,
                            }
                        )
        except Exception as e:
            QMessageBox.warning(self, "失敗", f"讀取 CSV 失敗：{e}")
            return

        if not weekdays and not date_rules:
            QMessageBox.warning(self, "失敗", "CSV 內容為空或格式不符。")
            return

        if not self.db_manager.replace_holiday_rules(weekdays, date_rules):
            QMessageBox.warning(self, "失敗", "匯入假日設定失敗。")
            return

        self._load_data()
        QMessageBox.information(self, "完成", "匯入完成。")
        self._load_data()

    def _export_csv(self) -> None:
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "匯出假日設定",
            "holidays.csv",
            "CSV 檔案 (*.csv)",
        )
        if not file_path:
            return

        payload = self.db_manager.get_holiday_rules_payload()
        try:
            with open(file_path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["entry_type", "calendar_type", "month", "day", "weekday"])
                for weekday in payload.get("weekdays", []):
                    writer.writerow(["weekday", "", "", "", int(weekday)])
                sorted_dates = sorted(
                    payload.get("dates", []),
                    key=lambda rule: (
                        0 if str(rule.get("calendar_type", "")).strip().lower() == "solar" else 1,
                        int(rule.get("month", 0) or 0),
                        int(rule.get("day", 0) or 0),
                    ),
                )
                for rule in sorted_dates:
                    writer.writerow(
                        [
                            "date",
                            str(rule.get("calendar_type", "") or ""),
                            int(rule.get("month", 0) or 0),
                            int(rule.get("day", 0) or 0),
                            "",
                        ]
                    )
        except Exception as e:
            QMessageBox.warning(self, "失敗", f"寫入 CSV 失敗：{e}")
            return

        QMessageBox.information(self, "完成", "匯出完成。")


class HolidayRuleEditDialog(QDialog):
    """日期假日規則編輯 popup。"""

    def __init__(
        self,
        parent=None,
        calendar_type: str = "solar",
        month: int = 1,
        day: int = 1,
    ):
        super().__init__(parent)
        self.setWindowTitle("日期假日設定")
        self.setModal(True)
        self.resize(320, 140)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        row = QHBoxLayout()
        self.combo_calendar_type = QComboBox()
        self.combo_calendar_type.addItem("國曆", "solar")
        self.combo_calendar_type.addItem("農曆", "lunar")

        self.combo_month = QComboBox()
        for m in range(1, 13):
            self.combo_month.addItem(f"{m} 月", m)

        self.combo_day = QComboBox()

        row.addWidget(self.combo_calendar_type)
        row.addWidget(self.combo_month)
        row.addWidget(self.combo_day)
        root.addLayout(row)

        actions = QHBoxLayout()
        actions.addStretch()
        self.btn_ok = QPushButton("確定")
        self.btn_cancel = QPushButton("取消")
        actions.addWidget(self.btn_ok)
        actions.addWidget(self.btn_cancel)
        root.addLayout(actions)

        self.combo_month.currentIndexChanged.connect(self._reload_day_combo)
        self.btn_ok.clicked.connect(self.accept)
        self.btn_cancel.clicked.connect(self.reject)

        attach_combo_wheel_behavior(self)

        self._wheel_combo_targets: Dict[object, QComboBox] = {}
        self._register_combo_wheel_targets()

        idx_calendar = self.combo_calendar_type.findData(calendar_type)
        self.combo_calendar_type.setCurrentIndex(max(0, idx_calendar))

        idx_month = self.combo_month.findData(month)
        self.combo_month.setCurrentIndex(max(0, idx_month))
        self._reload_day_combo(day)

    def _register_combo_wheel_targets(self) -> None:
        self._wheel_combo_targets.clear()
        for combo in (self.combo_calendar_type, self.combo_month, self.combo_day):
            self._wheel_combo_targets[combo] = combo
            combo.installEventFilter(self)
            line_edit = combo.lineEdit()
            if line_edit is not None:
                self._wheel_combo_targets[line_edit] = combo
                line_edit.installEventFilter(self)
            view = combo.view()
            if view is not None:
                self._wheel_combo_targets[view] = combo
                self._wheel_combo_targets[view.viewport()] = combo
                view.installEventFilter(self)
                view.viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Wheel:
            combo = self._wheel_combo_targets.get(obj)
            if combo is not None and combo.isEnabled() and combo.count() > 0:
                delta = event.angleDelta().y()
                if delta != 0:
                    steps = int(delta / 120)
                    if steps == 0:
                        steps = 1 if delta > 0 else -1
                    current_index = combo.currentIndex()
                    if current_index < 0:
                        current_index = 0
                    target_index = current_index - steps
                    if target_index < 0:
                        target_index = 0
                    elif target_index >= combo.count():
                        target_index = combo.count() - 1
                    if target_index != combo.currentIndex():
                        combo.setCurrentIndex(target_index)
                event.accept()
                return True
        return super().eventFilter(obj, event)

    def _reload_day_combo(self, preferred_day: Optional[int] = None) -> None:
        if preferred_day is None:
            preferred_day = int(self.combo_day.currentData() or 1)

        month = int(self.combo_month.currentData() or 1)
        max_days = 31
        if month in {4, 6, 9, 11}:
            max_days = 30
        elif month == 2:
            max_days = 29

        self.combo_day.blockSignals(True)
        self.combo_day.clear()
        for d in range(1, max_days + 1):
            self.combo_day.addItem(f"{d} 日", d)
        idx_day = self.combo_day.findData(min(preferred_day, max_days))
        self.combo_day.setCurrentIndex(max(0, idx_day))
        self.combo_day.blockSignals(False)

    def get_values(self) -> Dict[str, int | str]:
        return {
            "calendar_type": str(self.combo_calendar_type.currentData() or "solar"),
            "month": int(self.combo_month.currentData() or 1),
            "day": int(self.combo_day.currentData() or 1),
        }
