# staff_config_tab.py

import sys
from functools import partial
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QTableWidget, QTableWidgetItem, QLineEdit, QPushButton, QCheckBox,
    QHeaderView, QGroupBox, QMessageBox, QColorDialog
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor

from core_engine import weekdays_jp, Staff, SettingsManager

class StaffConfigTab(QWidget):
    def __init__(self, settings_manager: SettingsManager):
        super().__init__()
        self.settings_manager = settings_manager
        self._init_ui()
        self._connect_signals()

    def set_settings_manager(self, settings_manager: SettingsManager):
        self.settings_manager = settings_manager
        self.load_staff_list()

    def _connect_signals(self):
        self.table.itemSelectionChanged.connect(self._on_staff_selected)
        self.add_button.clicked.connect(self._add_or_update_staff)
        self.delete_button.clicked.connect(self._delete_staff)
        self.clear_form_button.clicked.connect(self._clear_form)
        self.color_picker_button.clicked.connect(self._open_color_picker)

    def _init_ui(self):
        main_layout = QVBoxLayout(self)

        self.table = QTableWidget()
        # ★★★ 修正: カラム数を4に減らし、ヘッダーラベルを変更 ★★★
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["稼働中", "名前", "カラー", "固定の不可曜日"])
        
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)

        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        main_layout.addWidget(self.table)

        form_group = QGroupBox("スタッフ情報編集")
        main_layout.addWidget(form_group)
        
        form_outer_layout = QVBoxLayout(form_group)
        form_layout = QFormLayout()
        
        self.name_input = QLineEdit()

        self.color_input = QLineEdit()
        self.color_input.setPlaceholderText("#RRGGBB形式 (例: #FFADAD)")
        self.color_picker_button = QPushButton("色を選択...")

        color_layout = QHBoxLayout()
        color_layout.setContentsMargins(0, 0, 0, 0)
        
        color_layout.addWidget(self.color_picker_button)
        color_layout.addWidget(self.color_input)

        self.weekdays_layout = QHBoxLayout()
        self.weekday_checkboxes = {}
        for day in weekdays_jp:
            checkbox = QCheckBox(day)
            self.weekday_checkboxes[day] = checkbox
            self.weekdays_layout.addWidget(checkbox)

        form_layout.addRow("名前:", self.name_input)
        form_layout.addRow("カラー:", color_layout)
        form_layout.addRow("固定の不可曜日:", self.weekdays_layout)
        form_outer_layout.addLayout(form_layout)

        button_layout = QHBoxLayout()
        self.add_button = QPushButton("スタッフを追加/更新")
        self.delete_button = QPushButton("選択したスタッフを削除")
        self.clear_form_button = QPushButton("フォームをクリア")
        button_layout.addStretch()
        button_layout.addWidget(self.clear_form_button)
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.delete_button)
        form_outer_layout.addLayout(button_layout)
    
    def _open_color_picker(self):
        current_color_str = self.color_input.text()
        initial_color = QColor(current_color_str) if QColor.isValidColor(current_color_str) else Qt.GlobalColor.white
        color = QColorDialog.getColor(initial_color, self, "色の選択")
        if color.isValid():
            self.color_input.setText(color.name())

    def load_staff_list(self):
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        staff_list = self.settings_manager.staff_manager.get_all_staff()
        
        for staff in sorted(staff_list, key=lambda s: s.name):
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            
            active_checkbox = QCheckBox()
            active_checkbox.setChecked(staff.is_active)
            active_checkbox.stateChanged.connect(partial(self._toggle_staff_active, staff.name))
            
            wrapper = QWidget()
            layout = QHBoxLayout(wrapper)
            layout.addWidget(active_checkbox)
            layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.setContentsMargins(0,0,0,0)
            self.table.setCellWidget(row_position, 0, wrapper)

            self.table.setItem(row_position, 1, QTableWidgetItem(staff.name))
            
            color_item = QTableWidgetItem(staff.color_code)
            try:
                color_item.setBackground(QColor(staff.color_code))
            except Exception: pass
            self.table.setItem(row_position, 2, color_item)
            
            impossible_days_str = ", ".join(sorted(list(staff.impossible_weekdays)))
            self.table.setItem(row_position, 3, QTableWidgetItem(impossible_days_str))
        
        self.table.blockSignals(False)
        self.table.resizeColumnsToContents()

    def _toggle_staff_active(self, staff_name, state):
        staff = self.settings_manager.staff_manager.get_staff_by_name(staff_name)
        if staff:
            staff.is_active = (state == Qt.CheckState.Checked.value)
            print(f"スタッフ '{staff.name}' の稼働状態が '{staff.is_active}' に変更されました。")

    def _on_staff_selected(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        staff_name_item = self.table.item(row, 1) 
        if not staff_name_item: return
        
        staff_name = staff_name_item.text()
        staff = self.settings_manager.staff_manager.get_staff_by_name(staff_name)

        if not staff:
            return

        self._clear_form(clear_selection=False)
        self.name_input.setText(staff.name)
        self.name_input.setReadOnly(True)
        self.color_input.setText(staff.color_code)

        for day, checkbox in self.weekday_checkboxes.items():
            checkbox.setChecked(day in staff.impossible_weekdays)
        
    def _add_or_update_staff(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "入力エラー", "スタッフ名を入力してください。")
            return
        
        color = self.color_input.text().strip()
        if not QColor.isValidColor(color) or not color.startswith('#'):
            QMessageBox.warning(self, "入力エラー", "カラーコードが不正な形式です。'#RRGGBB' の形式で入力してください。 (例: #ffadad)")
            return

        impossible_weekdays = {day for day, cb in self.weekday_checkboxes.items() if cb.isChecked()}
        
        existing_staff = self.settings_manager.staff_manager.get_staff_by_name(name)
        is_active = existing_staff.is_active if existing_staff else True

        new_staff = Staff(
            name=name,
            color_code=color.lower(),
            is_active=is_active,
            impossible_weekdays=impossible_weekdays
        )

        self.settings_manager.staff_manager.add_or_update_staff(new_staff)
        self.load_staff_list()
        self._clear_form()

    def _delete_staff(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "選択エラー", "削除するスタッフをリストから選択してください。")
            return

        row = selected_rows[0].row()
        staff_name_item = self.table.item(row, 1)
        if not staff_name_item: return

        staff_name = staff_name_item.text()

        reply = QMessageBox.question(self, '削除の確認', 
                                     f"本当にスタッフ '{staff_name}' を削除しますか？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            self.settings_manager.staff_manager.remove_staff_by_name(staff_name)
            self.load_staff_list()
            self._clear_form()

    def _clear_form(self, clear_selection=True):
        self.name_input.clear()
        self.name_input.setReadOnly(False)
        self.color_input.clear()
        for checkbox in self.weekday_checkboxes.values():
            checkbox.setChecked(False)
        if clear_selection:
            self.table.clearSelection()