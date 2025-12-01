import sys
from functools import partial
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout, QHBoxLayout, QGridLayout,
    QSpinBox, QGroupBox, QLabel, QCheckBox, QLineEdit, QFrame, QSpacerItem, QSizePolicy
)
from PySide6.QtCore import Qt
from core_engine import SettingsManager, weekdays_jp

class GeneralSettingsTab(QWidget):
    def __init__(self, settings_manager: SettingsManager):
        super().__init__()
        self.settings_manager = settings_manager
        self.fairness_checkboxes = {} 
        self.shifts_per_day_spinboxes = {} 
        self._init_ui()
        self._connect_signals()
        self.load_settings()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        
        general_group = QGroupBox("基本シフト設定")
        general_form_layout = QFormLayout(general_group)
        self.min_interval_spinbox = QSpinBox()
        self.min_interval_spinbox.setRange(0, 30)
        self.min_interval_spinbox.setSuffix(" 日")
        general_form_layout.addRow("最低勤務間隔:", self.min_interval_spinbox)
        self.max_consecutive_days_spinbox = QSpinBox()
        self.max_consecutive_days_spinbox.setRange(1, 30)
        self.max_consecutive_days_spinbox.setSuffix(" 日")
        general_form_layout.addRow("最大連勤日数:", self.max_consecutive_days_spinbox)
        self.ignore_rules_on_holidays_checkbox = QCheckBox("祝日には曜日ベースのルールを適用しない")
        general_form_layout.addRow(self.ignore_rules_on_holidays_checkbox)

        output_group = QGroupBox("出力設定")
        output_form_layout = QFormLayout(output_group)
        self.excel_title_input = QLineEdit()
        self.excel_title_input.setPlaceholderText("例: 日直・当直予定表")
        output_form_layout.addRow("Excel出力用タイトル:", self.excel_title_input)
        
        shifts_per_day_group = QGroupBox("曜日・祝日ごとの担当人数（範囲指定）")
        shifts_per_day_layout = QVBoxLayout(shifts_per_day_group)
        
        self.common_shifts_checkbox = QCheckBox("全曜日・祝日で共通の人数を設定する")
        shifts_per_day_layout.addWidget(self.common_shifts_checkbox)

        common_layout = QHBoxLayout()
        self.common_shifts_min_spinbox = QSpinBox()
        self.common_shifts_min_spinbox.setRange(0, 99)
        self.common_shifts_max_spinbox = QSpinBox()
        self.common_shifts_max_spinbox.setRange(0, 99)
        common_layout.addWidget(QLabel("共通人数:"))
        common_layout.addWidget(self.common_shifts_min_spinbox)
        common_layout.addWidget(QLabel("～"))
        common_layout.addWidget(self.common_shifts_max_spinbox)
        common_layout.addWidget(QLabel("人"))
        common_layout.addStretch()
        shifts_per_day_layout.addLayout(common_layout)

        # ★★★★★ ここからレイアウトを精密に調整 ★★★★★
        self.per_day_widget = QWidget()
        per_day_grid_layout = QGridLayout(self.per_day_widget)
        per_day_grid_layout.setContentsMargins(5, 10, 5, 10)
        per_day_grid_layout.setHorizontalSpacing(0)
        per_day_grid_layout.setVerticalSpacing(10)

        day_options = weekdays_jp + ("祝",)
        
        for i, day in enumerate(day_options):
            col = i % 4
            row = (i // 4) * 3 # 1グループで3行使う (ラベル、スピンボックス、線)

            day_label = QLabel(day)
            day_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            day_label.setStyleSheet("font-weight: bold;")

            min_spinbox = QSpinBox()
            min_spinbox.setRange(0, 99)
            min_spinbox.setFixedWidth(65)

            max_spinbox = QSpinBox()
            max_spinbox.setRange(0, 99)
            max_spinbox.setFixedWidth(65)
            
            self.shifts_per_day_spinboxes[day] = {'min': min_spinbox, 'max': max_spinbox}

            spinbox_layout = QHBoxLayout()
            spinbox_layout.setContentsMargins(0,0,0,0)
            spinbox_layout.setSpacing(0)
            spinbox_layout.addStretch(1) # 左のスペーサー
            spinbox_layout.addWidget(min_spinbox)
            spinbox_layout.addWidget(QLabel("～"))
            spinbox_layout.addWidget(max_spinbox)
            spinbox_layout.addStretch(1) # 右のスペーサー
            
            per_day_grid_layout.addWidget(day_label, row, col)
            per_day_grid_layout.addLayout(spinbox_layout, row + 1, col)
        
        # --- 行の区切り線を追加 ---
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        per_day_grid_layout.addWidget(line, 2, 0, 1, 4)
        
        shifts_per_day_layout.addWidget(self.per_day_widget)

        solver_group = QGroupBox("ソルバー設定")
        solver_form_layout = QFormLayout()
        self.max_solutions_spinbox = QSpinBox()
        self.max_solutions_spinbox.setRange(1, 1000)
        solver_form_layout.addRow("最大解探索数:", self.max_solutions_spinbox)
        self.fairness_tolerance_spinbox = QSpinBox()
        self.fairness_tolerance_spinbox.setRange(0, 10)
        self.fairness_tolerance_spinbox.setSuffix(" 回")
        solver_form_layout.addRow("公平性の許容差:", self.fairness_tolerance_spinbox)
        solver_group.setLayout(solver_form_layout)
        
        fairness_group = QGroupBox("公平性評価の対象")
        fairness_layout = QVBoxLayout(fairness_group)
        fairness_desc_label = QLabel("ここでチェックした曜日の担当回数が均等になるようにシフトが生成されます。")
        fairness_layout.addWidget(fairness_desc_label)
        checkbox_layout = QHBoxLayout()
        for day in day_options:
            checkbox = QCheckBox(day)
            self.fairness_checkboxes[day] = checkbox
            checkbox_layout.addWidget(checkbox)
        fairness_layout.addLayout(checkbox_layout)
        self.disperse_duties_checkbox = QCheckBox("同じ曜日の担当をできるだけ分散させる")
        self.disperse_duties_checkbox.setToolTip(
            "チェックを入れると、公平性対象の曜日や祝日について、担当機会をできるだけ分散させます"
            "（例：先週土曜に担当したスタッフは、今週の土曜を担当しにくくなります）。\n"
            "ただし、他に担当できる人がいないなど、シフト作成が困難な場合はこの限りではありません。"
        )
        fairness_layout.addWidget(self.disperse_duties_checkbox)
        
        #self.avoid_consecutive_weekday_checkbox = QCheckBox("同じ曜日の連続担当を避ける（上記でチェックした曜日のみ）")
        #self.avoid_consecutive_weekday_checkbox.setToolTip("チェックを入れると、例えば「先週の土曜担当」と「今週の土曜担当」が同じスタッフになるのを防ぎます。")
        # ★★★★★ ここで disperse_duties_checkbox と avoid_consecutive_weekday_checkbox を入れ替えます ★★★★★
        # 新しい分散機能は、古い連続担当回避機能の進化版なので、上に配置します。
        # 順番を入れ替え
        #fairness_layout.removeWidget(self.avoid_consecutive_weekday_checkbox)
        #fairness_layout.addWidget(self.avoid_consecutive_weekday_checkbox)

        main_layout.addWidget(general_group)
        main_layout.addWidget(output_group)
        main_layout.addWidget(shifts_per_day_group) 
        main_layout.addWidget(solver_group)
        main_layout.addWidget(fairness_group)
        main_layout.addStretch()

    def _disconnect_signals(self):
        self.min_interval_spinbox.blockSignals(True)
        self.max_consecutive_days_spinbox.blockSignals(True)
        self.ignore_rules_on_holidays_checkbox.blockSignals(True)
        self.excel_title_input.blockSignals(True)
        self.max_solutions_spinbox.blockSignals(True)
        self.fairness_tolerance_spinbox.blockSignals(True)
        self.common_shifts_checkbox.blockSignals(True)
        self.common_shifts_min_spinbox.blockSignals(True)
        self.common_shifts_max_spinbox.blockSignals(True)
        for spinbox_dict in self.shifts_per_day_spinboxes.values():
            spinbox_dict['min'].blockSignals(True)
            spinbox_dict['max'].blockSignals(True)
        for checkbox in self.fairness_checkboxes.values():
            checkbox.blockSignals(True)
        self.disperse_duties_checkbox.blockSignals(True) # ★追加
        #self.avoid_consecutive_weekday_checkbox.blockSignals(True)

    def _connect_signals(self):
        self.min_interval_spinbox.valueChanged.connect(lambda val: setattr(self.settings_manager, 'min_interval', val))
        self.max_consecutive_days_spinbox.valueChanged.connect(lambda val: setattr(self.settings_manager, 'max_consecutive_days', val))
        self.ignore_rules_on_holidays_checkbox.stateChanged.connect(lambda state: setattr(self.settings_manager, 'ignore_rules_on_holidays', state == Qt.CheckState.Checked.value))
        self.excel_title_input.textChanged.connect(lambda text: setattr(self.settings_manager, 'excel_title', text))
        self.max_solutions_spinbox.valueChanged.connect(lambda val: setattr(self.settings_manager, 'max_solutions', val))
        self.fairness_tolerance_spinbox.valueChanged.connect(lambda val: setattr(self.settings_manager, 'fairness_tolerance', val))
        self.common_shifts_checkbox.stateChanged.connect(self._update_shifts_per_day_mode)
        self.common_shifts_min_spinbox.valueChanged.connect(self._update_common_shifts_setting)
        self.common_shifts_max_spinbox.valueChanged.connect(self._update_common_shifts_setting)
        for spinbox_dict in self.shifts_per_day_spinboxes.values():
            spinbox_dict['min'].valueChanged.connect(self._update_per_day_shifts_setting)
            spinbox_dict['max'].valueChanged.connect(self._update_per_day_shifts_setting)
        for checkbox in self.fairness_checkboxes.values():
            checkbox.stateChanged.connect(self._update_fairness_group)
        self.disperse_duties_checkbox.stateChanged.connect(lambda state: setattr(self.settings_manager, 'disperse_duties', state == Qt.CheckState.Checked.value)) # ★追加
        #self.avoid_consecutive_weekday_checkbox.stateChanged.connect(lambda state: setattr(self.settings_manager, 'avoid_consecutive_same_weekday', state == Qt.CheckState.Checked.value))
        
        self.min_interval_spinbox.blockSignals(False)
        self.max_consecutive_days_spinbox.blockSignals(False)
        self.ignore_rules_on_holidays_checkbox.blockSignals(False)
        self.excel_title_input.blockSignals(False)
        self.max_solutions_spinbox.blockSignals(False)
        self.fairness_tolerance_spinbox.blockSignals(False)
        self.common_shifts_checkbox.blockSignals(False)
        self.common_shifts_min_spinbox.blockSignals(False)
        self.common_shifts_max_spinbox.blockSignals(False)
        for spinbox_dict in self.shifts_per_day_spinboxes.values():
            spinbox_dict['min'].blockSignals(False)
            spinbox_dict['max'].blockSignals(False)
        for checkbox in self.fairness_checkboxes.values():
            checkbox.blockSignals(False)
        self.disperse_duties_checkbox.blockSignals(False) # ★追加
        #self.avoid_consecutive_weekday_checkbox.blockSignals(False)

    def _update_shifts_per_day_mode(self):
        is_common = self.common_shifts_checkbox.isChecked()
        self.common_shifts_min_spinbox.setEnabled(is_common)
        self.common_shifts_max_spinbox.setEnabled(is_common)
        self.per_day_widget.setDisabled(is_common)
        if self.common_shifts_checkbox.signalsBlocked(): return
        if is_common:
            self._update_common_shifts_setting()
        else:
            self._update_per_day_shifts_setting()

    def _update_common_shifts_setting(self):
        if self.common_shifts_checkbox.isChecked():
            min_val = self.common_shifts_min_spinbox.value()
            max_val = self.common_shifts_max_spinbox.value()
            if min_val > max_val:
                self.common_shifts_max_spinbox.setValue(min_val)
                max_val = min_val
            self.settings_manager.shifts_per_day = {'min': min_val, 'max': max_val}
            print(f"共通担当人数が更新されました: {self.settings_manager.shifts_per_day}")

    def _update_per_day_shifts_setting(self):
        if not self.common_shifts_checkbox.isChecked():
            settings_dict = {}
            for day, spinbox_dict in self.shifts_per_day_spinboxes.items():
                min_val = spinbox_dict['min'].value()
                max_val = spinbox_dict['max'].value()
                if min_val > max_val:
                    spinbox_dict['max'].setValue(min_val)
                    max_val = min_val
                settings_dict[day] = {'min': min_val, 'max': max_val}
            self.settings_manager.shifts_per_day = settings_dict
            print(f"曜日別担当人数が更新されました: {self.settings_manager.shifts_per_day}")

    def _update_fairness_group(self):
        if any(cb.signalsBlocked() for cb in self.fairness_checkboxes.values()):
            return
        self.settings_manager.fairness_group.clear()
        for day, checkbox in self.fairness_checkboxes.items():
            if checkbox.isChecked():
                self.settings_manager.fairness_group.add(day)
        print(f"公平性グループが更新されました: {self.settings_manager.fairness_group}")

    def load_settings(self):
        self._disconnect_signals()
        
        self.min_interval_spinbox.setValue(self.settings_manager.min_interval)
        self.max_consecutive_days_spinbox.setValue(self.settings_manager.max_consecutive_days)
        self.ignore_rules_on_holidays_checkbox.setChecked(self.settings_manager.ignore_rules_on_holidays)
        self.excel_title_input.setText(self.settings_manager.excel_title)
        
        shifts_setting = self.settings_manager.shifts_per_day
        if isinstance(shifts_setting, int) or (isinstance(shifts_setting, dict) and 'min' not in shifts_setting.get(weekdays_jp[0], {})):
            self.common_shifts_checkbox.setChecked(True)
            min_val = shifts_setting if isinstance(shifts_setting, int) else shifts_setting.get('min', 1)
            max_val = shifts_setting if isinstance(shifts_setting, int) else shifts_setting.get('max', 1)
            self.common_shifts_min_spinbox.setValue(min_val)
            self.common_shifts_max_spinbox.setValue(max_val)
            for day, spinbox_dict in self.shifts_per_day_spinboxes.items():
                spinbox_dict['min'].setValue(min_val)
                spinbox_dict['max'].setValue(max_val)
        elif isinstance(shifts_setting, dict):
            self.common_shifts_checkbox.setChecked(False)
            avg_min = int(sum(d.get('min', 1) for d in shifts_setting.values()) / len(shifts_setting)) if shifts_setting else 1
            avg_max = int(sum(d.get('max', 1) for d in shifts_setting.values()) / len(shifts_setting)) if shifts_setting else 1
            self.common_shifts_min_spinbox.setValue(avg_min)
            self.common_shifts_max_spinbox.setValue(avg_max)
            for day, spinbox_dict in self.shifts_per_day_spinboxes.items():
                day_setting = shifts_setting.get(day, {'min': 1, 'max': 1})
                spinbox_dict['min'].setValue(day_setting.get('min', 1))
                spinbox_dict['max'].setValue(day_setting.get('max', 1))
        
        is_common = self.common_shifts_checkbox.isChecked()
        self.common_shifts_min_spinbox.setEnabled(is_common)
        self.common_shifts_max_spinbox.setEnabled(is_common)
        self.per_day_widget.setDisabled(is_common)

        self.max_solutions_spinbox.setValue(self.settings_manager.max_solutions)
        self.fairness_tolerance_spinbox.setValue(self.settings_manager.fairness_tolerance)
        
        for day, checkbox in self.fairness_checkboxes.items():
            checkbox.setChecked(day in self.settings_manager.fairness_group)

        self.disperse_duties_checkbox.setChecked(self.settings_manager.disperse_duties) # ★追加

        #self.avoid_consecutive_weekday_checkbox.setChecked(self.settings_manager.avoid_consecutive_same_weekday)

        self._connect_signals()

    def set_settings_manager(self, settings_manager: SettingsManager):
        self.settings_manager = settings_manager
        self.load_settings()