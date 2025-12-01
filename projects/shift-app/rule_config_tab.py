import sys
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QPushButton, QGroupBox, QComboBox, QListWidget,
    QMessageBox
)
from PySide6.QtCore import Qt

# core_engineから必要なクラスをインポート
from core_engine import weekdays_jp, SettingsManager, RuleBasedFixedShift, RuleBasedVacation

class RuleConfigTab(QWidget):
    """
    ルールベースの固定シフトと休暇を設定するためのタブUI。
    """
    def __init__(self, settings_manager: SettingsManager):
        super().__init__()
        self.settings_manager = settings_manager
        self._init_ui()
        self._connect_signals()
        self.load_rules()

    def set_settings_manager(self, settings_manager: SettingsManager):
        """メインウィンドウから新しいSettingsManagerが渡されたときに更新する"""
        self.settings_manager = settings_manager
        self.load_rules()

    def _connect_signals(self):
        self.add_fixed_button.clicked.connect(self._add_fixed_rule)
        self.delete_fixed_button.clicked.connect(self._delete_fixed_rule)
        self.add_vacation_button.clicked.connect(self._add_vacation_rule)
        self.delete_vacation_button.clicked.connect(self._delete_vacation_rule)

    def _init_ui(self):
        """UIの初期化とレイアウト設定"""
        main_layout = QHBoxLayout(self)

        # --- 固定シフト設定 ---
        fixed_shift_group = QGroupBox("固定シフト設定 (第N M曜日はこの人)")
        fixed_shift_layout = QVBoxLayout()

        # 入力フォーム
        fixed_form_layout = QFormLayout()
        self.fixed_week_combo = QComboBox()
        self.fixed_week_combo.addItems([f"第{i}" for i in range(1, 5)] + ["最終"])
        self.fixed_weekday_combo = QComboBox()
        self.fixed_weekday_combo.addItems(weekdays_jp)
        self.fixed_staff_combo = QComboBox() # スタッフ名は後でロード

        fixed_form_layout.addRow("週:", self.fixed_week_combo)
        fixed_form_layout.addRow("曜日:", self.fixed_weekday_combo)
        fixed_form_layout.addRow("スタッフ:", self.fixed_staff_combo)
        fixed_shift_layout.addLayout(fixed_form_layout)

        # 操作ボタン
        self.add_fixed_button = QPushButton("固定シフトを追加")
        fixed_shift_layout.addWidget(self.add_fixed_button)

        # 一覧
        self.fixed_list = QListWidget()
        fixed_shift_layout.addWidget(self.fixed_list)
        self.delete_fixed_button = QPushButton("選択したルールを削除")
        fixed_shift_layout.addWidget(self.delete_fixed_button)

        fixed_shift_group.setLayout(fixed_shift_layout)
        main_layout.addWidget(fixed_shift_group)

        # --- 休暇ルール設定 ---
        vacation_group = QGroupBox("休暇ルール設定 (第N M曜日はこの人休み)")
        vacation_layout = QVBoxLayout()

        # 入力フォーム
        vacation_form_layout = QFormLayout()
        self.vacation_week_combo = QComboBox()
        self.vacation_week_combo.addItems([f"第{i}" for i in range(1, 5)] + ["最終"])
        self.vacation_weekday_combo = QComboBox()
        self.vacation_weekday_combo.addItems(weekdays_jp)
        self.vacation_staff_combo = QComboBox() # スタッフ名は後でロード

        vacation_form_layout.addRow("週:", self.vacation_week_combo)
        vacation_form_layout.addRow("曜日:", self.vacation_weekday_combo)
        vacation_form_layout.addRow("スタッフ:", self.vacation_staff_combo)
        vacation_layout.addLayout(vacation_form_layout)

        # 操作ボタン
        self.add_vacation_button = QPushButton("休暇ルールを追加")
        vacation_layout.addWidget(self.add_vacation_button)

        # 一覧
        self.vacation_list = QListWidget()
        vacation_layout.addWidget(self.vacation_list)
        self.delete_vacation_button = QPushButton("選択したルールを削除")
        vacation_layout.addWidget(self.delete_vacation_button)

        vacation_group.setLayout(vacation_layout)
        main_layout.addWidget(vacation_group)

    def update_staff_list(self):
        """スタッフリストが変更されたときにコンボボックスを更新する"""
        staff_names = sorted([s.name for s in self.settings_manager.staff_manager.get_all_staff()])
        for combo in [self.fixed_staff_combo, self.vacation_staff_combo]:
            current_text = combo.currentText()
            combo.clear()
            combo.addItems(staff_names)
            if current_text in staff_names:
                combo.setCurrentText(current_text)

    def load_rules(self):
        """SettingsManagerからルールを読み込み、リストを更新する"""
        # 固定シフトのルールを読み込み
        self.fixed_list.clear()
        fixed_rules = sorted(self.settings_manager.rule_based_fixed_shifts, key=lambda r: (r.week_number, r.weekday_index, r.staff.name))
        for rule in fixed_rules:
            self.fixed_list.addItem(self._format_rule(rule))

        # 休暇のルールを読み込み
        self.vacation_list.clear()
        vacation_rules = sorted(self.settings_manager.rule_based_vacations, key=lambda r: (r.week_number, r.weekday_index, r.staff_name))
        for rule in vacation_rules:
            self.vacation_list.addItem(self._format_rule(rule))

    def _add_fixed_rule(self):
        staff_name = self.fixed_staff_combo.currentText()
        if not staff_name:
            QMessageBox.warning(self, "選択エラー", "スタッフを選択してください。")
            return

        staff = self.settings_manager.staff_manager.get_staff_by_name(staff_name)
        if not staff: return

        week_text = self.fixed_week_combo.currentText()
        weekday_text = self.fixed_weekday_combo.currentText()
        week_number = 5 if week_text == "最終" else int(week_text[1])
        weekday_index = weekdays_jp.index(weekday_text)

        new_rule = RuleBasedFixedShift(week_number, weekday_index, staff)

        if new_rule in self.settings_manager.rule_based_fixed_shifts:
            QMessageBox.information(self, "情報", "このルールは既に追加されています。")
            return

        self.settings_manager.rule_based_fixed_shifts.append(new_rule)
        self.load_rules()

    def _delete_fixed_rule(self):
        selected_items = self.fixed_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "選択エラー", "削除する固定シフトルールを選択してください。")
            return

        for item in selected_items:
            rule_to_delete = self._parse_rule(item.text(), 'fixed')
            if rule_to_delete in self.settings_manager.rule_based_fixed_shifts:
                self.settings_manager.rule_based_fixed_shifts.remove(rule_to_delete)
        self.load_rules()

    def _add_vacation_rule(self):
        staff_name = self.vacation_staff_combo.currentText()
        if not staff_name:
            QMessageBox.warning(self, "選択エラー", "スタッフを選択してください。")
            return

        week_text = self.vacation_week_combo.currentText()
        weekday_text = self.vacation_weekday_combo.currentText()
        week_number = 5 if week_text == "最終" else int(week_text[1])
        weekday_index = weekdays_jp.index(weekday_text)

        new_rule = RuleBasedVacation(week_number, weekday_index, staff_name)

        if new_rule in self.settings_manager.rule_based_vacations:
            QMessageBox.information(self, "情報", "このルールは既に追加されています。")
            return

        self.settings_manager.rule_based_vacations.append(new_rule)
        self.load_rules()

    def _delete_vacation_rule(self):
        selected_items = self.vacation_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "選択エラー", "削除する休暇ルールを選択してください。")
            return

        for item in selected_items:
            rule_to_delete = self._parse_rule(item.text(), 'vacation')
            if rule_to_delete in self.settings_manager.rule_based_vacations:
                self.settings_manager.rule_based_vacations.remove(rule_to_delete)
        self.load_rules()

    def _format_rule(self, rule):
        week_text = f"第{rule.week_number}" if rule.week_number != 5 else "最終"
        weekday_text = weekdays_jp[rule.weekday_index]
        staff_name = rule.staff.name if isinstance(rule, RuleBasedFixedShift) else rule.staff_name
        return f"{week_text}{weekday_text}: {staff_name}"

    def _parse_rule(self, text, rule_type):
        try:
            rule_part, staff_name = text.split(': ')
            week_number = 5 if rule_part.startswith("最終") else int(rule_part[1])
            weekday_text = rule_part[2:] if rule_part.startswith("最終") else rule_part[2:]
            weekday_index = weekdays_jp.index(weekday_text)

            if rule_type == 'fixed':
                staff = self.settings_manager.staff_manager.get_staff_by_name(staff_name)
                return RuleBasedFixedShift(week_number, weekday_index, staff) if staff else None
            elif rule_type == 'vacation':
                return RuleBasedVacation(week_number, weekday_index, staff_name)
        except (ValueError, IndexError):
            return None