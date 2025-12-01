import sys
import csv
import calendar
import datetime
import holidays
import os
import platform
import webbrowser
from dateutil.relativedelta import relativedelta

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QSplitter,
    QGroupBox, QLabel, QSpinBox, QComboBox, QPushButton,
    QMessageBox, QProgressDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QCheckBox, QListWidget, QDialog,
    QCalendarWidget, QTabWidget, QStyledItemDelegate, QStyleOptionViewItem,
    QRadioButton, QDialogButtonBox, QLineEdit # ★ 追加: QLineEdit を使用
)
from PySide6.QtCore import Qt, QDate, QThread, Signal, QRect, QPoint, QModelIndex
from PySide6.QtGui import QColor, QPainter, QBrush, QTextCharFormat

from excel_exporter import export_to_excel
from pdf_exporter import export_to_pdf
from core_engine import SettingsManager, ShiftScheduler, generate_calendar_with_holidays, weekdays_jp

# --- 出力オプションダイアログ ---
class OutputOptionsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ファイル出力オプション")

        layout = QVBoxLayout(self)

        # レイアウト選択
        layout_group = QGroupBox("レイアウトの選択")
        layout_form = QFormLayout(layout_group)
        self.grid_radio = QRadioButton("グリッド形式 (従来のカレンダー)")
        self.list_radio = QRadioButton("タイムラインリスト形式 (モダンな一覧表)")
        self.grid_radio.setChecked(True)
        note_label = QLabel("<small><i>複数人担当のシフトをグリッド形式で出力すると、<br>セルの内容が見切れる可能性があります。</i></small>")
        note_label.setWordWrap(True)
        layout_form.addRow(self.grid_radio)
        layout_form.addRow(self.list_radio)
        layout_form.addRow(note_label)
        layout.addWidget(layout_group)

        # ファイル形式選択
        format_group = QGroupBox("ファイル形式の選択")
        format_form = QFormLayout(format_group)
        self.excel_radio = QRadioButton("Excel (.xlsx)")
        self.pdf_radio = QRadioButton("PDF (.pdf)")
        self.excel_radio.setChecked(True)
        format_form.addRow(self.excel_radio)
        format_form.addRow(self.pdf_radio)
        layout.addWidget(format_group)

        # OK / Cancel ボタン
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def get_options(self):
        layout = 'grid' if self.grid_radio.isChecked() else 'list'
        file_format = 'excel' if self.excel_radio.isChecked() else 'pdf'
        return layout, file_format

class StaffColorDelegate(QStyledItemDelegate):
    # ... (変更なし) ...
    INDICATOR_SIZE = 10
    INDICATOR_MARGIN = 4

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex):
        option.displayAlignment = Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft
        super().paint(painter, option, index)

        staff_list = index.data(Qt.ItemDataRole.UserRole)
        
        if not staff_list:
            return

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        text = self.parent().item(index.row(), index.column()).text()
        text_rect = option.rect.adjusted(5, 2, -5, -2)
        actual_text_rect = painter.fontMetrics().boundingRect(text_rect, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft, text)
        
        start_x = actual_text_rect.right() + self.INDICATOR_MARGIN * 2
        start_y = actual_text_rect.top()
        
        current_x = start_x
        current_y = start_y

        for i, staff in enumerate(staff_list):
            if current_x + self.INDICATOR_SIZE > option.rect.right() - self.INDICATOR_MARGIN:
                current_x = start_x
                current_y += self.INDICATOR_SIZE + self.INDICATOR_MARGIN

            color = QColor(staff.color_code)
            painter.setBrush(QBrush(color))
            painter.setPen(Qt.PenStyle.NoPen)

            indicator_rect = QRect(
                current_x,
                current_y,
                self.INDICATOR_SIZE,
                self.INDICATOR_SIZE
            )
            
            if option.rect.contains(indicator_rect):
                painter.drawEllipse(indicator_rect)

            current_x += self.INDICATOR_SIZE + self.INDICATOR_MARGIN

        painter.restore()

class GenerationWorker(QThread):
    finished = Signal(object, str)

    def __init__(self, settings_manager: SettingsManager, year: int, month: int, 
                 total_adjustments: dict, fairness_adjustments: dict,
                 previous_counts: dict, previous_fairness_counts: dict,
                 last_month_end_dates: dict,
                 prev_month_consecutive_days: dict,
                 last_week_assignments: dict,
                 manual_vacations: dict, no_shift_dates: list,
                 manual_fixed_shifts: dict,
                 parent=None):
        super().__init__(parent)
        self.settings_manager = settings_manager
        self.year = year
        self.month = month
        self.total_adjustments = total_adjustments
        self.fairness_adjustments = fairness_adjustments
        self.previous_counts = previous_counts
        self.previous_fairness_counts = previous_fairness_counts
        self.last_month_end_dates = last_month_end_dates
        self.prev_month_consecutive_days = prev_month_consecutive_days
        self.last_week_assignments = last_week_assignments
        self.manual_vacations = manual_vacations
        self.no_shift_dates = no_shift_dates
        self.manual_fixed_shifts = manual_fixed_shifts
        # このプロパティは `_start_generation` で上書きされる
        self.disperse_duties = True
        self.fairness_group = set()
        self.past_schedules = {}

    def run(self):
        try:
            calendar_data = generate_calendar_with_holidays(self.year, self.month)
            scheduler = ShiftScheduler(
                self.settings_manager.staff_manager, 
                calendar_data,
                ignore_rules_on_holidays=self.settings_manager.ignore_rules_on_holidays
            )
            
            solutions_or_error = scheduler.solve(
                shifts_per_day=self.settings_manager.shifts_per_day,
                min_interval=self.settings_manager.min_interval,
                max_consecutive_days=self.settings_manager.max_consecutive_days,
                max_solutions=self.settings_manager.max_solutions,
                fairness_tolerance=self.settings_manager.fairness_tolerance,
                last_month_end_dates=self.last_month_end_dates,
                prev_month_consecutive_days=self.prev_month_consecutive_days,
                last_week_assignments=self.last_week_assignments,
                avoid_consecutive_same_weekday=self.settings_manager.avoid_consecutive_same_weekday, # 引数は残すがエンジン側で使われない
                no_shift_dates=self.no_shift_dates,
                manual_fixed_shifts=self.manual_fixed_shifts,
                rule_based_fixed_shifts=self.settings_manager.rule_based_fixed_shifts,
                rule_based_vacations=self.settings_manager.rule_based_vacations,
                vacations=self.manual_vacations,
                fairness_group=self.fairness_group, # ★★★★★ `_start_generation`から渡されたものを使う
                total_adjustments=self.total_adjustments,
                fairness_adjustments=self.fairness_adjustments, # ★★★★★ ここにカンマを追加 ★★★★★
                disperse_duties=self.disperse_duties,
                past_schedules=self.past_schedules
            )
            
            if isinstance(solutions_or_error, str):
                self.finished.emit([], solutions_or_error)
            else:
                self.finished.emit(solutions_or_error, "")

        except Exception as e:
            error_message = f"シフト生成中に予期せぬエラーが発生しました:\n{e}"
            print(error_message)
            import traceback
            traceback.print_exc()
            self.finished.emit([], error_message)

class GenerationTab(QWidget):
    def __init__(self, settings_manager: SettingsManager):
        super().__init__()
        # ... (self.プロパティの初期化は変更なし) ...
        self.settings_manager = settings_manager
        self.worker = None
        self.solutions = []
        self.prev_month_schedule = {}
        self.last_month_end_dates = {}
        self.prev_month_consecutive_days = {}
        self.last_week_assignments = {}
        self.last_save_directory = os.path.expanduser("~")
        self._init_ui()
        self._connect_signals()
        self.update_options_ui()
        
        self.preview_table.setItemDelegate(StaffColorDelegate(self.preview_table))
        # History view wiring
        try:
            self.history_reload_button.clicked.connect(self._refresh_history_list)
            self.history_open_dir_button.clicked.connect(self._open_history_dir)
            self.history_delete_button.clicked.connect(self._delete_selected_history)
            self.history_table.itemSelectionChanged.connect(self._on_history_selected)
            self._refresh_history_list()
        except Exception:
            pass

    def set_settings_manager(self, settings_manager: SettingsManager):
        # ... (変更なし) ...
        self.settings_manager = settings_manager
        print("GenerationTabのSettingsManagerが更新されました。")
        self.update_options_ui()

    def _init_ui(self):
        # ... (前半のUI初期化は変更なし) ...
        main_layout = QHBoxLayout(self)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_widget.setFixedWidth(400)
        date_group = QGroupBox("1. 対象年月")
        date_form_layout = QFormLayout()
        current_date = QDate.currentDate()
        self.year_spinbox = QSpinBox()
        self.year_spinbox.setRange(2020, 2050)
        self.year_spinbox.setValue(current_date.year())
        self.month_combo = QComboBox()
        self.month_combo.addItems([f"{i}月" for i in range(1, 13)])
        self.month_combo.setCurrentIndex(current_date.month() - 1)
        date_form_layout.addRow("年:", self.year_spinbox)
        date_form_layout.addRow("月:", self.month_combo)
        date_group.setLayout(date_form_layout)
        left_layout.addWidget(date_group)
        options_group = QGroupBox("2. オプション設定 (任意)")
        options_main_layout = QVBoxLayout(options_group)
        options_tab_widget = QTabWidget()
        options_main_layout.addWidget(options_tab_widget)
        monthly_tab = QWidget()
        monthly_layout = QVBoxLayout(monthly_tab)
        monthly_sub_tab_widget = QTabWidget()
        monthly_layout.addWidget(monthly_sub_tab_widget)
        fixed_shift_sub_tab = QWidget()
        fixed_shift_layout = QVBoxLayout(fixed_shift_sub_tab)
        fixed_shift_group = QGroupBox("月限定固定シフト")
        fixed_shift_inner_layout = QVBoxLayout()
        fixed_shift_form_layout = QFormLayout()
        self.fixed_shift_staff_combo = QComboBox()
        self.add_fixed_shift_button = QPushButton("固定日を追加・編集...")
        fixed_shift_form_layout.addRow("スタッフ:", self.fixed_shift_staff_combo)
        fixed_shift_form_layout.addRow(self.add_fixed_shift_button)
        self.fixed_shift_list = QListWidget()
        self.delete_fixed_shift_button = QPushButton("選択した固定シフトを削除")
        fixed_shift_inner_layout.addLayout(fixed_shift_form_layout)
        fixed_shift_inner_layout.addWidget(self.fixed_shift_list)
        fixed_shift_inner_layout.addWidget(self.delete_fixed_shift_button)
        fixed_shift_group.setLayout(fixed_shift_inner_layout)
        fixed_shift_layout.addWidget(fixed_shift_group)
        fixed_shift_layout.addStretch()
        vacation_sub_tab = QWidget()
        vacation_layout = QVBoxLayout(vacation_sub_tab)
        vacation_group = QGroupBox("月限定の休暇")
        vacation_inner_layout = QVBoxLayout()
        vacation_form_layout = QFormLayout()
        self.vacation_staff_combo = QComboBox()
        vacation_form_layout.addRow("スタッフ:", self.vacation_staff_combo)
        self.add_vacation_button = QPushButton("休暇日を追加・編集...")
        vacation_form_layout.addRow(self.add_vacation_button)
        self.vacation_list = QListWidget()
        self.delete_vacation_button = QPushButton("選択した休暇を削除")
        vacation_inner_layout.addLayout(vacation_form_layout)
        vacation_inner_layout.addWidget(self.vacation_list)
        vacation_inner_layout.addWidget(self.delete_vacation_button)
        vacation_group.setLayout(vacation_inner_layout)
        vacation_layout.addWidget(vacation_group)
        vacation_layout.addStretch()
        no_shift_sub_tab = QWidget()
        no_shift_layout = QVBoxLayout(no_shift_sub_tab)
        no_shift_group = QGroupBox("担当者不要日")
        no_shift_inner_layout = QVBoxLayout()
        self.add_no_shift_button = QPushButton("日付を選択・編集...")
        self.no_shift_list = QListWidget()
        self.delete_no_shift_button = QPushButton("選択した日を削除")
        no_shift_inner_layout.addWidget(self.add_no_shift_button)
        no_shift_inner_layout.addWidget(self.no_shift_list)
        no_shift_inner_layout.addWidget(self.delete_no_shift_button)
        no_shift_group.setLayout(no_shift_inner_layout)
        no_shift_layout.addWidget(no_shift_group)
        no_shift_layout.addStretch()
        monthly_sub_tab_widget.addTab(fixed_shift_sub_tab, "固定シフト")
        monthly_sub_tab_widget.addTab(vacation_sub_tab, "休暇")
        monthly_sub_tab_widget.addTab(no_shift_sub_tab, "不要日")
        fairness_tab = QWidget()
        fairness_layout = QVBoxLayout(fairness_tab)
        adj_group = QGroupBox("公平性調整値")
        adj_layout = QVBoxLayout()
        adj_layout.addWidget(QLabel("特定のスタッフを優遇/冷遇する場合に設定します。\n+1: 1回分担当が少なくなるように調整"))
        self.adjustment_table = QTableWidget()
        self.adjustment_table.setColumnCount(3)
        self.adjustment_table.setHorizontalHeaderLabels(["スタッフ", "総回数", "特別日"])
        self.adjustment_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.adjustment_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.adjustment_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        adj_layout.addWidget(self.adjustment_table)
        adj_group.setLayout(adj_layout)
        fairness_layout.addWidget(adj_group)
        history_group = QGroupBox("過去のシフト履歴")
        history_layout = QVBoxLayout()
        self.history_checkbox = QCheckBox("過去2ヶ月の履歴を考慮して公平性を計算する")
        self.history_checkbox.setChecked(True)
        self.history_label = QLabel("履歴を読み込んでいます...")
        self.history_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.history_label.setStyleSheet("background-color: #f0f0f0; padding: 5px; border-radius: 3px;")
        history_layout.addWidget(self.history_checkbox)
        history_layout.addWidget(self.history_label)
        history_group.setLayout(history_layout)
        fairness_layout.addWidget(history_group)
        fairness_layout.addStretch()
        options_tab_widget.addTab(monthly_tab, "月限定")
        options_tab_widget.addTab(fairness_tab, "公平性")
        left_layout.addWidget(options_group)
        self.generate_button = QPushButton("シフトを生成！")
        self.generate_button.setStyleSheet("font-size: 16px; padding: 10px;")
        left_layout.addWidget(self.generate_button)
        left_layout.addStretch()
        splitter.addWidget(left_widget)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        result_splitter = QSplitter(Qt.Orientation.Vertical)
        solutions_group = QGroupBox("生成されたシフトパターン")
        solutions_layout = QVBoxLayout()
        self.solutions_table = QTableWidget()
        self.solutions_table.setColumnCount(2)
        self.solutions_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.solutions_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.solutions_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.solutions_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        solutions_layout.addWidget(self.solutions_table)
        solutions_group.setLayout(solutions_layout)
        result_splitter.addWidget(solutions_group)

        # --- Right bottom: tabs (Preview / History) ---
        preview_group = QGroupBox("選択したパターン / 履歴")
        preview_layout = QVBoxLayout(preview_group)

        right_tabs = QTabWidget()

        # Tab 1: Preview
        preview_tab = QWidget()
        preview_tab_layout = QVBoxLayout(preview_tab)
        preview_splitter = QSplitter(Qt.Orientation.Vertical)
        self.preview_table = QTableWidget()
        preview_splitter.addWidget(self.preview_table)
        action_widget = QWidget()
        action_layout = QHBoxLayout(action_widget)
        action_layout.setContentsMargins(0, 10, 0, 0)
        self.save_history_button = QPushButton("このシフトを履歴として保存")
        self.export_file_button = QPushButton("ファイルに出力...")
        self.save_history_button.setEnabled(False)
        self.export_file_button.setEnabled(False)
        action_layout.addWidget(self.save_history_button)
        action_layout.addStretch()
        action_layout.addWidget(self.export_file_button)
        preview_splitter.addWidget(action_widget)
        preview_splitter.setSizes([600, 120])
        preview_tab_layout.addWidget(preview_splitter)
        right_tabs.addTab(preview_tab, "プレビュー")

        # Tab 2: History
        history_tab = QWidget()
        history_tab_layout = QVBoxLayout(history_tab)
        history_splitter = QSplitter(Qt.Orientation.Vertical)
        # top: list
        history_top = QWidget()
        history_top_layout = QVBoxLayout(history_top)
        history_toolbar = QHBoxLayout()
        self.history_reload_button = QPushButton("再読込")
        self.history_open_dir_button = QPushButton("フォルダを開く")
        self.history_delete_button = QPushButton("削除")
        history_toolbar.addWidget(self.history_reload_button)
        history_toolbar.addWidget(self.history_open_dir_button)
        history_toolbar.addWidget(self.history_delete_button)
        history_toolbar.addStretch()
        history_top_layout.addLayout(history_toolbar)
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(5)
        self.history_table.setHorizontalHeaderLabels(["年月", "ファイル名", "登録日", "総回数合計", "特別日合計"])
        self.history_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.history_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.history_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.history_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.history_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        history_top_layout.addWidget(self.history_table)
        history_splitter.addWidget(history_top)
        # bottom: preview + summary
        history_bottom = QWidget()
        history_bottom_layout = QHBoxLayout(history_bottom)
        self.history_preview_table = QTableWidget()
        self.history_preview_table.setMinimumHeight(200)
        history_bottom_layout.addWidget(self.history_preview_table, 2)
        self.history_summary_table = QTableWidget()
        self.history_summary_table.setColumnCount(3)
        self.history_summary_table.setHorizontalHeaderLabels(["スタッフ", "総回数", "特別日"])
        self.history_summary_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.history_summary_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.history_summary_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        history_bottom_layout.addWidget(self.history_summary_table, 1)
        history_splitter.addWidget(history_bottom)
        history_splitter.setSizes([300, 500])
        history_tab_layout.addWidget(history_splitter)
        right_tabs.addTab(history_tab, "履歴")

        preview_layout.addWidget(right_tabs)
        result_splitter.addWidget(preview_group)
        # ★★★★★ UI変更ここまで ★★★★★
        
        right_layout.addWidget(result_splitter)
        splitter.addWidget(right_widget)
        splitter.setSizes([400, 800])

    def _connect_signals(self):
        # ... (既存のシグナル接続) ...
        self.generate_button.clicked.connect(self._start_generation)
        self.solutions_table.itemSelectionChanged.connect(self._update_preview_and_actions)
        self.year_spinbox.valueChanged.connect(self._load_and_display_history)
        self.month_combo.currentIndexChanged.connect(self._load_and_display_history)
        self.history_checkbox.stateChanged.connect(self._load_and_display_history)
        self.add_vacation_button.clicked.connect(self._add_manual_vacation)
        self.delete_vacation_button.clicked.connect(self._delete_manual_vacation)
        self.add_no_shift_button.clicked.connect(self._add_no_shift_dates)
        self.delete_no_shift_button.clicked.connect(self._delete_no_shift_dates)
        self.add_fixed_shift_button.clicked.connect(self._add_manual_fixed_shift)
        self.delete_fixed_shift_button.clicked.connect(self._delete_manual_fixed_shift)

        # ★★★★★ 新しいシグナル接続 ★★★★★
        self.save_history_button.clicked.connect(self._save_history)
        self.export_file_button.clicked.connect(self._export_file)

    def update_options_ui(self):
        # ... (変更なし) ...
        staff_list = sorted(self.settings_manager.staff_manager.get_active_staff(), key=lambda s: s.name)
        staff_names = [s.name for s in staff_list]
        self.adjustment_table.setRowCount(len(staff_list))
        for i, staff in enumerate(staff_list):
            name_item = QTableWidgetItem(staff.name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.adjustment_table.setItem(i, 0, name_item)
            total_spin_box = QSpinBox()
            total_spin_box.setRange(-20, 20)
            total_spin_box.setButtonSymbols(QSpinBox.ButtonSymbols.PlusMinus)
            self.adjustment_table.setCellWidget(i, 1, total_spin_box)
            fairness_spin_box = QSpinBox()
            fairness_spin_box.setRange(-20, 20)
            fairness_spin_box.setButtonSymbols(QSpinBox.ButtonSymbols.PlusMinus)
            self.adjustment_table.setCellWidget(i, 2, fairness_spin_box)
        all_staff_list = sorted(self.settings_manager.staff_manager.get_all_staff(), key=lambda s: s.name)
        all_staff_names = [s.name for s in all_staff_list]
        for combo in [self.vacation_staff_combo, self.fixed_shift_staff_combo]:
            current_text = combo.currentText()
            combo.clear()
            combo.addItems(all_staff_names)
            if current_text in all_staff_names:
                combo.setCurrentText(current_text)
        self._load_and_display_history()

    # ... (_add_manual_fixed_shift, _delete_manual_fixed_shift など、中間のメソッドは変更なし) ...
    def _add_manual_fixed_shift(self):
        staff_name = self.fixed_shift_staff_combo.currentText()
        if not staff_name:
            QMessageBox.warning(self, "エラー", "固定するスタッフを選択してください。")
            return
        year = self.year_spinbox.value()
        month = self.month_combo.currentIndex() + 1
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{staff_name} の固定日選択 ({year}年{month}月)")
        layout = QVBoxLayout(dialog)
        calendar_widget = QCalendarWidget()
        calendar_widget.setGridVisible(True)
        calendar_widget.setMinimumDate(QDate(year, month, 1))
        calendar_widget.setMaximumDate(QDate(year, month, 1).addMonths(1).addDays(-1))
        # クリック検出のため SingleSelection を維持（内部では複数選択を自前管理）
        calendar_widget.setSelectionMode(QCalendarWidget.SelectionMode.SingleSelection)
        calendar_widget.setSelectedDate(QDate(year, month, 1))
        add_button = QPushButton("選択した日付を追加 →")
        clear_button = QPushButton("選択の全解除")
        clear_list_button = QPushButton("右リストを全クリア")
        temp_list = QListWidget()
        # 右側リストはクリックでその行を削除できるようにする
        def _tmp_remove_item_fixed(item):
            temp_list.takeItem(temp_list.row(item))
        temp_list.itemClicked.connect(_tmp_remove_item_fixed)
        button_box = QHBoxLayout()
        ok_button = QPushButton("完了")
        cancel_button = QPushButton("キャンセル")
        button_box.addStretch()
        button_box.addWidget(ok_button)
        button_box.addWidget(cancel_button)
        calendar_layout = QHBoxLayout()
        calendar_layout.addWidget(calendar_widget)
        # ボタンを縦に配置（追加/全解除）
        button_col = QVBoxLayout()
        button_col.addWidget(add_button)
        button_col.addWidget(clear_button)
        button_col.addWidget(clear_list_button)
        button_col.addStretch()
        calendar_layout.addLayout(button_col)
        calendar_layout.addWidget(temp_list)
        layout.addLayout(calendar_layout)
        layout.addLayout(button_box)
        current_no_shift_dates = {self.no_shift_list.item(i).text() for i in range(self.no_shift_list.count())}
        for i in range(self.fixed_shift_list.count()):
            item_text = self.fixed_shift_list.item(i).text()
            if item_text.endswith(f": {staff_name}"):
                date_str = item_text.split(': ')[0]
                temp_list.addItem(date_str)
                # 初期ハイライト
                parts = [int(x) for x in date_str.split('-')]
                qd = QDate(parts[0], parts[1], parts[2])
                fmt = QTextCharFormat(); fmt.setBackground(QColor('#CCE5FF')); fmt.setFontWeight(75)
                calendar_widget.setDateTextFormat(qd, fmt)
        # クリックでトグル選択（右側リストへは即追加しない）
        selected_set = {temp_list.item(i).text() for i in range(temp_list.count())}
        def toggle_date(qdate: QDate):
            nonlocal selected_set
            selected_date_str = qdate.toString("yyyy-MM-dd")
            if selected_date_str in selected_set:
                selected_set.remove(selected_date_str)
                fmt = QTextCharFormat()
                calendar_widget.setDateTextFormat(qdate, fmt)
            else:
                selected_set.add(selected_date_str)
                fmt = QTextCharFormat(); fmt.setBackground(QColor('#CCE5FF')); fmt.setFontWeight(75)
                calendar_widget.setDateTextFormat(qdate, fmt)
        calendar_widget.clicked.connect(toggle_date)
        def clear_selection():
            nonlocal selected_set
            for dstr in list(selected_set):
                try:
                    y, m, d = [int(x) for x in dstr.split('-')]
                    qd = QDate(y, m, d)
                    fmt = QTextCharFormat()
                    calendar_widget.setDateTextFormat(qd, fmt)
                except Exception:
                    pass
            selected_set.clear()
        clear_button.clicked.connect(clear_selection)
        def clear_right_list():
            temp_list.clear()
        clear_list_button.clicked.connect(clear_right_list)
        def add_date_to_list():
            # 現在ハイライト中の選択集合をまとめて右側リストへ反映
            changed = False
            for selected_date_str in sorted(selected_set):
                if selected_date_str in current_no_shift_dates:
                    QMessageBox.warning(dialog, "ルール衝突", f"{selected_date_str} は担当者不要日に設定されているため、固定シフトは追加できません。")
                    continue
                dup = False
                for i in range(self.fixed_shift_list.count()):
                    item = self.fixed_shift_list.item(i)
                    if item.text().startswith(selected_date_str) and not item.text().endswith(f": {staff_name}"):
                        QMessageBox.warning(dialog, "重複エラー", f"{selected_date_str} は既に他のスタッフで固定されています。")
                        dup = True
                        break
                if dup:
                    continue
                if not temp_list.findItems(selected_date_str, Qt.MatchFlag.MatchExactly):
                    temp_list.addItem(selected_date_str)
                    changed = True
            if changed:
                temp_list.sortItems()
        add_button.clicked.connect(add_date_to_list)
        ok_button.clicked.connect(dialog.accept)
        cancel_button.clicked.connect(dialog.reject)
        if dialog.exec():
            for i in range(self.fixed_shift_list.count() - 1, -1, -1):
                if self.fixed_shift_list.item(i).text().endswith(f": {staff_name}"):
                    self.fixed_shift_list.takeItem(i)
            final_dates = [temp_list.item(i).text() for i in range(temp_list.count())]
            for date_str in final_dates:
                self.fixed_shift_list.addItem(f"{date_str}: {staff_name}")
            self.fixed_shift_list.sortItems()

    def _delete_manual_fixed_shift(self):
        selected_items = self.fixed_shift_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "エラー", "削除する固定シフトをリストから選択してください。")
            return
        for item in selected_items:
            self.fixed_shift_list.takeItem(self.fixed_shift_list.row(item))

    def _add_no_shift_dates(self):
        year = self.year_spinbox.value()
        month = self.month_combo.currentIndex() + 1
        dialog = QDialog(self)
        dialog.setWindowTitle(f"担当者不要日の選択 ({year}年{month}月)")
        layout = QVBoxLayout(dialog)
        calendar_widget = QCalendarWidget()
        calendar_widget.setGridVisible(True)
        calendar_widget.setMinimumDate(QDate(year, month, 1))
        calendar_widget.setMaximumDate(QDate(year, month, 1).addMonths(1).addDays(-1))
        calendar_widget.setSelectionMode(QCalendarWidget.SelectionMode.SingleSelection)
        calendar_widget.setSelectedDate(QDate(year, month, 1))
        add_button = QPushButton("選択した日付を追加 →")
        clear_button = QPushButton("選択の全解除")
        clear_list_button = QPushButton("右リストを全クリア")
        temp_list = QListWidget()
        def _tmp_remove_item_noshift(item):
            temp_list.takeItem(temp_list.row(item))
        temp_list.itemClicked.connect(_tmp_remove_item_noshift)
        button_box = QHBoxLayout()
        ok_button = QPushButton("完了")
        cancel_button = QPushButton("キャンセル")
        button_box.addStretch()
        button_box.addWidget(ok_button)
        button_box.addWidget(cancel_button)
        calendar_layout = QHBoxLayout()
        calendar_layout.addWidget(calendar_widget)
        button_col = QVBoxLayout()
        button_col.addWidget(add_button)
        button_col.addWidget(clear_button)
        button_col.addWidget(clear_list_button)
        button_col.addStretch()
        calendar_layout.addLayout(button_col)
        calendar_layout.addWidget(temp_list)
        layout.addLayout(calendar_layout)
        layout.addLayout(button_box)
        current_fixed_shifts = {}
        for i in range(self.fixed_shift_list.count()):
            text = self.fixed_shift_list.item(i).text()
            date_str, staff_name = text.split(': ')
            current_fixed_shifts[date_str] = staff_name
        for i in range(self.no_shift_list.count()):
            temp_list.addItem(self.no_shift_list.item(i).text())
            dstr = self.no_shift_list.item(i).text()
            y,m,dd = [int(x) for x in dstr.split('-')]
            qd = QDate(y, m, dd)
            fmt = QTextCharFormat(); fmt.setBackground(QColor('#FFE8CC')); fmt.setFontWeight(75)
            calendar_widget.setDateTextFormat(qd, fmt)
        selected_set = {temp_list.item(i).text() for i in range(temp_list.count())}
        def toggle_date_noshift(qdate: QDate):
            nonlocal selected_set
            selected_date_str = qdate.toString("yyyy-MM-dd")
            if selected_date_str in selected_set:
                selected_set.remove(selected_date_str)
                fmt = QTextCharFormat(); calendar_widget.setDateTextFormat(qdate, fmt)
            else:
                selected_set.add(selected_date_str)
                fmt = QTextCharFormat(); fmt.setBackground(QColor('#FFE8CC')); fmt.setFontWeight(75)
                calendar_widget.setDateTextFormat(qdate, fmt)
        calendar_widget.clicked.connect(toggle_date_noshift)
        def clear_selection_noshift():
            nonlocal selected_set
            for dstr in list(selected_set):
                try:
                    y, m, d = [int(x) for x in dstr.split('-')]
                    qd = QDate(y, m, d)
                    fmt = QTextCharFormat()
                    calendar_widget.setDateTextFormat(qd, fmt)
                except Exception:
                    pass
            selected_set.clear()
        clear_button.clicked.connect(clear_selection_noshift)
        def clear_right_list_noshift():
            temp_list.clear()
        clear_list_button.clicked.connect(clear_right_list_noshift)
        def add_date_to_list():
            changed = False
            for selected_date_str in sorted(selected_set):
                if selected_date_str in current_fixed_shifts:
                    staff_name = current_fixed_shifts[selected_date_str]
                    reply = QMessageBox.question(dialog, "ルール衝突の確認",
                                                 f"{selected_date_str} には {staff_name} の固定シフトが設定されています。\n"
                                                 "担当者不要日に設定すると、この固定シフトは無視されますがよろしいですか？",
                                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                                 QMessageBox.StandardButton.No)
                    if reply == QMessageBox.StandardButton.No:
                        continue
                if not temp_list.findItems(selected_date_str, Qt.MatchFlag.MatchExactly):
                    temp_list.addItem(selected_date_str)
                    changed = True
            if changed:
                temp_list.sortItems()
        add_button.clicked.connect(add_date_to_list)
        ok_button.clicked.connect(dialog.accept)
        cancel_button.clicked.connect(dialog.reject)
        if dialog.exec():
            self.no_shift_list.clear()
            final_dates = [temp_list.item(i).text() for i in range(temp_list.count())]
            self.no_shift_list.addItems(final_dates)

    def _delete_no_shift_dates(self):
        selected_items = self.no_shift_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "エラー", "削除する日付をリストから選択してください。")
            return
        for item in selected_items:
            self.no_shift_list.takeItem(self.no_shift_list.row(item))

    def _add_manual_vacation(self):
        staff_name = self.vacation_staff_combo.currentText()
        if not staff_name:
            QMessageBox.warning(self, "エラー", "休暇を追加するスタッフを選択してください。")
            return
        year = self.year_spinbox.value()
        month = self.month_combo.currentIndex() + 1
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{staff_name} の休暇日選択 ({year}年{month}月)")
        layout = QVBoxLayout(dialog)
        calendar_widget = QCalendarWidget()
        calendar_widget.setGridVisible(True)
        calendar_widget.setMinimumDate(QDate(year, month, 1))
        calendar_widget.setMaximumDate(QDate(year, month, 1).addMonths(1).addDays(-1))
        calendar_widget.setSelectionMode(QCalendarWidget.SelectionMode.SingleSelection)
        calendar_widget.setSelectedDate(QDate(year, month, 1))
        add_button = QPushButton("選択した日付を追加 →")
        clear_button = QPushButton("選択の全解除")
        clear_list_button = QPushButton("右リストを全クリア")
        temp_list = QListWidget()
        def _tmp_remove_item_vac(item):
            temp_list.takeItem(temp_list.row(item))
        temp_list.itemClicked.connect(_tmp_remove_item_vac)
        button_box = QHBoxLayout()
        ok_button = QPushButton("完了")
        cancel_button = QPushButton("キャンセル")
        button_box.addStretch()
        button_box.addWidget(ok_button)
        button_box.addWidget(cancel_button)
        calendar_layout = QHBoxLayout()
        calendar_layout.addWidget(calendar_widget)
        button_col = QVBoxLayout()
        button_col.addWidget(add_button)
        button_col.addWidget(clear_button)
        button_col.addWidget(clear_list_button)
        button_col.addStretch()
        calendar_layout.addLayout(button_col)
        calendar_layout.addWidget(temp_list)
        layout.addLayout(calendar_layout)
        layout.addLayout(button_box)
        for i in range(self.vacation_list.count()):
            item_text = self.vacation_list.item(i).text()
            if item_text.startswith(f"{staff_name}:"):
                dates_part = item_text.split(': ')[1]
                for d_str in dates_part.split(', '):
                    temp_list.addItem(d_str)
                    # 既存休暇をハイライト
                    try:
                        day = int(d_str.replace('日',''))
                        qd = QDate(year, month, day)
                        fmt = QTextCharFormat(); fmt.setBackground(QColor('#FFD6E7')); fmt.setFontWeight(75)
                        calendar_widget.setDateTextFormat(qd, fmt)
                    except Exception:
                        pass
        selected_days = {int(temp_list.item(i).text().replace('日','')) for i in range(temp_list.count()) if temp_list.item(i).text().endswith('日')}
        def toggle_date_vac(qdate: QDate):
            nonlocal selected_days
            day = qdate.day()
            if day in selected_days:
                selected_days.remove(day)
                fmt = QTextCharFormat(); calendar_widget.setDateTextFormat(qdate, fmt)
            else:
                selected_days.add(day)
                fmt = QTextCharFormat(); fmt.setBackground(QColor('#FFD6E7')); fmt.setFontWeight(75)
                calendar_widget.setDateTextFormat(qdate, fmt)
        calendar_widget.clicked.connect(toggle_date_vac)
        def clear_selection_vac():
            nonlocal selected_days
            for day in list(selected_days):
                try:
                    qd = QDate(year, month, day)
                    fmt = QTextCharFormat()
                    calendar_widget.setDateTextFormat(qd, fmt)
                except Exception:
                    pass
            selected_days.clear()
        clear_button.clicked.connect(clear_selection_vac)
        def clear_right_list_vac():
            temp_list.clear()
        clear_list_button.clicked.connect(clear_right_list_vac)
        def add_date_to_list():
            # 選択済み（日数）をまとめて反映
            existing = {temp_list.item(i).text() for i in range(temp_list.count())}
            for day in sorted(selected_days):
                label = f"{day}日"
                if label not in existing:
                    temp_list.addItem(label)
            items = [temp_list.item(i).text() for i in range(temp_list.count())]
            items.sort(key=lambda x: int(x.replace('日', '')))
            temp_list.clear(); temp_list.addItems(items)
        add_button.clicked.connect(add_date_to_list)
        ok_button.clicked.connect(dialog.accept)
        cancel_button.clicked.connect(dialog.reject)
        if dialog.exec():
            for i in range(self.vacation_list.count() - 1, -1, -1):
                if self.vacation_list.item(i).text().startswith(f"{staff_name}:"):
                    self.vacation_list.takeItem(i)
            final_dates = [temp_list.item(i).text() for i in range(temp_list.count())]
            if final_dates:
                dates_str = ", ".join(final_dates)
                item_text = f"{staff_name}: {dates_str}"
                self.vacation_list.addItem(item_text)

    def _delete_manual_vacation(self):
        selected_items = self.vacation_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "エラー", "削除する休暇をリストから選択してください。")
            return
        for item in selected_items:
            self.vacation_list.takeItem(self.vacation_list.row(item))

    def _load_and_display_history(self):
        # ... (変更なし) ...
        self.past_total_counts = {}
        self.past_fairness_counts = {}
        self.last_month_end_dates = {}
        self.prev_month_schedule = {}
        self.prev_month_consecutive_days = {}
        self.last_week_assignments = {}
        if not self.history_checkbox.isChecked():
            self.history_label.setText("過去の履歴は考慮しません。")
            return
        year = self.year_spinbox.value()
        month = self.month_combo.currentIndex() + 1
        target_date = datetime.date(year, month, 1)
        prev1_month_date = target_date - relativedelta(months=1)
        prev2_month_date = target_date - relativedelta(months=2)
        history1 = self.settings_manager.load_history(prev1_month_date.year, prev1_month_date.month)
        history2 = self.settings_manager.load_history(prev2_month_date.year, prev2_month_date.month)
        histories_found = []
        if history1:
            histories_found.append(f"{prev1_month_date.year}年{prev1_month_date.month}月")
            for name, count in history1.get("counts", {}).items():
                self.past_total_counts[name] = self.past_total_counts.get(name, 0) + count
            for name, count in history1.get("fairness_group_counts", {}).items():
                self.past_fairness_counts[name] = self.past_fairness_counts.get(name, 0) + count
        if history2:
            histories_found.append(f"{prev2_month_date.year}年{prev2_month_date.month}月")
            for name, count in history2.get("counts", {}).items():
                self.past_total_counts[name] = self.past_total_counts.get(name, 0) + count
            for name, count in history2.get("fairness_group_counts", {}).items():
                self.past_fairness_counts[name] = self.past_fairness_counts.get(name, 0) + count
        if history1 and "schedule" in history1:
            last_dates = {}
            temp_schedule = {}
            staff_map = {s.name: s for s in self.settings_manager.staff_manager.get_all_staff()}
            schedule_by_date = {datetime.date.fromisoformat(d["date"]): d["staff_names"] for d in history1["schedule"]}
            for day_data in history1["schedule"]:
                date_obj = datetime.date.fromisoformat(day_data["date"])
                staff_obj_list = []
                for staff_name in day_data["staff_names"]:
                    if staff_name in staff_map:
                        staff_obj_list.append(staff_map[staff_name])
                        last_dates[staff_name] = date_obj
                temp_schedule[date_obj] = staff_obj_list
            self.last_month_end_dates = last_dates
            self.prev_month_schedule = temp_schedule
            _, days_in_prev_month = calendar.monthrange(prev1_month_date.year, prev1_month_date.month)
            for staff_name in staff_map.keys():
                consecutive_days = 0
                for i in range(days_in_prev_month, 0, -1):
                    d = datetime.date(prev1_month_date.year, prev1_month_date.month, i)
                    if d in schedule_by_date and staff_name in schedule_by_date[d]:
                        consecutive_days += 1
                    else:
                        break
                if consecutive_days > 0:
                    self.prev_month_consecutive_days[staff_name] = consecutive_days
                last_day_obj = datetime.date(prev1_month_date.year, prev1_month_date.month, days_in_prev_month)
                for i in range(7):
                    d = last_day_obj - datetime.timedelta(days=i)
                    if d in schedule_by_date and staff_name in schedule_by_date[d]:
                        self.last_week_assignments[staff_name] = d.weekday()
                        break
            print(f"前月の最終勤務日情報を読み込みました: {self.last_month_end_dates}")
            print(f"前月からの連勤日数を読み込みました: {self.prev_month_consecutive_days}")
            print(f"前月最終週の担当曜日を読み込みました: {self.last_week_assignments}")
        if not histories_found:
            self.history_label.setText("過去2ヶ月の履歴が見つかりませんでした。")
        else:
            self.history_label.setText(f"読み込み成功: {', '.join(histories_found)}")

    def _start_generation(self):
        # ... (変更なし) ...
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "情報", "現在、シフトを生成中です。")
            return
        total_adjustments = {}
        fairness_adjustments = {}
        active_staff_list = self.settings_manager.staff_manager.get_active_staff()
        active_staff_names = {s.name for s in active_staff_list}
        for i in range(self.adjustment_table.rowCount()):
            name = self.adjustment_table.item(i, 0).text()
            if name in active_staff_names:
                total_spin_box = self.adjustment_table.cellWidget(i, 1)
                fairness_spin_box = self.adjustment_table.cellWidget(i, 2)
                total_adjustments[name] = total_spin_box.value()
                fairness_adjustments[name] = fairness_spin_box.value()
        previous_counts = {}
        previous_fairness_counts = {}
        if self.history_checkbox.isChecked():
            previous_counts = self.past_total_counts
            previous_fairness_counts = self.past_fairness_counts
        manual_vacations = {}
        year = self.year_spinbox.value()
        month = self.month_combo.currentIndex() + 1
        for i in range(self.vacation_list.count()):
            item_text = self.vacation_list.item(i).text()
            staff_name, dates_part = item_text.split(': ')
            dates = []
            for d_str in dates_part.split(', '):
                day = int(d_str.replace('日', ''))
                dates.append(datetime.date(year, month, day))
            manual_vacations[staff_name] = dates
        no_shift_dates = []
        for i in range(self.no_shift_list.count()):
            date_str = self.no_shift_list.item(i).text()
            no_shift_dates.append(datetime.date.fromisoformat(date_str))
        manual_fixed_shifts = {}
        all_staff = self.settings_manager.staff_manager.get_all_staff()
        staff_map = {s.name: s for s in all_staff}
        for i in range(self.fixed_shift_list.count()):
            item_text = self.fixed_shift_list.item(i).text()
            date_str, staff_name = item_text.split(': ')
            date_obj = datetime.date.fromisoformat(date_str)
            staff_obj = staff_map.get(staff_name)
            if staff_obj:
                if date_obj not in manual_fixed_shifts:
                    manual_fixed_shifts[date_obj] = []
                manual_fixed_shifts[date_obj].append(staff_obj)
        past_schedules_for_solver = {}
        if self.history_checkbox.isChecked():
            # 過去2ヶ月分の履歴を結合
            target_date = datetime.date(self.year_spinbox.value(), self.month_combo.currentIndex() + 1, 1)
            for i in range(1, 3): # 1, 2ヶ月前
                past_date = target_date - relativedelta(months=i)
                history = self.settings_manager.load_history(past_date.year, past_date.month)
                if history and "schedule" in history:
                    for day_data in history["schedule"]:
                        past_schedules_for_solver[day_data["date"]] = day_data["staff_names"]
        self.solutions_table.setRowCount(0)
        self.preview_table.setColumnCount(0)
        self.preview_table.setRowCount(0)
        self.solutions = []
        self.save_history_button.setEnabled(False) # ボタンを無効化
        self.export_file_button.setEnabled(False) # ボタンを無効化
        self.progress_dialog = QProgressDialog("シフトを生成しています...", "キャンセル", 0, 0, self)
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setWindowTitle("処理中")
        self.progress_dialog.canceled.connect(self._cancel_generation)
        self.worker = GenerationWorker(
            self.settings_manager, year, month,
            total_adjustments, fairness_adjustments, 
            previous_counts, previous_fairness_counts,
            self.last_month_end_dates,
            self.prev_month_consecutive_days,
            self.last_week_assignments,
            manual_vacations, no_shift_dates,
            manual_fixed_shifts
        )
        # GenerationWorkerのinitとrunも修正が必要
        self.worker.disperse_duties = self.settings_manager.disperse_duties
        self.worker.fairness_group = self.settings_manager.fairness_group
        self.worker.past_schedules = past_schedules_for_solver
        self.worker.finished.connect(self._on_generation_finished)
        self.worker.start()
        self.progress_dialog.show()

    def _cancel_generation(self):
        # ... (変更なし) ...
        if self.worker and self.worker.isRunning(): self.worker.terminate(); self.worker.wait(); self.worker = None; print("シフト生成がキャンセルされました。")

    def _on_generation_finished(self, solutions, error_message: str):
        # ... (変更なし) ...
        if hasattr(self, 'progress_dialog') and not self.progress_dialog.wasCanceled(): self.progress_dialog.close()
        if self.worker is None: return
        if error_message:
            QMessageBox.warning(self, "生成エラー", error_message)
            return
        self.solutions = solutions
        if solutions:
            QMessageBox.information(self, "成功", f"{len(solutions)}件のシフトパターンが見つかりました！")
            self._update_solutions_table()
        else:
            QMessageBox.warning(self, "結果なし", "条件を満たすシフトパターンが見つかりませんでした。")
        self.worker = None

    def _update_solutions_table(self):
        # ... (変更なし) ...
        fairness_group_str = ", ".join(sorted(list(self.settings_manager.fairness_group)))
        header_label = f"各スタッフの担当回数 (総回数 / {fairness_group_str} 回数)"
        self.solutions_table.setHorizontalHeaderLabels(["パターン", header_label])
        self.solutions_table.setRowCount(len(self.solutions))
        active_staff_list = sorted(self.settings_manager.staff_manager.get_active_staff(), key=lambda s: s.name)
        if not active_staff_list: return
        for i, sol_data in enumerate(self.solutions):
            pattern_item = QTableWidgetItem(f"パターン {i+1}")
            self.solutions_table.setItem(i, 0, pattern_item)
            counts_str_parts = []
            total_counts = sol_data["counts"]
            fairness_counts = sol_data["fairness_group_counts"]
            for staff in active_staff_list:
                total = total_counts.get(staff.name, 0)
                fairness_val = fairness_counts.get(staff.name, 0)
                counts_str_parts.append(f"{staff.name}: {total} / {fairness_val}")
            counts_str = ",  ".join(counts_str_parts)
            counts_item = QTableWidgetItem(counts_str)
            self.solutions_table.setItem(i, 1, counts_item)
        self.solutions_table.resizeRowsToContents()

    # ★★★★★ メソッド名を変更 ★★★★★
    def _update_preview_and_actions(self):
        selected_rows = self.solutions_table.selectionModel().selectedRows()
        if not selected_rows:
            self.save_history_button.setEnabled(False)
            self.export_file_button.setEnabled(False)
            return
        
        self._update_preview() # プレビュー更新処理を呼び出す
        self.save_history_button.setEnabled(True)
        self.export_file_button.setEnabled(True)

    def _update_preview(self):
        selected_rows = self.solutions_table.selectionModel().selectedRows()
        if not selected_rows: return

        selected_row = selected_rows[0].row()
        if selected_row >= len(self.solutions): return

        schedule_data = self.solutions[selected_row]["schedule"]
        # ... (プレビューテーブルの描画ロジックは変更なし) ...
        year = self.year_spinbox.value()
        month = self.month_combo.currentIndex() + 1
        max_staff_per_day = 0
        if schedule_data:
            max_staff_per_day = max(len(staff_list) for staff_list in schedule_data.values() if staff_list is not None)
        is_indicator_mode = (max_staff_per_day > 1)
        cal = calendar.monthcalendar(year, month)
        self.preview_table.setRowCount(len(cal))
        self.preview_table.setColumnCount(7)
        self.preview_table.setHorizontalHeaderLabels(list(weekdays_jp))
        no_shift_dates = {datetime.date.fromisoformat(self.no_shift_list.item(i).text()) for i in range(self.no_shift_list.count())}
        prev_month_date = datetime.date(year, month, 1) - relativedelta(months=1)
        _, days_in_prev_month = calendar.monthrange(prev_month_date.year, prev_month_date.month)
        first_weekday = datetime.date(year, month, 1).weekday()
        for row_idx, week in enumerate(cal):
            for col_idx, day in enumerate(week):
                item = QTableWidgetItem()
                item.setTextAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
                self.preview_table.setItem(row_idx, col_idx, item)
                if day == 0:
                    if row_idx == 0:
                        prev_month_day = days_in_prev_month - (first_weekday - 1 - col_idx)
                        date = datetime.date(prev_month_date.year, prev_month_date.month, prev_month_day)
                        staff_list = self.prev_month_schedule.get(date)
                        cell_text = f"{prev_month_day}\n"
                        if staff_list:
                            cell_text += "\n".join([s.name for s in staff_list])
                        item.setText(cell_text)
                        item.setForeground(QColor("#AAAAAA"))
                        item.setData(Qt.ItemDataRole.UserRole, staff_list)
                    continue
                date = datetime.date(year, month, day)
                staff_list = schedule_data.get(date)
                cell_text = f"{day}\n"
                if staff_list:
                    cell_text += "\n".join([s.name for s in staff_list])
                item.setText(cell_text)
                item.setData(Qt.ItemDataRole.UserRole, staff_list)
                if date in no_shift_dates:
                    item.setBackground(QColor("#CCCCCC"))
                elif staff_list:
                    if is_indicator_mode:
                        item.setBackground(QColor("white"))
                    else:
                        item.setBackground(QColor(staff_list[0].color_code))
                else:
                    item.setBackground(QColor("white"))
                jp_holidays = holidays.JP(years=year)
                is_holiday_date = date in jp_holidays
                if col_idx >= 5 or is_holiday_date:
                    item.setForeground(QColor("red"))
                else:
                    item.setForeground(QColor("black"))
        self.preview_table.resizeRowsToContents()
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    # ===== 履歴ビュー関連 =====
    def _history_dir(self) -> str:
        try:
            return self.settings_manager.history_dir
        except Exception:
            return os.path.join(os.path.expanduser('~'), 'shift_history')

    def _refresh_history_list(self):
        try:
            dirpath = self._history_dir()
            files = []
            if os.path.isdir(dirpath):
                for name in os.listdir(dirpath):
                    p = os.path.join(dirpath, name)
                    if not os.path.isfile(p):
                        continue
                    ym = None
                    if name.endswith('.json'):
                        # YYYY-MM.json
                        if len(name) == 12 and name[:4].isdigit() and name[4] == '-' and name[5:7].isdigit():
                            ym = (int(name[:4]), int(name[5:7]))
                        # history_YYYY-MM.json
                        elif name.startswith('history_') and len(name) == 20 and name[8:12].isdigit() and name[12] == '-' and name[13:15].isdigit():
                            ym = (int(name[8:12]), int(name[13:15]))
                    if ym:
                        files.append((p, ym[0], ym[1], name))
            files.sort(key=lambda x: (x[1], x[2]), reverse=True)
            self.history_table.setRowCount(len(files))
            for i, (path, y, m, name) in enumerate(files):
                try:
                    ts = os.path.getmtime(path)
                    dt = datetime.datetime.fromtimestamp(ts)
                    saved_at = dt.strftime('%Y-%m-%d %H:%M')
                except Exception:
                    saved_at = ''
                data = self.settings_manager.load_history(y, m)
                total_sum = sum((data.get('counts', {}) or {}).values()) if data else 0
                fairness_sum = sum((data.get('fairness_group_counts', {}) or {}).values()) if data else 0
                self.history_table.setItem(i, 0, QTableWidgetItem(f"{y}-{m:02d}"))
                fi = QTableWidgetItem(name)
                fi.setData(Qt.ItemDataRole.UserRole, (path, y, m))
                self.history_table.setItem(i, 1, fi)
                self.history_table.setItem(i, 2, QTableWidgetItem(saved_at))
                self.history_table.setItem(i, 3, QTableWidgetItem(str(total_sum)))
                self.history_table.setItem(i, 4, QTableWidgetItem(str(fairness_sum)))
            self.history_table.resizeRowsToContents()
        except Exception as e:
            print('history refresh error:', e)

    def _on_history_selected(self):
        try:
            rows = self.history_table.selectionModel().selectedRows()
            if not rows:
                self.history_preview_table.clear()
                self.history_summary_table.clear()
                return
            r = rows[0].row()
            info = self.history_table.item(r, 1).data(Qt.ItemDataRole.UserRole)
            if not info:
                return
            path, y, m = info
            data = self.settings_manager.load_history(y, m)
            if not data:
                return
            schedule = {}
            for d in data.get('schedule', []) or []:
                try:
                    dd = datetime.date.fromisoformat(d.get('date'))
                    schedule[dd] = d.get('staff_names', [])
                except Exception:
                    continue
            self._render_history_preview(y, m, schedule)
            self._render_history_summary(data)
        except Exception as e:
            print('history select error:', e)

    def _render_history_preview(self, year: int, month: int, schedule: dict):
        cal = calendar.monthcalendar(year, month)
        self.history_preview_table.setRowCount(len(cal))
        self.history_preview_table.setColumnCount(7)
        self.history_preview_table.setHorizontalHeaderLabels(list(weekdays_jp))
        jp_holidays = holidays.JP(years=year)
        for row_idx, week in enumerate(cal):
            for col_idx, day in enumerate(week):
                item = QTableWidgetItem()
                item.setTextAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
                if day == 0:
                    self.history_preview_table.setItem(row_idx, col_idx, item)
                    continue
                date = datetime.date(year, month, day)
                names = schedule.get(date, [])
                cell_text = f"{day}\n" + ("\n".join(names) if names else "")
                item.setText(cell_text)
                is_hol = (col_idx >= 5) or (date in jp_holidays)
                item.setForeground(QColor("red" if is_hol else "black"))
                self.history_preview_table.setItem(row_idx, col_idx, item)
        self.history_preview_table.resizeRowsToContents()
        self.history_preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    def _render_history_summary(self, data: dict):
        counts = data.get('counts', {}) or {}
        fcounts = data.get('fairness_group_counts', {}) or {}
        names = sorted(counts.keys())
        self.history_summary_table.setRowCount(len(names))
        for i, name in enumerate(names):
            self.history_summary_table.setItem(i, 0, QTableWidgetItem(name))
            self.history_summary_table.setItem(i, 1, QTableWidgetItem(str(counts.get(name, 0))))
            self.history_summary_table.setItem(i, 2, QTableWidgetItem(str(fcounts.get(name, 0))))
        self.history_summary_table.resizeRowsToContents()

    def _open_history_dir(self):
        try:
            dirpath = self._history_dir()
            if platform.system() == 'Windows':
                os.startfile(dirpath)
            else:
                webbrowser.open(dirpath)
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"フォルダを開けませんでした:\n{e}")

    def _delete_selected_history(self):
        rows = self.history_table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.information(self, "情報", "削除する履歴を選択してください。")
            return
        r = rows[0].row()
        info = self.history_table.item(r, 1).data(Qt.ItemDataRole.UserRole)
        if not info:
            return
        path, y, m = info
        reply = QMessageBox.question(self, "確認", f"{y}年{m}月の履歴を削除しますか？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
            self._refresh_history_list()
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"削除に失敗しました:\n{e}")

    # ★★★★★ 新しいメソッド ★★★★★
    def _save_history(self):
        selected_rows = self.solutions_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "エラー", "履歴として保存するシフトパターンを選択してください。")
            return
        selected_row = selected_rows[0].row()
        solution_data = self.solutions[selected_row]
        year = self.year_spinbox.value()
        month = self.month_combo.currentIndex() + 1
        
        # 履歴が既に存在するかチェック
        if self.settings_manager.history_exists(year, month):
            reply = QMessageBox.question(
                self,
                '履歴の上書き確認',
                f"{year}年{month}月の履歴は既に存在します。\n"
                "新しい内容で上書きしますか？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            # 「いいえ」が選択されたら処理を中断
            if reply == QMessageBox.StandardButton.No:
                QMessageBox.information(self, "情報", "履歴の保存をキャンセルしました。")
                return

        # ユーザーが「はい」を選択したか、履歴が存在しない場合に保存処理を実行
        success = self.settings_manager.save_history(year, month, solution_data)
        
        if success:
            QMessageBox.information(self, "成功", f"{year}年{month}月のシフトを履歴として保存しました。")
            # 履歴を保存したら、公平性計算のために再読み込みを促す
            self._load_and_display_history() 
        else:
            QMessageBox.critical(self, "エラー", "履歴の保存に失敗しました。")
        # ★★★★★ 変更ここまで ★★★★★

    # ★★★★★ 新しいメソッド ★★★★★
    def _export_file(self):
        selected_rows = self.solutions_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "エラー", "出力するシフトパターンを選択してください。")
            return
        
        dialog = OutputOptionsDialog(self)
        if not dialog.exec():
            return

        layout_type, file_format = dialog.get_options()
        selected_row = selected_rows[0].row()
        solution_data = self.solutions[selected_row]
        schedule_data = solution_data["schedule"]
        year = self.year_spinbox.value()
        month = self.month_combo.currentIndex() + 1

        if file_format == 'excel':
            self._export_excel(year, month, selected_row, schedule_data, layout_type)
        elif file_format == 'pdf':
            self._export_pdf(year, month, selected_row, schedule_data, layout_type)
            # QMessageBox.information(self, "未実装", "PDF出力機能は現在開発中です。")

    def _export_excel(self, year, month, selected_row, schedule_data, layout_type):
        # GeneralSettingsTab 側の出力先フォルダを優先的に使用
        base_dir = self.last_save_directory
        try:
            if hasattr(self, 'output_dir_provider') and self.output_dir_provider is not None:
                base_dir = self.output_dir_provider.get_output_directory() or base_dir
        except Exception:
            pass
        self.last_save_directory = base_dir

        # シンプルな既定名に変更し、同名があれば (2), (3), ... を付与
        base_name = f"{year}_{month:02d}"
        ext = ".xlsx"
        default_path = os.path.join(base_dir, base_name + ext)
        if os.path.exists(default_path):
            i = 2
            while True:
                candidate = os.path.join(base_dir, f"{base_name} ({i}){ext}")
                if not os.path.exists(candidate):
                    default_path = candidate
                    break
                i += 1
        
        filepath, _ = QFileDialog.getSaveFileName(self, "Excelファイルを保存", default_path, "Excel Workbook (*.xlsx)")

        if not filepath:
            return

        # 上書き防止: 同名が存在する場合は (2), (3)... を付与
        filepath = self._ensure_unique_path(filepath)

        # 保存に成功したら、ディレクトリを記憶する
        self.last_save_directory = os.path.dirname(filepath)
            
        success, error_msg = export_to_excel(
            filepath, year, month, 
            self.settings_manager.excel_title,
            schedule_data, 
            self.settings_manager.staff_manager,
            self.prev_month_schedule,
            format_type=layout_type
        )
        if success:
            QMessageBox.information(self, "成功", f"Excelファイルを出力しました。\n{filepath}")
            try:
                if platform.system() == "Windows": os.startfile(os.path.realpath(filepath))
                else: webbrowser.open(os.path.realpath(filepath))
            except Exception as e:
                QMessageBox.warning(self, "ファイルオープンエラー", f"ファイルを開けませんでした:\n{e}")
        else:
            QMessageBox.critical(self, "Excel出力エラー", f"Excelファイルの出力中にエラーが発生しました:\n{error_msg}")

    def _export_pdf(self, year, month, selected_row, schedule_data, layout_type):
        base_dir = self.last_save_directory
        try:
            if hasattr(self, 'output_dir_provider') and self.output_dir_provider is not None:
                base_dir = self.output_dir_provider.get_output_directory() or base_dir
        except Exception:
            pass
        self.last_save_directory = base_dir

        base_name = f"{year}_{month:02d}"
        ext = ".pdf"
        default_path = os.path.join(base_dir, base_name + ext)
        if os.path.exists(default_path):
            i = 2
            while True:
                candidate = os.path.join(base_dir, f"{base_name} ({i}){ext}")
                if not os.path.exists(candidate):
                    default_path = candidate
                    break
                i += 1
        filepath, _ = QFileDialog.getSaveFileName(self, "PDFファイルを保存", default_path, "PDF Document (*.pdf)")

        if not filepath:
            return

        # 上書き防止: 同名が存在する場合は (2), (3)... を付与
        filepath = self._ensure_unique_path(filepath)

        self.last_save_directory = os.path.dirname(filepath)
            
        success, error_msg = export_to_pdf(
            filepath, year, month, 
            self.settings_manager.excel_title,
            schedule_data, 
            self.settings_manager.staff_manager,
            format_type=layout_type,
            prev_month_schedule=self.prev_month_schedule 
        )
        if success:
            QMessageBox.information(self, "成功", f"PDFファイルを出力しました。\n{filepath}")
            try:
                if platform.system() == "Windows": os.startfile(os.path.realpath(filepath))
                else: webbrowser.open(os.path.realpath(filepath))
            except Exception as e:
                QMessageBox.warning(self, "ファイルオープンエラー", f"ファイルを開けませんでした:\n{e}")
        else:
            QMessageBox.critical(self, "PDF出力エラー", f"PDFファイルの出力中にエラーが発生しました:\n{error_msg}")

    def _ensure_unique_path(self, path: str) -> str:
        """指定パスが既存なら ' (2)', ' (3)' を拡張子前に付与して重複回避する。"""
        try:
            base, ext = os.path.splitext(path)
            if not os.path.exists(path):
                return path
            i = 2
            while True:
                candidate = f"{base} ({i}){ext}"
                if not os.path.exists(candidate):
                    return candidate
                i += 1
        except Exception:
            return path

    def set_output_dir_provider(self, provider):
        """Basic settings tab (GeneralSettingsTab) を渡すためのフック。"""
        self.output_dir_provider = provider
