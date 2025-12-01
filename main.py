# main.py

import sys
import json
import os
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QFileDialog, QMessageBox
)
from PySide6.QtGui import QAction, QCloseEvent
from PySide6.QtCore import QRect, QStandardPaths 

from core_engine import SettingsManager
from staff_config_tab import StaffConfigTab
from rule_config_tab import RuleConfigTab
from general_settings_tab import GeneralSettingsTab
from generation_tab import GenerationTab


class MainWindow(QMainWindow):
    APP_CONFIG_FILE = "app_config.json"

    def __init__(self):
        super().__init__()
        self.setWindowTitle("シフト表自動作成アプリ")

        desired_width = 1200
        desired_height = 800
        screen = self.screen().availableGeometry()
        width = min(desired_width, int(screen.width() * 0.95))
        height = min(desired_height, int(screen.height() * 0.90))
        self.resize(width, height)
        self.setMinimumSize(width, height)

        # 1. アプリケーションのデータ保存用パスを取得
        # 例: C:/Users/ユーザー名/AppData/Roaming/YourOrganizationName/ShiftGenerator
        # このパスは __main__ ブロックで設定された情報に基づきます
        self.app_data_path = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppDataLocation)
        
        # フォルダが存在しない場合は作成
        if not os.path.exists(self.app_data_path):
            os.makedirs(self.app_data_path)

        # 2. アプリ設定ファイルのパスを絶対パスで定義
        self.APP_CONFIG_FILE = os.path.join(self.app_data_path, "app_config.json")
        
        # 3. SettingsManagerに履歴保存用の絶対パスを渡す
        history_path = os.path.join(self.app_data_path, "shift_history")
        # SettingsManagerの__init__内でフォルダがなければ作成される
        self.settings_manager = SettingsManager(history_dir=history_path)

        self.current_filepath = None
        # ★★★★★ 変更点 1: last_save_directoryプロパティを初期化 ★★★★★
        # デフォルトはユーザーのホームディレクトリ
        self.last_save_directory = os.path.expanduser("~") 

        self._create_menu()
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.tabs.currentChanged.connect(self.on_tab_changed)

        self.staff_tab = StaffConfigTab(self.settings_manager)
        self.rule_tab = RuleConfigTab(self.settings_manager)
        self.general_settings_tab = GeneralSettingsTab(self.settings_manager)
        self.generation_tab = GenerationTab(self.settings_manager)
        self.generation_tab.set_output_dir_provider(self.general_settings_tab)

        self.tabs.addTab(self.staff_tab, "スタッフ設定")
        self.tabs.addTab(self.rule_tab, "ルール設定")
        self.tabs.addTab(self.general_settings_tab, "基本設定")
        self.tabs.addTab(self.generation_tab, "シフト生成と結果")

        self._load_app_config()
        self._center_window()

    def _center_window(self):
        # ... (変更なし) ...
        screen_geometry = self.screen().availableGeometry()
        frame_rect = self.frameGeometry()
        screen_center = screen_geometry.center()
        frame_rect.moveCenter(screen_center)
        y_offset = 0
        new_top_left = frame_rect.topLeft()
        self.move(new_top_left.x(), max(0, new_top_left.y() - y_offset))

    def on_tab_changed(self, index):
        # ... (変更なし) ...
        self.rule_tab.update_staff_list() 
        self.generation_tab.update_options_ui()

    def _create_menu(self):
        # ... (変更なし) ...
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("ファイル")
        load_action = QAction("設定を読み込む...", self)
        load_action.triggered.connect(self._load_settings)
        file_menu.addAction(load_action)
        save_action = QAction("設定を保存", self)
        save_action.triggered.connect(self._save_settings)
        file_menu.addAction(save_action)
        save_as_action = QAction("名前を付けて設定を保存...", self)
        save_as_action.triggered.connect(self._save_settings_as)
        file_menu.addAction(save_as_action)
        file_menu.addSeparator()
        exit_action = QAction("終了", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def _load_settings_from_path(self, filepath: str):
        # ▼▼▼▼▼ このメソッドを再度、全面的に書き換え ▼▼▼▼▼
        if not os.path.exists(filepath):
            QMessageBox.warning(self, "エラー", f"設定ファイルが見つかりません:\n{filepath}")
            self.current_filepath = None
            return

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # 1. 読み込んだデータから、新しいSettingsManagerインスタンスを生成
            new_settings = SettingsManager.from_dict(data)

            # 2. ★重要★ 現在のインスタンスから、安全な履歴パスを引き継ぐ
            new_settings.history_dir = self.settings_manager.history_dir
            
            # 3. MainWindowが保持するインスタンスを、新しいものに置き換える
            self.settings_manager = new_settings
            
            # 4. 各タブに、新しいインスタンスを再設定する
            self.staff_tab.set_settings_manager(self.settings_manager)
            self.rule_tab.set_settings_manager(self.settings_manager)
            self.general_settings_tab.set_settings_manager(self.settings_manager)
            self.generation_tab.set_settings_manager(self.settings_manager)
            
            # 5. UIに新しい設定を反映させる (set_settings_managerの中でloadが呼ばれるが、念のため明示的に呼ぶ)
            self.staff_tab.load_staff_list()
            self.rule_tab.load_rules()
            self.general_settings_tab.load_settings()
            self.generation_tab.update_options_ui()

            # 6. ウィンドウの状態を更新
            self.current_filepath = filepath
            self.setWindowTitle(f"シフト表自動作成アプリ - {os.path.basename(filepath)}")
            print(f"設定 '{filepath}' を読み込み、UIを更新しました。")

        except Exception as e:
            QMessageBox.critical(self, "エラー", f"設定ファイルの読み込みに失敗しました。\nファイルが破損しているか、形式が不正です。\n詳細: {e}")
            self.current_filepath = None
            self.setWindowTitle("シフト表自動作成アプリ")
        # ▲▲▲▲▲ ここまで書き換え ▲▲▲▲▲

    def _load_settings(self):
        # ... (変更なし) ...
        filepath, _ = QFileDialog.getOpenFileName(self, "設定ファイルを開く", "", "JSON Files (*.json)")
        if filepath:
            self._load_settings_from_path(filepath)

    def _save_settings(self):
        # ... (変更なし) ...
        if not self.current_filepath:
            self._save_settings_as()
        else:
            self.settings_manager.save_to_json(self.current_filepath)
            QMessageBox.information(self, "成功", f"設定を保存しました。\n{self.current_filepath}")

    def _save_settings_as(self):
        # ... (変更なし) ...
        filepath, _ = QFileDialog.getSaveFileName(self, "設定を名前を付けて保存", "", "JSON Files (*.json)")
        if filepath:
            self.current_filepath = filepath
            self._save_settings()
            self.setWindowTitle(f"シフト表自動作成アプリ - {os.path.basename(filepath)}")

    def _load_app_config(self):
        if os.path.exists(self.APP_CONFIG_FILE):
            try:
                with open(self.APP_CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                # ★★★★★ 変更点 2: 保存場所のパスを読み込む ★★★★★
                last_dir = config.get("last_save_directory")
                if last_dir and os.path.isdir(last_dir):
                    self.last_save_directory = last_dir
                # GenerationTab / GeneralSettingsTab にパスを渡す
                self.generation_tab.last_save_directory = self.last_save_directory
                try:
                    self.general_settings_tab.set_output_directory(self.last_save_directory)
                except Exception:
                    pass
                
                last_file = config.get("last_opened_file")
                if last_file and os.path.exists(last_file):
                    print(f"前回終了時の設定ファイル '{last_file}' を読み込みます。")
                    self._load_settings_from_path(last_file)
                else:
                    self._reset_ui_to_default()
            except (json.JSONDecodeError, KeyError):
                self._reset_ui_to_default()
        else:
            self._reset_ui_to_default()

    def _save_app_config(self):
        # self.APP_CONFIG_FILE を保存する前に、親ディレクトリが存在することを確認
        app_config_dir = os.path.dirname(self.APP_CONFIG_FILE)
        if not os.path.exists(app_config_dir):
            os.makedirs(app_config_dir)

        # ★★★★★ 変更点 3: 保存場所のパスを保存する ★★★★★
        # GeneralSettingsTab の出力先設定を優先して取得
        try:
            self.last_save_directory = self.general_settings_tab.get_output_directory()
            # GenerationTab 側にも反映
            self.generation_tab.last_save_directory = self.last_save_directory
        except Exception:
            # 互換: 取得できない場合は GenerationTab から
            self.last_save_directory = self.generation_tab.last_save_directory
        config = {
            "last_opened_file": self.current_filepath,
            "last_save_directory": self.last_save_directory
        }
        try:
            with open(self.APP_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
            print(f"アプリ設定を '{self.APP_CONFIG_FILE}' に保存しました。")
        except Exception as e:
            print(f"アプリ設定の保存中にエラーが発生しました: {e}")

    def _reset_ui_to_default(self):
        # SettingsManagerを再生成する際も、正しい履歴パスを指定する
        history_path = os.path.join(self.app_data_path, "shift_history")
        self.settings_manager = SettingsManager(history_dir=history_path)        
        self.current_filepath = None
        self.staff_tab.set_settings_manager(self.settings_manager)
        self.rule_tab.set_settings_manager(self.settings_manager)
        self.general_settings_tab.set_settings_manager(self.settings_manager)
        self.generation_tab.set_settings_manager(self.settings_manager)
        self.staff_tab.load_staff_list()
        self.rule_tab.load_rules()
        self.general_settings_tab.load_settings()
        self.generation_tab.update_options_ui()
        self.setWindowTitle("シフト表自動作成アプリ")
        print("UIを初期状態にリセットしました。")

    def closeEvent(self, event: QCloseEvent):
        self._save_app_config()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # QStandardPathsが正しく動作するために、アプリ情報を設定する
    app.setOrganizationName("YourOrganizationName") # あなたの組織名や名前に変更
    app.setApplicationName("ShiftGenerator")

    window = MainWindow()
    window.show()
    sys.exit(app.exec())