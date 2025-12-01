# hooks/hook-holidays.py

from PyInstaller.utils.hooks import collect_data_files

# holidays.countries 以下の全ファイルをデータファイルとして収集する
datas = collect_data_files('holidays.countries', include_py_files=True)