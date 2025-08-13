#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel UI自動化の設定ファイル
タイミング、パス、その他の設定を管理
"""

import os

class ExcelConfig:
    """Excel自動化の設定クラス"""
    
    # タイミング設定（秒）
    TIMING = {
        'excel_startup': 3,      # Excel起動待機時間
        'window_wait': 10,       # ウィンドウ表示待機時間
        'window_activation': 0.5, # ウィンドウアクティベーション待機時間
        'cell_selection': 0.5,   # セル選択待機時間
        'text_input': 0.5,       # テキスト入力待機時間
        'file_operation': 1,     # ファイル操作待機時間
        'dialog_wait': 1,        # ダイアログ待機時間
        'dialog_check_interval': 0.5, # ダイアログチェック間隔
        'dialog_timeout': 10,    # ダイアログ待機タイムアウト
        'ribbon_operation': 1, # リボン操作待機時間
    }
    
    # Excel関連設定
    EXCEL = {
        'process_name': 'excel.exe',
        'window_title_pattern': '.*Excel.*'
    }
    
    # キーボードショートカット
    SHORTCUTS = {
        'open_file': '^o',           # Ctrl+O
        'save_file': '^s',           # Ctrl+S
        'save_as': '^+s',            # Ctrl+Shift+S
        'close_workbook': '^w',      # Ctrl+W
        'find': '^f',                # Ctrl+F
        'select_all': '^a',          # Ctrl+A
        'go_to': '^g',               # Ctrl+G
        'insert_row': '^+{+}',       # Ctrl+Shift++
        'delete_row': '^-',          # Ctrl+-
    }
    
    # セル参照設定
    CELL_REFERENCE = {
        'start_column': 'A',
        'max_columns': 26,  # A-Z
        'max_rows': 1000,
    }
    
    # ログ設定
    LOGGING = {
        'level': 'DEBUG',
        'format': '%(asctime)s - %(levelname)s - %(message)s',
        'file': 'excel_automation.log',
    }
    
    # エラーハンドリング設定
    ERROR_HANDLING = {
        'max_retries': 3,
        'retry_delay': 1,
        'continue_on_error': True,
    }
    
    @classmethod
    def get_timing(cls, key, default=None):
        """タイミング設定を取得"""
        if default is None:
            return cls.TIMING.get(key, 1.0)
        return cls.TIMING.get(key, default)
    
    @classmethod
    def get_shortcut(cls, key):
        """ショートカットキーを取得"""
        return cls.SHORTCUTS.get(key, '')
    
    @classmethod
    def get_ribbon_tab_key(cls, tab_name):
        """リボンタブキーを取得"""
        return cls.RIBBON_TABS.get(tab_name, '')
    
    @classmethod
    def get_ribbon_button_key(cls, button_name):
        """リボンボタンキーを取得"""
        return cls.RIBBON_BUTTONS.get(button_name, '')
    
    @classmethod
    def get_excel_setting(cls, key):
        """Excel設定を取得"""
        return cls.EXCEL.get(key, '')
    
    @classmethod
    def update_timing(cls, key, value):
        """タイミング設定を更新"""
        if key in cls.TIMING:
            cls.TIMING[key] = value
    
    @classmethod
    def get_cell_address(cls, row, column):
        """セルアドレスを生成"""
        if column < 0 or column >= cls.CELL_REFERENCE['max_columns']:
            raise ValueError(f"列番号が範囲外です: {column}")
        
        column_letter = chr(65 + column)
        row_number = row + 1
        return f"{column_letter}{row_number}"
    
    @classmethod
    def get_range_address(cls, start_row, start_col, end_row, end_col):
        """範囲アドレスを生成"""
        start_cell = cls.get_cell_address(start_row, start_col)
        end_cell = cls.get_cell_address(end_row, end_col)
        return f"{start_cell}:{end_cell}"
