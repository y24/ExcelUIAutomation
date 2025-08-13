import time
import os
import winreg
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto.findwindows import find_window, find_windows
import logging
from config import ExcelConfig, EnvironmentConfig

# 環境設定を取得
env_config = EnvironmentConfig.get_config()

# ログファイルのクリーンアップ（スクリプト実行ごと）
def cleanup_log_file():
    """スクリプト実行ごとにログファイルをクリーンアップ"""
    try:
        log_file_path = ExcelConfig.LOGGING['file']
        if os.path.exists(log_file_path):
            os.remove(log_file_path)
            print(f"前回のログファイルを削除しました: {log_file_path}")
    except Exception as e:
        print(f"ログファイル削除エラー（無視可能）: {e}")

# ログファイルをクリーンアップ
cleanup_log_file()

# ログ設定
logging.basicConfig(
    level=getattr(logging, ExcelConfig.LOGGING['level']), 
    format=ExcelConfig.LOGGING['format'],
    filename=ExcelConfig.LOGGING['file'],
    encoding='utf-8'  # UTF-8エンコーディングを明示的に指定
)
logger = logging.getLogger(__name__)

def get_excel_path():
    """レジストリからExcelのインストールパスを取得"""
    try:
        # Office 2016以降（App Paths）
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe") as key:
            path, _ = winreg.QueryValueEx(key, "")
            if os.path.exists(path):
                logger.info(f"レジストリからExcelパスを取得: {path}")
                return path
    except Exception as e:
        logger.debug(f"App Pathsからの取得に失敗: {e}")
    
    try:
        # Office 2016/2019/365
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot") as key:
            path, _ = winreg.QueryValueEx(key, "Path")
            excel_path = os.path.join(path, "EXCEL.EXE")
            if os.path.exists(excel_path):
                logger.info(f"レジストリからExcelパスを取得: {excel_path}")
                return excel_path
    except Exception as e:
        logger.debug(f"Office 16.0からの取得に失敗: {e}")
    
    logger.warning("レジストリからExcelパスを取得できませんでした")
    return None

class ExcelUIAutomation:
    def __init__(self):
        self.app = None
        self.excel_window = None
        self.workbook = None
        self.copied_files = []  # コピーしたファイルのパスを記録
        
    def wait_for_dialog(self, title_patterns, timeout=None, check_interval=None):
        """
        指定されたタイトルパターンに一致するダイアログが表示されるまで待機
        
        Args:
            title_patterns (str or list): ダイアログタイトルのパターン（文字列またはリスト）
            timeout (float): 最大待機時間（秒）（Noneの場合は設定ファイルの値を使用）
            check_interval (float): チェック間隔（秒）（Noneの場合は設定ファイルの値を使用）
            
        Returns:
            tuple: (ダイアログが見つかったかどうか, ダイアログウィンドウオブジェクト)
        """
        try:
            # タイトルパターンをリストに統一
            if isinstance(title_patterns, str):
                title_patterns = [title_patterns]
            
            # 設定ファイルからデフォルト値を取得
            if timeout is None:
                timeout = ExcelConfig.get_timing('dialog_timeout', 10)
            if check_interval is None:
                check_interval = ExcelConfig.get_timing('dialog_check_interval', 0.5)
            
            logger.info(f"ダイアログの表示を待機中... (タイムアウト: {timeout}秒, パターン: {title_patterns})")
            
            start_time = time.time()
            while time.time() - start_time < timeout:
                try:
                    # 各タイトルパターンでダイアログを検索
                    for pattern in title_patterns:
                        try:
                            dialog_window = find_window(title_re=f".*{pattern}.*")
                            if dialog_window:
                                logger.info(f"ダイアログを検出しました: {pattern}")
                                return True, dialog_window
                        except Exception as e:
                            logger.debug(f"ダイアログ検索エラー（パターン: {pattern}）: {e}")
                            continue
                    
                    # より汎用的な検索（Excel関連のダイアログ）
                    try:
                        excel_dialogs = find_windows(title_re=".*Microsoft Excel.*")
                        for dialog in excel_dialogs:
                            dialog_title = dialog.window_text()
                            logger.debug(f"検出されたダイアログ: {dialog_title}")
                            if any(pattern.lower() in dialog_title.lower() for pattern in title_patterns):
                                logger.info(f"ダイアログを検出しました: {dialog_title}")
                                return True, dialog
                    except Exception as e:
                        logger.debug(f"汎用ダイアログ検索エラー: {e}")
                    
                    time.sleep(check_interval)
                    
                except Exception as e:
                    logger.debug(f"ダイアログ待機中のエラー: {e}")
                    time.sleep(check_interval)
            
            logger.warning(f"ダイアログの表示待機がタイムアウトしました (パターン: {title_patterns})")
            return False, None
            
        except Exception as e:
            logger.error(f"ダイアログ待機エラー: {e}")
            return False, None
    
    def is_dialog_present(self, title_patterns):
        """
        指定されたタイトルパターンに一致するダイアログが現在表示されているかチェック
        
        Args:
            title_patterns (str or list): ダイアログタイトルのパターン（文字列またはリスト）
            
        Returns:
            tuple: (ダイアログが表示されているかどうか, ダイアログウィンドウオブジェクト)
        """
        try:
            # タイトルパターンをリストに統一
            if isinstance(title_patterns, str):
                title_patterns = [title_patterns]
            
            # 各タイトルパターンでダイアログを検索
            for pattern in title_patterns:
                try:
                    dialog_window = find_window(title_re=f".*{pattern}.*")
                    if dialog_window:
                        logger.info(f"ダイアログが表示されています: {pattern}")
                        return True, dialog_window
                except Exception as e:
                    logger.debug(f"ダイアログ検索エラー（パターン: {pattern}）: {e}")
                    continue
            
            # より汎用的な検索
            try:
                excel_dialogs = find_windows(title_re=".*Microsoft Excel.*")
                for dialog in excel_dialogs:
                    dialog_title = dialog.window_text()
                    logger.debug(f"検出されたダイアログ: {dialog_title}")
                    if any(pattern.lower() in dialog_title.lower() for pattern in title_patterns):
                        logger.info(f"ダイアログが表示されています: {dialog_title}")
                        return True, dialog
            except Exception as e:
                logger.debug(f"汎用ダイアログ検索エラー: {e}")
            
            return False, None
            
        except Exception as e:
            logger.error(f"ダイアログ存在確認エラー: {e}")
            return False, None
    
    def handle_dialog(self, title_patterns, action='ok', timeout=10):
        """
        ダイアログを処理する（表示を待機してから適切なアクションを実行）
        
        Args:
            title_patterns (str or list): ダイアログタイトルのパターン（文字列またはリスト）
            action (str): 実行するアクション（'ok', 'cancel', 'yes', 'no', 'enable'）
            timeout (float): ダイアログ表示待機時間（秒）
            
        Returns:
            bool: 処理が成功したかどうか
        """
        try:
            # ダイアログの表示を待機
            dialog_found, dialog_window = self.wait_for_dialog(title_patterns, timeout)
            
            if not dialog_found:
                logger.info(f"ダイアログは表示されませんでした (パターン: {title_patterns})")
                return True  # ダイアログが表示されない場合は成功とみなす
            
            # アクションを実行
            logger.info(f"ダイアログでアクション '{action}' を実行")
            
            # ダイアログをアクティブにする
            try:
                if dialog_window:
                    dialog_window.set_focus()
                    time.sleep(ExcelConfig.get_timing('dialog_wait', 0.5))
            except Exception as e:
                logger.debug(f"ダイアログアクティベートエラー: {e}")
            
            # ダイアログが完全に表示されるまで少し待機
            time.sleep(ExcelConfig.get_timing('dialog_wait'))
            
            # アクションに応じたキーを送信（キーを直接渡す）
            send_keys(action)
            
            time.sleep(ExcelConfig.get_timing('dialog_wait', 0.5))
            logger.info(f"ダイアログの処理が完了しました")
            return True
                
        except Exception as e:
            logger.error(f"ダイアログ処理エラー: {e}")
            return False
    
    def wait_and_handle_dialogs(self, dialog_configs, timeout=10):
        """
        複数のダイアログ設定を順次チェックして処理
        
        Args:
            dialog_configs (list): ダイアログ設定のリスト
                [{'title_patterns': ['パターン1', 'パターン2'], 'action': 'ok'}, ...]
            timeout (float): 各ダイアログの待機時間（秒）
            
        Returns:
            bool: すべての処理が成功したかどうか
        """
        try:
            success = True
            for config in dialog_configs:
                title_patterns = config.get('title_patterns', [])
                action = config.get('action', 'ok')
                
                # ダイアログの表示を待機してから処理
                if not self.handle_dialog(title_patterns, action, timeout):
                    success = False
                    logger.warning(f"ダイアログの処理に失敗しました (パターン: {title_patterns})")
            
            return success
            
        except Exception as e:
            logger.error(f"複数ダイアログ処理エラー: {e}")
            return False

    def activate_excel_window(self, max_retries=3, retry_delay=1.0):
        """
        Excelウィンドウをアクティベートする汎用的なメソッド
        
        Args:
            max_retries (int): 最大リトライ回数（デフォルト: 3）
            retry_delay (float): リトライ間隔（秒）（デフォルト: 1.0）
            
        Returns:
            bool: アクティベートに成功したかどうか
        """
        try:
            if not self.app or not self.excel_window:
                logger.warning("Excelアプリケーションまたはウィンドウが初期化されていません")
                return False
            
            for attempt in range(max_retries):
                try:
                    logger.info(f"Excelウィンドウのアクティベートを試行中... (試行 {attempt + 1}/{max_retries})")
                    
                    # 方法1: pywinautoのset_focus()を使用
                    try:
                        self.excel_window.set_focus()
                        time.sleep(ExcelConfig.get_timing('window_activation'))
                        logger.info("pywinautoのset_focus()でExcelウィンドウをアクティベートしました")
                        return True
                    except Exception as e:
                        logger.debug(f"set_focus()でのアクティベートに失敗: {e}")
                    
                    # 方法2: ウィンドウハンドルを使用してアクティベート
                    try:
                        import win32gui
                        import win32con
                        
                        # ウィンドウハンドルを取得
                        hwnd = self.excel_window.handle
                        if hwnd:
                            # ウィンドウを前面に表示
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                            time.sleep(ExcelConfig.get_timing('window_activation'))
                            
                            # ウィンドウをアクティブにする
                            win32gui.SetForegroundWindow(hwnd)
                            time.sleep(ExcelConfig.get_timing('window_activation'))
                            
                            logger.info("win32guiを使用してExcelウィンドウをアクティベートしました")
                            return True
                    except Exception as e:
                        logger.debug(f"win32guiでのアクティベートに失敗: {e}")
                    
                    # 方法3: Alt+Tabを使用してExcelウィンドウに切り替え
                    try:
                        # Alt+Tabでウィンドウを切り替え
                        send_keys('%{TAB}')
                        time.sleep(ExcelConfig.get_timing('window_activation'))
                        
                        # さらに確実にするため、Altキーを押してリリース
                        send_keys('%')
                        time.sleep(ExcelConfig.get_timing('window_activation'))
                        
                        logger.info("Alt+Tabを使用してExcelウィンドウをアクティベートしました")
                        return True
                    except Exception as e:
                        logger.debug(f"Alt+Tabでのアクティベートに失敗: {e}")
                    
                    # 方法4: ウィンドウタイトルで検索してアクティベート
                    try:
                        import win32gui
                        import win32con
                        
                        def enum_windows_callback(hwnd, target_title):
                            if win32gui.IsWindowVisible(hwnd):
                                window_title = win32gui.GetWindowText(hwnd)
                                if target_title.lower() in window_title.lower():
                                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                                    time.sleep(ExcelConfig.get_timing('window_activation'))
                                    win32gui.SetForegroundWindow(hwnd)
                                    time.sleep(ExcelConfig.get_timing('window_activation'))
                                    return False  # 列挙を停止
                            return True
                        
                        # Excelウィンドウを検索してアクティベート
                        win32gui.EnumWindows(enum_windows_callback, "Excel")
                        logger.info("ウィンドウタイトル検索でExcelウィンドウをアクティベートしました")
                        return True
                    except Exception as e:
                        logger.debug(f"ウィンドウタイトル検索でのアクティベートに失敗: {e}")
                    
                    # リトライ前の待機
                    if attempt < max_retries - 1:
                        logger.info(f"アクティベートに失敗しました。{retry_delay}秒後にリトライします...")
                        time.sleep(retry_delay)
                    
                except Exception as e:
                    logger.debug(f"アクティベート試行 {attempt + 1} でエラー: {e}")
                    if attempt < max_retries - 1:
                        time.sleep(retry_delay)
            
            logger.warning(f"Excelウィンドウのアクティベートに失敗しました（{max_retries}回試行）")
            return False
            
        except Exception as e:
            logger.error(f"Excelウィンドウアクティベートエラー: {e}")
            return False
    
    def ensure_excel_active(self, operation_name="操作"):
        """
        操作前にExcelウィンドウがアクティブであることを保証するヘルパーメソッド
        
        Args:
            operation_name (str): 実行する操作の名前（ログ用）
            
        Returns:
            bool: Excelウィンドウがアクティブになったかどうか
        """
        try:
            logger.info(f"{operation_name}の前にExcelウィンドウをアクティベート中...")
            if self.activate_excel_window():
                logger.info(f"{operation_name}の準備が完了しました")
                return True
            else:
                logger.warning(f"{operation_name}の準備に失敗しましたが、操作を続行します")
                return False
        except Exception as e:
            logger.error(f"Excelウィンドウアクティベート確認エラー: {e}")
            return False

    def start_excel(self, file_path=None):
        """Excelを起動し、指定されたファイルを開く"""
        try:
            # 起動前に復旧ファイルを削除
            self._cleanup_recovery_files()
            
            # ファイルが指定されている場合、信頼できる場所にコピー
            if file_path and os.path.exists(file_path):
                # デスクトップにコピーして保護ビューを回避
                import shutil
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                safe_file_path = os.path.join(desktop_path, os.path.basename(file_path))
                shutil.copy2(file_path, safe_file_path)
                # コピーしたファイルのパスを記録
                self.copied_files.append(safe_file_path)
                file_path = safe_file_path
                logger.info(f"ファイルを信頼できる場所にコピーしました: {safe_file_path}")
            
            # レジストリからExcelのパスを取得
            excel_path = get_excel_path()
            
            # Excelの一般的なインストールパスを試す
            excel_paths = [
                excel_path,  # レジストリから取得したパスを最初に試す
                r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
                r"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE",
                r"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE",
                r"C:\Program Files\Microsoft Office\root\Office15\EXCEL.EXE",
                r"C:\Program Files (x86)\Microsoft Office\root\Office15\EXCEL.EXE"
            ]
            
            # 有効なパスを見つける
            valid_excel_path = None
            for path in excel_paths:
                if path is None:
                    continue
                try:
                    if os.path.exists(path) or path == "excel.exe":
                        valid_excel_path = path
                        break
                except:
                    continue
            
            if valid_excel_path is None:
                logger.error("Excelが見つかりません。Excelがインストールされているか確認してください。")
                return False
            
            # Excelを起動
            if file_path and os.path.exists(file_path):
                # 既存のファイルを開く（保護ビューを無効にするオプション付き）
                cmd = f'"{valid_excel_path}" "{file_path}" /e'
                self.app = Application().start(cmd)
                logger.info(f"Excelファイルを開きました: {file_path}")
            else:
                # 新しいExcelを起動
                self.app = Application().start(valid_excel_path)
                logger.info("新しいExcelを起動しました")
            
            # Excelウィンドウが表示されるまで待機
            time.sleep(ExcelConfig.get_timing('excel_startup'))
            
            # メインウィンドウを取得
            self.excel_window = self.app.window(title_re=ExcelConfig.get_excel_setting('window_title_pattern'))
            self.excel_window.wait('visible', timeout=ExcelConfig.get_timing('window_wait'))
            
            return True
            
        except Exception as e:
            logger.error(f"Excel起動エラー: {e}")
            logger.error("詳細なエラー情報:")
            import traceback
            traceback.print_exc()
            return False
    
    def open_file(self, file_path):
        """ファイルを開く"""
        try:
            # Excelウィンドウをアクティベート
            self.ensure_excel_active("ファイルを開く")
            
            # Ctrl+O でファイルを開く
            send_keys(ExcelConfig.get_shortcut('open_file'))
            time.sleep(ExcelConfig.get_timing('file_operation'))
            
            # ファイルパスを入力
            send_keys(file_path)
            time.sleep(ExcelConfig.get_timing('text_input'))
            
            # Enter で開く
            send_keys('{ENTER}')
            time.sleep(ExcelConfig.get_timing('file_operation'))
            
            logger.info(f"ファイルを開きました: {file_path}")
            return True
            
        except Exception as e:
            logger.error(f"ファイルを開くエラー: {e}")
            return False
    
    def save_file(self, file_path=None):
        """ファイルを保存"""
        try:
            # Excelウィンドウをアクティベート
            self.ensure_excel_active("ファイル保存")
            
            if file_path:
                # Ctrl+Shift+S で名前を付けて保存
                send_keys(ExcelConfig.get_shortcut('save_as'))
                time.sleep(ExcelConfig.get_timing('file_operation'))
                send_keys(file_path)
                time.sleep(ExcelConfig.get_timing('text_input'))
                send_keys('{ENTER}')
            else:
                # Ctrl+S で保存
                send_keys(ExcelConfig.get_shortcut('save_file'))
            
            time.sleep(ExcelConfig.get_timing('file_operation'))
            logger.info("ファイルを保存しました")
            return True
            
        except Exception as e:
            logger.error(f"ファイル保存エラー: {e}")
            return False
    
    def select_cell(self, row, column):
        """セルを選択"""
        try:
            # Excelウィンドウをアクティベート
            self.ensure_excel_active("セル選択")
            
            # セルに移動
            cell_address = ExcelConfig.get_cell_address(row, column)
            send_keys(ExcelConfig.get_shortcut('go_to'))  # Ctrl+G でジャンプ
            time.sleep(ExcelConfig.get_timing('cell_selection'))
            send_keys(cell_address)
            time.sleep(ExcelConfig.get_timing('cell_selection'))
            send_keys('{ENTER}')
            time.sleep(ExcelConfig.get_timing('cell_selection'))
            
            logger.info(f"セル {cell_address} を選択しました")
            return True
            
        except Exception as e:
            logger.error(f"セル選択エラー: {e}")
            return False
    
    def input_text(self, text):
        """テキストを入力"""
        try:
            # Excelウィンドウをアクティベート
            self.ensure_excel_active("テキスト入力")
            
            send_keys(text)
            time.sleep(ExcelConfig.get_timing('text_input'))
            send_keys('{ENTER}')
            logger.info(f"テキストを入力しました: {text}")
            return True
            
        except Exception as e:
            logger.error(f"テキスト入力エラー: {e}")
            return False

    def click_ribbon_shortcut(self, shortcut_key):
        """短縮キー形式でリボン操作を実行（例: "H>AC" でホームタブの中央揃え、"M>M>D" で数式タブ>名前の定義>名前の定義）"""
        try:
            # Excelウィンドウをアクティベート
            self.ensure_excel_active("リボン操作")
            
            # Altキーでリボンにアクセス
            send_keys('%')
            time.sleep(ExcelConfig.get_timing('text_input'))
            
            # 短縮キーの形式を解析
            if '>' in shortcut_key:
                parts = [part.strip().upper() for part in shortcut_key.split('>')]
                
                # 各段階の短縮キーを順次送信
                for i, key in enumerate(parts):
                    send_keys(key)
                    time.sleep(ExcelConfig.get_timing('ribbon_operation'))
                
                logger.info(f"リボン短縮キー '{shortcut_key}' を実行しました")
                return True
            else:
                # タブのみの短縮キーの場合
                send_keys(shortcut_key.upper())
                time.sleep(ExcelConfig.get_timing('ribbon_operation'))
                # タブキー送信後、Enterキーで抜ける
                send_keys('{ENTER}')
                time.sleep(ExcelConfig.get_timing('ribbon_operation'))
                logger.info(f"リボンタブ短縮キー '{shortcut_key}' を実行しました")
                return True
                    
        except Exception as e:
            logger.error(f"リボン短縮キー実行エラー: {e}")
            return False

    def close_dialog(self):
        """ダイアログを閉じる"""
        try:
            # Excelウィンドウをアクティベート
            self.ensure_excel_active("ダイアログを閉じる")
            
            send_keys('{ESC}')
            time.sleep(ExcelConfig.get_timing('dialog_wait'))
            logger.info("ダイアログを閉じました")
            return True
            
        except Exception as e:
            logger.error(f"ダイアログ閉じるエラー: {e}")
            return False

    def close_excel(self):
        """Excelを閉じる"""
        try:
            if self.app:
                # Excelウィンドウをアクティベート
                self.ensure_excel_active("Excelを閉じる")
                
                # 正常にExcelを閉じる（Ctrl+W でワークブックを閉じる）
                send_keys(ExcelConfig.get_shortcut('close_workbook'))
                time.sleep(ExcelConfig.get_timing('file_operation'))
                logger.info("ワークブックを閉じました")
                
                # さらに確実に閉じるため、Alt+F4 を使用
                # send_keys('%{F4}')
                # time.sleep(ExcelConfig.get_timing('file_operation'))
                
                # それでも残っている場合は強制終了
                # try:
                #     if self.app.is_process_running():
                #         self.app.kill()
                #         logger.info("Excelを終了しました")
                #     else:
                #         pass
                # except:
                #     self.app.kill()
                #     logger.info("Excelを終了しました")
                
                # 復旧ファイルを削除
                self._cleanup_recovery_files()
                
            return True
            
        except Exception as e:
            logger.error(f"Excel終了エラー: {e}")
            # エラーが発生した場合は強制終了
            try:
                if self.app:
                    self.app.kill()
            except:
                pass
            # エラーが発生しても復旧ファイルは削除
            self._cleanup_recovery_files()
            # エラー時はコピーしたファイルもクリーンナップ
            self._cleanup_copied_files()
            return False
    
    def _cleanup_recovery_files(self):
        """復旧ファイルを削除"""
        try:
            import glob
            # Excelの復旧ファイルの一般的な場所
            recovery_paths = [
                os.path.expanduser("~/AppData/Local/Microsoft/Office/UnsavedFiles"),
                os.path.expanduser("~/AppData/Roaming/Microsoft/Excel"),
            ]
            
            # 復旧ファイルのパターン（より具体的に）
            recovery_patterns = [
                "*.xlsx~*",
                "*.xls~*",
                "*[Recovered]*",
                "*~$*.xlsx",
                "*~$*.xls"
            ]
            
            # デスクトップの復旧ファイルのみを削除（より厳密に）
            desktop_path = os.path.expanduser("~/Desktop")
            if os.path.exists(desktop_path):
                # デスクトップでは、より具体的な復旧ファイルパターンのみを対象とする
                desktop_recovery_patterns = [
                    "*[オリジナル].xlsx",
                    "*[オリジナル].xls",
                    "*[Recovered].xlsx",
                    "*[Recovered].xls",
                    "*~$*.xlsx",
                    "*~$*.xls"
                ]
                
                for pattern in desktop_recovery_patterns:
                    files = glob.glob(os.path.join(desktop_path, pattern))
                    for file_path in files:
                        try:
                            # コピーしたファイルは削除しない
                            if file_path in self.copied_files:
                                logger.debug(f"コピーしたファイルのため削除をスキップ: {file_path}")
                                continue
                            
                            # ファイル名をチェックして、通常のファイルでないことを確認
                            file_name = os.path.basename(file_path)
                            if any(keyword in file_name for keyword in ["[オリジナル]", "[Recovered]", "~$"]):
                                os.remove(file_path)
                                logger.info(f"デスクトップの復旧ファイルを削除しました: {file_path}")
                        except Exception as e:
                            logger.debug(f"デスクトップ復旧ファイル削除エラー（無視可能）: {e}")
            
            # その他の場所の復旧ファイルを削除
            for recovery_path in recovery_paths:
                if os.path.exists(recovery_path):
                    for pattern in recovery_patterns:
                        files = glob.glob(os.path.join(recovery_path, pattern))
                        for file_path in files:
                            try:
                                os.remove(file_path)
                                logger.info(f"復旧ファイルを削除しました: {file_path}")
                            except Exception as e:
                                logger.debug(f"復旧ファイル削除エラー（無視可能）: {e}")
                                
        except Exception as e:
            logger.debug(f"復旧ファイル削除エラー（無視可能）: {e}")
    
    def _cleanup_copied_files(self):
        """コピーしたファイルをクリーンアップ"""
        try:
            for file_path in self.copied_files:
                if os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                        logger.info(f"コピーしたファイルを削除しました: {file_path}")
                    except Exception as e:
                        logger.debug(f"コピーしたファイル削除エラー（無視可能）: {e}")
            # リストをクリア
            self.copied_files.clear()
        except Exception as e:
            logger.debug(f"コピーしたファイルクリーンアップエラー（無視可能）: {e}")

def main():
    """メイン実行例"""
    excel_auto = ExcelUIAutomation()
    
    try:
        print("Excelを起動中...")
        # Excelを起動
        if excel_auto.start_excel(ExcelConfig.get_excel_setting('default_file')):
            print("Excelが正常に起動しました")
            time.sleep(ExcelConfig.get_timing('excel_startup'))
            
            # ===== セル操作のデモ =====
            # print("セルA1にテキストを入力中...")
            # # セルA1にテキストを入力
            # excel_auto.select_cell(0, 0)  # A1
            # excel_auto.input_text("Hello Excel!")
            
            # print("セルB1に数式を入力中...")
            # # セルB1に数式を入力
            # excel_auto.select_cell(0, 1)  # B1
            # excel_auto.input_text("=A1")
            
            # print("セルC1に数値を入力中...")
            # # セルC1に数値を入力
            # excel_auto.select_cell(0, 2)  # C1
            # excel_auto.input_text("1000")
            
            # ===== リボン操作のデモ =====
            # print("データタブをクリック中...")
            # excel_auto.click_ribbon_shortcut("A")

            # print("ホーム > 中央揃えを実行中...")
            # excel_auto.select_cell(0, 2)  # C1
            # excel_auto.click_ribbon_shortcut("H>AC")
            
            # ===== ダイアログ処理機能のデモ =====
            print("ダイアログ処理機能のデモ...")

            # print("名前の定義ダイアログを表示中...")
            # excel_auto.click_ribbon_shortcut("M>M>D")

            # 単一のダイアログタイトルパターンをチェック
            if excel_auto.is_dialog_present("新しい名前"):
                print("名前の定義ダイアログが表示されています。キャンセルします")
                excel_auto.handle_dialog("新しい名前", "{ESC}")
            
            print("ファイルを保存中...")
            # ファイルを保存
            excel_auto.save_file()

            print("セルA2にテキストを入力中...")
            # セルA2にテキストを入力
            excel_auto.select_cell(1, 0)  # A2
            excel_auto.input_text("保存ダイアログの表示")

            print("Excelを閉じます")
            excel_auto.close_excel()

            # # 複数のダイアログ設定を一括処理
            dialog_sequence = [
                {'title_patterns': ['保存の確認', 'Save As', 'Microsoft Excel'], 'action': 's'},
                {'title_patterns': ['エラー', 'Error'], 'action': '{ENTER}'}
            ]
            print("保存ダイアログが表示されているため、ファイルを保存します")
            excel_auto.wait_and_handle_dialogs(dialog_sequence, timeout=15)
            
            print("処理が完了しました")
            # 少し待ってから閉じる
            time.sleep(ExcelConfig.get_timing('excel_startup'))
        else:
            print("Excelの起動に失敗しました")
            
    except Exception as e:
        logger.error(f"実行エラー: {e}")
        print("詳細なエラー情報:")
        import traceback
        traceback.print_exc()
        print("\nExcelがインストールされているか、または正しく起動できるか確認してください。")
        print("また、管理者権限で実行してみてください。")
    
    finally:
        # excel_auto.close_excel()
        pass

if __name__ == "__main__":
    main() 