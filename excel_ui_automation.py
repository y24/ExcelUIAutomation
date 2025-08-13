import time
import os
import winreg
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import logging
from config import ExcelConfig, EnvironmentConfig

# 環境設定を取得
env_config = EnvironmentConfig.get_config()

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
            send_keys(text)
            time.sleep(ExcelConfig.get_timing('text_input'))
            send_keys('{ENTER}')
            logger.info(f"テキストを入力しました: {text}")
            return True
            
        except Exception as e:
            logger.error(f"テキスト入力エラー: {e}")
            return False

    def click_ribbon_shortcut(self, shortcut_key):
        """短縮キー形式でリボン操作を実行（例: "H>AC" でホームタブの中央揃え）"""
        try:
            # Altキーでリボンにアクセス
            send_keys('%')
            time.sleep(ExcelConfig.get_timing('text_input'))
            
            # 短縮キーの形式を解析
            if '>' in shortcut_key:
                parts = shortcut_key.split('>')
                if len(parts) == 2:
                    tab_shortcut = parts[0].strip()
                    button_shortcut = parts[1].strip()
                    
                    # タブの短縮キーを送信
                    send_keys(tab_shortcut.upper())
                    time.sleep(ExcelConfig.get_timing('ribbon_operation'))
                    
                    # ボタンの短縮キーを送信
                    send_keys(button_shortcut.upper())
                    time.sleep(ExcelConfig.get_timing('ribbon_operation'))
                    
                    logger.info(f"リボン短縮キー '{shortcut_key}' を実行しました")
                    return True
                else:
                    logger.error(f"無効な短縮キー形式: {shortcut_key}")
                    return False
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
                # 正常にExcelを閉じる（Ctrl+W でワークブックを閉じる）
                send_keys(ExcelConfig.get_shortcut('close_workbook'))
                time.sleep(ExcelConfig.get_timing('file_operation'))
                logger.info("ワークブックを閉じました")
                
                # 保存確認ダイアログが表示された場合の処理
                try:
                    # 保存確認ダイアログの「保存しない」を選択
                    send_keys('n')  # No
                    time.sleep(ExcelConfig.get_timing('dialog_wait'))
                    logger.info("保存確認ダイアログを閉じました")
                except:
                    pass
                
                # さらに確実に閉じるため、Alt+F4 を使用
                send_keys('%{F4}')
                time.sleep(ExcelConfig.get_timing('file_operation'))
                
                # それでも残っている場合は強制終了
                try:
                    if self.app.is_process_running():
                        self.app.kill()
                        logger.info("Excelを強制終了しました")
                    else:
                        logger.info("Excelを正常に閉じました")
                except:
                    self.app.kill()
                    logger.info("Excelを強制終了しました")
                
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
            
            print("セルA1にテキストを入力中...")
            # セルA1にテキストを入力
            excel_auto.select_cell(0, 0)  # A1
            excel_auto.input_text("Hello Excel!")
            
            print("セルB1に数式を入力中...")
            # セルB1に数式を入力
            excel_auto.select_cell(0, 1)  # B1
            excel_auto.input_text("=A1")
            
            print("セルC1に数値を入力中...")
            # セルC1に数値を入力
            excel_auto.select_cell(0, 2)  # C1
            excel_auto.input_text("1000")
            
            # ===== リボン操作のデモ =====
            print("データタブをクリック中...")
            excel_auto.click_ribbon_shortcut("A")

            print("ホーム > 中央揃えを実行中...")
            excel_auto.select_cell(0, 2)  # C1
            excel_auto.click_ribbon_shortcut("H>AC")
            
            print("ファイルを保存中...")
            # ファイルを保存
            excel_auto.save_file()
            
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
        excel_auto.close_excel()

if __name__ == "__main__":
    main() 