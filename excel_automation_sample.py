import time
import logging
from excel_automation_helper import ExcelUIAutomation
from excel_automation_configs import ExcelConfig

# ログ設定
logger = logging.getLogger(__name__)

def main():
    """メイン実行例"""
    excel_auto = ExcelUIAutomation()
    
    try:
        print("Excelを起動中...")
        # Excelを起動
        if excel_auto.start_excel("demo.xlsx"):
            print("Excelが正常に起動しました")
            
            print("===== セル操作のデモ =====")
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
            
            print("===== リボン操作のデモ =====")
            print("データタブをクリック中...")
            excel_auto.click_ribbon_shortcut("A")

            print("ホーム > 中央揃えを実行中...")
            excel_auto.select_cell(0, 2)  # C1
            excel_auto.click_ribbon_shortcut("H>AC")

            print("ファイルを保存中...")
            # ファイルを保存
            excel_auto.save_file()

            print("===== ダイアログ処理機能のデモ =====")

            print("名前の定義ダイアログを表示中...")
            excel_auto.click_ribbon_shortcut("M>M>D")

            # 単一のダイアログタイトルパターンをチェック
            dialog_found, dialog_window = excel_auto.is_dialog_present("新しい名前")
            if dialog_found:
                print("新しい名前という名前のダイアログが表示されています。キャンセルします")
                excel_auto.handle_dialog("新しい名前", "{ESC}")

            print("セルA2にテキストを入力中...")
            # セルA2にテキストを入力
            excel_auto.select_cell(1, 0)  # A2
            excel_auto.input_text("保存ダイアログの表示確認用")

            print("ワークブックを閉じます")
            excel_auto.close_workbook()

            # 保存ダイアログの処理
            dialog_sequence = [
                {'title_patterns': ['保存の確認', 'Save As', 'Microsoft Excel'], 'key_action': 's'}, #保存ダイアログでsキーを押下
                {'title_patterns': ['エラー', 'Error'], 'key_action': '{ENTER}'}  #複数のダイアログが連続して出る場合はこのように記述
            ]
            print("保存ダイアログが表示されている場合、ファイルを保存します")
            excel_auto.wait_and_handle_dialogs(dialog_sequence)
            
            print("処理が完了しました")
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
        # Excelを終了
        excel_auto.exit_excel()
        pass

if __name__ == "__main__":
    main()
