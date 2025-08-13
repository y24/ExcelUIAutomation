import logging
from utils.excel_automation_helper import ExcelAutomationHelper

def main():
    """実装サンプル"""
    excel = ExcelAutomationHelper()
    
    try:
        print("Excelを起動中...")
        if excel.start_excel("templates/demo.xlsx"):
            print("Excelが正常に起動しました")
            
            print("===== セル操作のデモ =====")

            print("セルA1にテキストを入力中...")
            excel.select_cell(0, 0)  # A1
            excel.input_text("Hello Excel!")
            
            print("セルB1に数式を入力中...")
            excel.select_cell(0, 1)  # B1
            excel.input_text("=A1")
            
            print("セルC1に数値を入力中...")
            excel.select_cell(0, 2)  # C1
            excel.input_text("1000")
            
            print("===== リボン操作のデモ =====")
            print("データタブをクリック中...")
            excel.click_ribbon_shortcut("A")

            print("ホーム > 中央揃えを実行中...")
            excel.select_cell(0, 2)  # C1
            excel.click_ribbon_shortcut("H>AC")

            print("ファイルを保存中...")
            excel.save_file()

            print("===== ダイアログ処理機能のデモ =====")

            print("名前の定義ダイアログを表示中...")
            excel.click_ribbon_shortcut("M>M>D")
            
            # ダイアログを待機して処理する例
            print("'新しい名前'ダイアログが表示されるまで待機します...")
            dialog_found, dialog_window = excel.wait_for_dialog("新しい名前", timeout=10)
            if dialog_found:
                print("'新しい名前'ダイアログが表示されました。キャンセルします")
                excel.handle_dialog("新しい名前", "{ESC}")
            else:
                print("タイムアウト: ダイアログが表示されませんでした")

            print("セルA2にテキストを入力中...")
            excel.select_cell(1, 0)  # A2
            excel.input_text("保存ダイアログの表示確認用")

            print("ワークブックを閉じます")
            excel.close_workbook()

            # 複数ダイアログを処理する例
            dialog_sequence = [
                {'title_patterns': ['保存の確認', 'Save As', 'Microsoft Excel'], 'key_action': 's'}, #保存ダイアログでsキーを押下
                # {'title_patterns': ['エラー', 'Error'], 'key_action': '{ENTER}'}  #複数のダイアログが連続して出る場合はこのように記述
            ]
            print("保存確認ダイアログの保存ボタンを実行します")
            excel.wait_and_handle_dialogs(dialog_sequence)
            
            print("デモが完了しました")
        else:
            print("Excelの起動に失敗しました")
            
    except Exception as e:
        # ログ設定
        logger = logging.getLogger(__name__)
        logger.error(f"実行エラー: {e}")
        print("Traceback:")
        import traceback
        traceback.print_exc()
    
    finally:
        print("Excelを終了します")
        excel.exit_excel()

if __name__ == "__main__":
    main()
