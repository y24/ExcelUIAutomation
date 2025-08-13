import logging
from utils.excel_automation_helper import ExcelAutomationHelper

def main():
    """メイン実行例"""
    excel = ExcelAutomationHelper()
    
    try:
        print("Excelを起動中...")
        # Excelを起動
        if excel.start_excel("templates/demo.xlsx"):
            print("Excelが正常に起動しました")
            
            print("===== セル操作のデモ =====")
            print("セルA1にテキストを入力中...")
            # セルA1にテキストを入力
            excel.select_cell(0, 0)  # A1
            excel.input_text("Hello Excel!")
            
            print("セルB1に数式を入力中...")
            # セルB1に数式を入力
            excel.select_cell(0, 1)  # B1
            excel.input_text("=A1")
            
            print("セルC1に数値を入力中...")
            # セルC1に数値を入力
            excel.select_cell(0, 2)  # C1
            excel.input_text("1000")
            
            print("===== リボン操作のデモ =====")
            print("データタブをクリック中...")
            excel.click_ribbon_shortcut("A")

            print("ホーム > 中央揃えを実行中...")
            excel.select_cell(0, 2)  # C1
            excel.click_ribbon_shortcut("H>AC")

            print("ファイルを保存中...")
            # ファイルを保存
            excel.save_file()

            print("===== ダイアログ処理機能のデモ =====")

            print("名前の定義ダイアログを表示中...")
            excel.click_ribbon_shortcut("M>M>D")

            # 単一のダイアログタイトルパターンをチェック
            dialog_found, dialog_window = excel.is_dialog_present("新しい名前")
            if dialog_found:
                print("新しい名前という名前のダイアログが表示されています。キャンセルします")
                excel.handle_dialog("新しい名前", "{ESC}")

            print("セルA2にテキストを入力中...")
            # セルA2にテキストを入力
            excel.select_cell(1, 0)  # A2
            excel.input_text("保存ダイアログの表示確認用")

            print("ワークブックを閉じます")
            excel.close_workbook()

            # 保存ダイアログの処理
            dialog_sequence = [
                {'title_patterns': ['保存の確認', 'Save As', 'Microsoft Excel'], 'key_action': 's'}, #保存ダイアログでsキーを押下
                {'title_patterns': ['エラー', 'Error'], 'key_action': '{ENTER}'}  #複数のダイアログが連続して出る場合はこのように記述
            ]
            print("保存ダイアログが表示されている場合、ファイルを保存します")
            excel.wait_and_handle_dialogs(dialog_sequence)
            
            print("処理が完了しました")
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
        # Excelを終了
        print("Excelを終了します")
        excel.exit_excel()
        pass

if __name__ == "__main__":
    main()
