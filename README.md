# Excel UI自動化サンプル

WindowsでExcel操作をUI自動化するためのPythonプロジェクトです。Pywinautoを使用してExcelのUI操作を自動化します。

## 機能

### 基本機能
- Excelの起動・終了
- ファイルの開く・保存
- セルの選択・入力
- テキスト・数式の入力
- リボン操作（短縮キー形式）
- ダイアログ処理
- ウィンドウアクティベーション

### 主要なメソッド
- `start_excel(file_path)` - Excelを起動
- `select_cell(row, column)` - セルを選択
- `input_text(text)` - テキストを入力
- `click_ribbon_shortcut(shortcut)` - リボン操作
- `save_file()` - ファイルを保存
- `handle_dialog(title_patterns, action)` - ダイアログ処理
- `exit_excel()` - Excelを終了

## セットアップ

### 1. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 2. 必要なソフトウェア
- Windows 10/11
- Microsoft Excel
- Python 3.7以上

## 使用方法

### 基本的な使用例

```python
from utils.excel_automation_helper import ExcelAutomationHelper

# Excel自動化オブジェクトを作成
excel_auto = ExcelAutomationHelper()

# Excelを起動
excel_auto.start_excel("templates/demo.xlsx")

# セルA1にテキストを入力
excel_auto.select_cell(0, 0)  # A1
excel_auto.input_text("Hello Excel!")

# リボン操作
excel_auto.click_ribbon_shortcut("H>AC")  # ホーム > 中央揃え

# ファイルを保存
excel_auto.save_file()

# Excelを閉じる
excel_auto.exit_excel()
```

### ダイアログ処理の例

```python
# 単一ダイアログの処理
dialog_found, dialog_window = excel_auto.wait_for_dialog("保存の確認", timeout=10)
if dialog_found:
    excel_auto.handle_dialog("保存の確認", "no")

# 複数ダイアログの一括処理
dialog_configs = [
    {'title_patterns': ['保存の確認', 'Save As'], 'key_action': 's'},
    {'title_patterns': ['エラー', 'Error'], 'key_action': '{ENTER}'}
]
excel_auto.wait_and_handle_dialogs(dialog_configs)
```

## 実行方法

```bash
python excel_automation_sample.py
```

## リボン操作の短縮キー

### タブ
- `H` - ホーム
- `N` - 挿入
- `A` - データ
- `M` - 数式
- `R` - 校閲

### よく使用される操作
- `H>B` - ホーム > 太字
- `H>AC` - ホーム > 中央揃え
- `H>I` - ホーム > 斜体
- `H>U` - ホーム > 下線
- `N>CH` - 挿入 > グラフ

## 注意事項

1. **Excelのバージョン**: Microsoft Excel 2016以上を推奨
2. **権限**: 管理者権限が必要な場合があります
3. **画面解像度**: 高解像度ディスプレイでは調整が必要な場合があります
4. **タイミング**: システムの性能によって待機時間の調整が必要な場合があります

## トラブルシューティング

1. **Excelが起動しない**
   - Excelがインストールされているか確認
   - パスが正しく設定されているか確認

2. **セル選択がうまくいかない**
   - システムの性能を考慮して待機時間を調整

3. **リボン操作が動作しない**
   - 短縮キーが正しく指定されているか確認

4. **ダイアログ処理がうまくいかない**
   - ダイアログのタイトルパターンが正しく指定されているか確認
   - タイムアウト時間を適切に設定

## プロジェクト構造

```
ExcelUIAutomation/
├── excel_automation_sample.py    # 実行サンプル
├── requirements.txt              # 依存関係
├── README.md                     # このファイル
├── templates/
│   └── demo.xlsx                 # サンプルファイル
└── utils/
    ├── excel_automation_helper.py    # メイン機能
    └── excel_automation_configs.py   # 設定ファイル
```
