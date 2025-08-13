# Excel UI自動化サンプル

WindowsでExcel操作をUI自動化するためのPythonプロジェクトです。Pywinautoを使用してExcelのUI操作を自動化します。

## 機能

### 基本機能 (`excel_automation_helper.py`)
- Excelの起動・終了
- ファイルの開く・保存
- セルの選択・入力
- テキスト・数式の入力
- コピー・ペースト
- セルの書式設定（太字、斜体、下線、通貨形式）
- **ウィンドウアクティベーション機能**
  - 他のウィンドウがアクティブになっている場合でも、Excelウィンドウを自動的にアクティベート
  - 複数の方法でウィンドウアクティベーションを試行（pywinauto、win32gui、Alt+Tab、ウィンドウタイトル検索）
  - リトライ機能付きで確実なアクティベーション
- **リボン操作機能**
  - リボンタブのクリック（ホーム、挿入、データ、数式、校閲など）
  - リボンボタンのクリック（太字、中央揃え、グラフ、ピボットテーブルなど）
  - **短縮キー形式でのリボン操作** - Configファイルに追記不要
  - リボンダイアログの開閉（セルの書式設定、条件付き書式など）
  - リボンギャラリーの使用（スタイル、フォント、図形など）
  - リボンコマンドパスでの実行（タブ > グループ > コマンド）
- **柔軟なダイアログ処理機能（新機能）**
  - ダイアログタイトルパターンを直接指定（複数指定可）
  - 保存確認、ファイル上書き確認、エラー、保護ビューなどのダイアログを自動処理
  - 複数ダイアログの一括処理
  - カスタマイズ可能なダイアログ設定

### 高度な機能 (`advanced_excel_automation.py`)
- 既存のExcelプロセスへの接続
- 範囲選択
- 行・列の挿入・削除
- データのソート・フィルター
- グラフの作成
- 印刷プレビュー
- 検索と置換
- シートの保護
- コメントの追加

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
from excel_automation_helper import ExcelUIAutomation

# Excel自動化オブジェクトを作成
excel_auto = ExcelUIAutomation()

# Excelを起動
excel_auto.start_excel("demo.xlsx")

# セルA1にテキストを入力
excel_auto.select_cell(0, 0)  # A1
excel_auto.input_text("Hello Excel!")

# ファイルを保存
excel_auto.save_file()

# リボン操作の例
excel_auto.click_ribbon_tab("ホーム")
excel_auto.click_ribbon_button("太字", "ホーム")
excel_auto.click_ribbon_tab("挿入")
excel_auto.click_ribbon_button("グラフ", "挿入")

# Excelを閉じる
excel_auto.close_excel()
```

### 柔軟なダイアログ処理機能の使用例

```python
from excel_automation_helper import ExcelUIAutomation

excel_auto = ExcelUIAutomation()

# 単一のダイアログタイトルパターンをチェック
if excel_auto.is_dialog_present("保存の確認"):
    print("保存確認ダイアログが表示されています")
    excel_auto.handle_dialog("保存の確認", "no")

# 複数のダイアログタイトルパターンをチェック
dialog_found, dialog_window = excel_auto.wait_for_dialog(["エラー", "Error", "警告"], timeout=5)
if dialog_found:
    print("エラーダイアログを検出しました")
    excel_auto.handle_dialog(["エラー", "Error", "警告"], "ok")

# 複数のダイアログ設定を一括処理
dialog_configs = [
    {'title_patterns': ['保存の確認', 'Save As'], 'action': 'no'},
    {'title_patterns': ['ファイルの上書き確認', 'File Already Exists'], 'action': 'yes'},
    {'title_patterns': ['保護されたビュー', 'Protected View'], 'action': 'enable'},
    {'title_patterns': ['エラー', 'Error'], 'action': 'ok'}
]
excel_auto.wait_and_handle_dialogs(dialog_configs)

# カスタムアクションでのダイアログ処理
excel_auto.handle_dialog("保存の確認", "yes")  # 保存する
excel_auto.handle_dialog("ファイルの上書き確認", "no")  # 上書きしない
```

### 高度な使用例

```python
from advanced_excel_automation import AdvancedExcelUIAutomation

# 高度なExcel自動化オブジェクトを作成
excel_auto = AdvancedExcelUIAutomation()

# 既存のExcelに接続
if excel_auto.connect_to_existing_excel():
    # データを入力
    excel_auto.select_cell(0, 0)
    excel_auto.input_text("商品名")
    
    # 範囲を選択してフィルターを適用
    excel_auto.select_range(0, 0, 10, 2)
    excel_auto.filter_data()
    
    # グラフを作成
    excel_auto.create_chart("column")
```

## 実行方法

### 基本スクリプトの実行
```bash
python excel_automation_helper.py
```

### 高度なスクリプトの実行
```bash
python advanced_excel_automation.py
```

## 注意事項

1. **Excelのバージョン**: Microsoft Excel 2016以上を推奨
2. **権限**: 管理者権限が必要な場合があります
3. **画面解像度**: 高解像度ディスプレイでは調整が必要な場合があります
4. **タイミング**: システムの性能によって`time.sleep()`の値を調整してください
5. **エラーハンドリング**: 実行中にエラーが発生した場合は、ログを確認してください
6. **ウィンドウアクティベーション**: 他のウィンドウがアクティブになっている場合でも、自動的にExcelウィンドウをアクティベートします

## トラブルシューティング

### よくある問題

1. **Excelが起動しない**
   - Excelがインストールされているか確認
   - パスが正しく設定されているか確認

2. **セル選択がうまくいかない**
   - `time.sleep()`の値を増やす
   - システムの性能を考慮して調整

3. **キーボードショートカットが動作しない**
   - Excelの言語設定を確認
   - 他のアプリケーションがキーボードを占有していないか確認

4. **ウィンドウアクティベーションがうまくいかない**
   - Excelウィンドウが最小化されていないか確認
   - 他のアプリケーションがウィンドウをブロックしていないか確認
   - 必要に応じて`activate_excel_window()`メソッドのリトライ回数を増やす

5. **ダイアログ処理がうまくいかない**
   - ダイアログのタイトルパターンが正しく指定されているか確認
   - タイムアウト時間を適切に設定
   - ダイアログが表示されるまで十分な待機時間を確保
   - 複数のタイトルパターンを指定して柔軟性を高める

### ログの確認

スクリプト実行時には詳細なログが出力されます。エラーが発生した場合は、ログメッセージを確認してください。

## カスタマイズ

### リボン操作の使用例

```python
# 基本的なリボン操作
excel_auto.click_ribbon_tab("ホーム")
excel_auto.click_ribbon_tab("挿入")
excel_auto.click_ribbon_tab("データ")

# リボンボタンのクリック
excel_auto.click_ribbon_button("太字", "ホーム")
excel_auto.click_ribbon_button("中央揃え", "ホーム")
excel_auto.click_ribbon_button("グラフ", "挿入")

# 新しい短縮キー形式（推奨）
excel_auto.click_ribbon_shortcut("H>B")      # ホームタブの太字
excel_auto.click_ribbon_shortcut("H>AC")     # ホームタブの中央揃え
excel_auto.click_ribbon_shortcut("N>CH")     # 挿入タブのグラフ
excel_auto.click_ribbon_shortcut("H")        # ホームタブのみクリック

# 短縮キー一覧の表示
excel_auto.show_ribbon_shortcuts()

# リボンダイアログの開閉
excel_auto.open_ribbon_dialog("セルの書式設定", "ホーム")
excel_auto.close_ribbon_dialog()

# リボンギャラリーの使用
excel_auto.use_ribbon_gallery("スタイルギャラリー", 0, "ホーム")

# リボンコマンドパスでの実行
excel_auto.execute_ribbon_command("ホーム > フォント > 太字")
excel_auto.execute_ribbon_command("挿入 > グラフ")

# 改善された確実なリボン操作
excel_auto.apply_bold_format()           # 太字を適用
excel_auto.apply_center_alignment()      # 中央揃えを適用
excel_auto.apply_currency_format()       # 通貨形式を適用
excel_auto.open_format_cells_dialog()    # セルの書式設定ダイアログを開く
```

## 短縮キー形式について

新しい短縮キー形式では、Excelのリボンアクセスキーを直接使用します。Configファイルに追記する必要がなく、Altキーを押した後に表示されるキーをそのまま使用できます。

### 短縮キーの形式
- `"H>B"` - ホームタブの太字（Alt + H + B）
- `"H>AC"` - ホームタブの中央揃え（Alt + H + AC）
- `"N>CH"` - 挿入タブのグラフ（Alt + N + CH）
- `"H"` - ホームタブのみクリック（Alt + H）

### 動作原理
短縮キー形式は、Excelのリボンアクセスキーを直接使用します：
1. `Alt`キーを押してリボンにアクセス
2. タブのアクセスキー（例：`H`）を送信
3. ボタンのアクセスキー（例：`AC`）を送信

### 利用可能な短縮キー例

#### タブ
- `H` - ホーム
- `N` - 挿入
- `P` - ページレイアウト
- `M` - 数式
- `A` - データ
- `R` - 校閲
- `W` - 表示
- `D` - 開発

### 柔軟なダイアログ処理機能の詳細

#### ダイアログタイトルパターンの指定方法
- 単一パターン: `"保存の確認"`
- 複数パターン: `["保存の確認", "Save As", "名前を付けて保存"]`
- 部分一致: `"エラー"` で "エラー" を含むタイトルを検索

#### 利用可能なアクション
- `yes` - 「はい」ボタンをクリック
- `no` - 「いいえ」ボタンをクリック
- `ok` - 「OK」ボタンをクリック
- `cancel` - 「キャンセル」ボタンをクリック
- `enable` - 「編集を有効にする」ボタンをクリック（保護ビュー用）

#### よく使用されるダイアログタイトルパターン例
- 保存関連: `["保存の確認", "Save As", "名前を付けて保存"]`
- 上書き確認: `["ファイルの上書き確認", "File Already Exists"]`
- エラー: `["エラー", "Error", "警告"]`
- 保護ビュー: `["保護されたビュー", "Protected View"]`

#### ホームタブのボタン例
- `B` - 太字
- `I` - 斜体
- `U` - 下線
- `AC` - 中央揃え
- `AL` - 左揃え
- `AR` - 右揃え
- `H` - 塗りつぶし
- `FC` - フォント色
- `C` - コピー
- `V` - 貼り付け
- `X` - 切り取り
- `FN` - フォント
- `FS` - フォントサイズ
- `AT` - 上揃え
- `AB` - 下揃え
- `CU` - 通貨
- `PE` - パーセント
- `TH` - 桁区切り

#### 挿入タブのボタン例
- `PT` - ピボットテーブル
- `CH` - グラフ
- `PI` - 画像
- `SH` - 図形
- `TB` - テーブル

#### データタブのボタン例
- `S` - 並べ替え
- `F` - フィルター
- `RD` - 重複の削除
- `TC` - テキストを列に分割

#### 数式タブのボタン例
- `FX` - 関数の挿入
- `AS` - 自動合計
- `RF` - 最近使用した関数

#### 校閲タブのボタン例
- `SP` - スペルチェック
- `TR` - 翻訳
- `CM` - コメント

**注意**: 実際のキーはExcelのバージョンや言語設定によって異なる場合があります。Altキーを押してリボンにアクセスし、表示されるキーを確認してください。

### ウィンドウアクティベーション機能の使用例

```python
# 手動でExcelウィンドウをアクティベート
excel_auto.activate_excel_window(max_retries=5, retry_delay=2.0)

# 操作前にExcelウィンドウがアクティブであることを保証
excel_auto.ensure_excel_active("カスタム操作")

# カスタム機能の例
def custom_function(self):
    """カスタム機能の例"""
    try:
        # 操作前にExcelウィンドウをアクティベート
        self.ensure_excel_active("カスタム機能")
        
        # カスタム処理をここに記述
        send_keys('your_custom_shortcut')
        time.sleep(1)
        
        logger.info("カスタム機能を実行しました")
        return True
        
    except Exception as e:
        logger.error(f"カスタム機能エラー: {e}")
        return False
```

### キーボードショートカットの変更

Excelのバージョンや言語設定によってキーボードショートカットが異なる場合があります。必要に応じて調整してください。
