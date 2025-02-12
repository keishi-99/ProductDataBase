using OfficeOpenXml;
using ProductDatabase.Other;
using System.Data;
using System.Data.SQLite;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Rendering;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class HistoryWindow : Form {
        public ProductInformation ProductInfo { get; }

        private readonly List<string> _listColFilter = [];
        private bool IsListPrint => ProductInfo.PrintType is 5 or 6;
        private bool IsCheckSheetPrint => ProductInfo.PrintType is 6 or 7;

        public HistoryWindow(ProductInformation productInfo) {
            InitializeComponent();
            ProductInfo = productInfo;
            // 最大サイズをディスプレイサイズに合わせる
            if (Screen.PrimaryScreen != null) {
                var h = Screen.PrimaryScreen.Bounds.Height;
                var w = Screen.PrimaryScreen.Bounds.Width;
                DataBaseDataGridView.MaximumSize = new Size(w, h);
            }
            DataBaseDataGridView.AllowUserToAddRows = false;
            DataBaseDataGridView.AllowUserToDeleteRows = false;
            DataBaseDataGridView.AllowUserToResizeRows = false;
            DataBaseDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
            DataBaseDataGridView.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(DataBaseDataGridView.Font, FontStyle.Bold);
            DataBaseDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            DataBaseDataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataBaseDataGridView.ReadOnly = true;
            DataBaseDataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            DataBaseDataGridView.RowTemplate.DefaultCellStyle.Padding = new Padding(5);
            DataBaseDataGridView.RowTemplate.Height += 10;
        }

        private void LoadEvents() {
            try {
                Font = new System.Drawing.Font(ProductInfo.FontName, ProductInfo.FontSize);

                CategoryRadioButton1.Checked = true;
                CategoryComboBox.SelectedIndex = 0;

                switch (ProductInfo.RadioButtonFlg) {
                    case 1:
                        if (ProductInfo.RegType == 0) { CategoryRadioButton2.Visible = false; }
                        CategoryRadioButton2.Text = "在庫";
                        CategoryRadioButton3.Visible = false;
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = false;
                        GenerationReportButton.Visible = false;
                        GenerationListButton.Visible = IsListPrint;
                        GenerationCheckSheetButton.Visible = IsCheckSheetPrint;
                        break;
                    case 2:
                        CategoryRadioButton2.Text = "全てのタイプ";
                        CategoryRadioButton3.Text = "シリアル";
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = false;
                        GenerationListButton.Visible = IsListPrint;
                        GenerationCheckSheetButton.Visible = IsCheckSheetPrint;
                        break;
                    case 3:
                        CategoryRadioButton1.Visible = false;
                        CategoryRadioButton2.Visible = false;
                        CategoryRadioButton3.Visible = false;
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = false;
                        GenerationReportButton.Visible = false;
                        GenerationListButton.Visible = IsListPrint;
                        GenerationCheckSheetButton.Visible = IsCheckSheetPrint;
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ViewSubstrateRegistrationLog() {
            StockCheckBox.Visible = false;
            AllSubstrateCheckBox.Visible = false;
            CategoryRadioButton3.Visible = false;
            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new System.Data.DataTable();

            var query = $"""SELECT _rowid_, * FROM "{ProductInfo.ProductName}_Substrate" WHERE SubstrateModel = @SubstrateModel ORDER BY _rowid_ DESC""";
            using SQLiteCommand command = new(query, con);
            command.Parameters.AddWithValue("@SubstrateModel", ProductInfo.SubstrateModel);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = historyTable;

            _listColFilter.Clear();
            _listColFilter.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                var headerValue = DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty;
                if (headerValue != null) { _listColFilter.Add(headerValue); }
            }

            DataBaseDataGridView.Columns[0].HeaderCell.Value = "ID";
            DataBaseDataGridView.Columns[1].HeaderCell.Value = "基板名";
            DataBaseDataGridView.Columns[2].HeaderCell.Value = "基板型式";
            DataBaseDataGridView.Columns[3].HeaderCell.Value = "製造番号";
            DataBaseDataGridView.Columns[4].HeaderCell.Value = "注文番号";
            DataBaseDataGridView.Columns[5].HeaderCell.Value = "追加量";
            DataBaseDataGridView.Columns[6].HeaderCell.Value = "減少量";
            DataBaseDataGridView.Columns[7].HeaderCell.Value = "不良";
            DataBaseDataGridView.Columns[8].HeaderCell.Value = "使用製品名";
            DataBaseDataGridView.Columns[9].HeaderCell.Value = "使用製番";
            DataBaseDataGridView.Columns[10].HeaderCell.Value = "使用注番";
            DataBaseDataGridView.Columns[11].HeaderCell.Value = "Rev";
            DataBaseDataGridView.Columns[12].HeaderCell.Value = "担当者";
            DataBaseDataGridView.Columns[13].HeaderCell.Value = "登録日";
            DataBaseDataGridView.Columns[14].HeaderCell.Value = "コメント";
            DataBaseDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            CategoryComboBox.Items.Clear();
            CategoryComboBox.Items.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? string.Empty);
            }
        }
        private void ViewProductRegistrationLog() {

            GenerationReportButton.Visible = true;
            GenerationListButton.Visible = IsListPrint;
            GenerationCheckSheetButton.Visible = IsCheckSheetPrint;

            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new System.Data.DataTable();

            var query = $"""SELECT * FROM "{ProductInfo.ProductName}_Product" WHERE ProductModel = @ProductModel ORDER BY ID DESC""";
            using SQLiteCommand command = new(query, con);
            command.Parameters.AddWithValue("@ProductModel", ProductInfo.ProductModel);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = historyTable;

            _listColFilter.Clear();
            _listColFilter.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                var headerValue = DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty;
                if (headerValue != null) { _listColFilter.Add(headerValue); }
            }

            DataBaseDataGridView.Columns[0].HeaderCell.Value = "ID";
            DataBaseDataGridView.Columns[1].HeaderCell.Value = "注文番号";
            DataBaseDataGridView.Columns[2].HeaderCell.Value = "製造番号";
            DataBaseDataGridView.Columns[3].HeaderCell.Value = "製品名";
            DataBaseDataGridView.Columns[4].HeaderCell.Value = "製品型式";
            DataBaseDataGridView.Columns[5].HeaderCell.Value = "数量";
            DataBaseDataGridView.Columns[6].HeaderCell.Value = "担当者";
            DataBaseDataGridView.Columns[7].HeaderCell.Value = "登録日";
            DataBaseDataGridView.Columns[8].HeaderCell.Value = "Rev";
            DataBaseDataGridView.Columns[9].HeaderCell.Value = "Group";
            DataBaseDataGridView.Columns[10].HeaderCell.Value = "シリアル先頭";
            DataBaseDataGridView.Columns[11].HeaderCell.Value = "シリアル末尾";
            DataBaseDataGridView.Columns[12].HeaderCell.Value = "末番";
            DataBaseDataGridView.Columns[13].HeaderCell.Value = "コメント";
            DataBaseDataGridView.Columns[14].HeaderCell.Value = "使用基板";
            DataBaseDataGridView.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            DataBaseDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            CategoryComboBox.Items.Clear();
            CategoryComboBox.Items.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? string.Empty);
            }
        }
        private void ViewProductRegistrationAllTypesLog() {

            GenerationReportButton.Visible = true;
            GenerationListButton.Visible = IsListPrint;
            GenerationCheckSheetButton.Visible = IsCheckSheetPrint;

            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new System.Data.DataTable();

            var query = $"""SELECT * FROM "{ProductInfo.ProductName}_Product" ORDER BY ID DESC""";
            using SQLiteCommand command = new(query, con);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = historyTable;

            _listColFilter.Clear();
            _listColFilter.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                var headerValue = DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty;
                if (headerValue != null) { _listColFilter.Add(headerValue); }
            }

            DataBaseDataGridView.Columns[0].HeaderCell.Value = "ID";
            DataBaseDataGridView.Columns[1].HeaderCell.Value = "注文番号";
            DataBaseDataGridView.Columns[2].HeaderCell.Value = "製造番号";
            DataBaseDataGridView.Columns[3].HeaderCell.Value = "製品名";
            DataBaseDataGridView.Columns[4].HeaderCell.Value = "製品型式";
            DataBaseDataGridView.Columns[5].HeaderCell.Value = "数量";
            DataBaseDataGridView.Columns[6].HeaderCell.Value = "担当者";
            DataBaseDataGridView.Columns[7].HeaderCell.Value = "登録日";
            DataBaseDataGridView.Columns[8].HeaderCell.Value = "Rev";
            DataBaseDataGridView.Columns[9].HeaderCell.Value = "Group";
            DataBaseDataGridView.Columns[10].HeaderCell.Value = "シリアル先頭";
            DataBaseDataGridView.Columns[11].HeaderCell.Value = "シリアル末尾";
            DataBaseDataGridView.Columns[12].HeaderCell.Value = "末番";
            DataBaseDataGridView.Columns[13].HeaderCell.Value = "コメント";
            DataBaseDataGridView.Columns[14].HeaderCell.Value = "使用基板";
            DataBaseDataGridView.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            DataBaseDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            CategoryComboBox.Items.Clear();
            CategoryComboBox.Items.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? string.Empty);
            }
        }
        private void ViewReprintLog() {
            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new System.Data.DataTable();

            var query = $"""SELECT _rowid_, * FROM Reprint WHERE ProductModel = @ProductModel ORDER BY _rowid_ DESC""";
            using SQLiteCommand command = new(query, con);
            command.Parameters.AddWithValue("@ProductModel", ProductInfo.ProductModel);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = historyTable;

            _listColFilter.Clear();
            _listColFilter.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                var headerValue = DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty;
                if (headerValue != null) { _listColFilter.Add(headerValue); }
            }

            DataBaseDataGridView.Columns[0].HeaderCell.Value = "ID";
            DataBaseDataGridView.Columns[1].HeaderCell.Value = "印刷対象";
            DataBaseDataGridView.Columns[2].HeaderCell.Value = "注文番号";
            DataBaseDataGridView.Columns[3].HeaderCell.Value = "製造番号";
            DataBaseDataGridView.Columns[4].HeaderCell.Value = "製品名";
            DataBaseDataGridView.Columns[5].HeaderCell.Value = "製品型式";
            DataBaseDataGridView.Columns[6].HeaderCell.Value = "数量";
            DataBaseDataGridView.Columns[7].HeaderCell.Value = "担当者";
            DataBaseDataGridView.Columns[8].HeaderCell.Value = "登録日";
            DataBaseDataGridView.Columns[9].HeaderCell.Value = "Rev";
            DataBaseDataGridView.Columns[10].HeaderCell.Value = "シリアル末尾";
            DataBaseDataGridView.Columns[11].HeaderCell.Value = "末番";
            DataBaseDataGridView.Columns[12].HeaderCell.Value = "コメント";
            DataBaseDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            CategoryComboBox.Items.Clear();
            CategoryComboBox.Items.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? string.Empty);
            }
        }
        private void ViewSubstrateStockLog() {
            StockCheckBox.Visible = true;
            AllSubstrateCheckBox.Visible = true;
            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new System.Data.DataTable();

            var otherSubstrate = !AllSubstrateCheckBox.Checked ? " AND SubstrateModel = @SubstrateModel" : string.Empty;
            var inStock = StockCheckBox.Checked ? " AND Stock > 0" : string.Empty;

            var query = $"""
                        SELECT *
                         FROM "{ProductInfo.ProductName}_StockView"
                         WHERE 1=1{otherSubstrate}{inStock}
                        """;

            using SQLiteCommand command = new(query, con);
            command.Parameters.AddWithValue("@SubstrateModel", ProductInfo.SubstrateModel);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = historyTable;

            _listColFilter.Clear();
            _listColFilter.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                var headerValue = DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty;
                if (headerValue != null) { _listColFilter.Add(headerValue); }
            }

            DataBaseDataGridView.Columns[0].HeaderCell.Value = "基板名";
            DataBaseDataGridView.Columns[1].HeaderCell.Value = "基板型式";
            DataBaseDataGridView.Columns[2].HeaderCell.Value = "製造番号";
            DataBaseDataGridView.Columns[3].HeaderCell.Value = "注文番号";
            DataBaseDataGridView.Columns[4].HeaderCell.Value = "残数";
            DataBaseDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            CategoryComboBox.Items.Clear();
            CategoryComboBox.Items.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? string.Empty);
            }
        }
        private void ViewSerialLog() {
            GenerationReportButton.Visible = false;
            GenerationListButton.Visible = false;
            GenerationCheckSheetButton.Visible = false;

            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new System.Data.DataTable();

            var query = $"""SELECT _rowid_, * FROM "{ProductInfo.ProductName}_Serial" ORDER BY _rowid_ DESC""";
            using SQLiteCommand command = new(query, con);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = historyTable;

            _listColFilter.Clear();
            _listColFilter.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                var headerValue = DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty;
                if (headerValue != null) { _listColFilter.Add(headerValue); }
            }

            DataBaseDataGridView.Columns[0].HeaderCell.Value = "ID";
            DataBaseDataGridView.Columns[1].HeaderCell.Value = "シリアル";
            DataBaseDataGridView.Columns[2].HeaderCell.Value = "注文番号";
            DataBaseDataGridView.Columns[3].HeaderCell.Value = "製造番号";
            DataBaseDataGridView.Columns[4].HeaderCell.Value = "製品名";
            DataBaseDataGridView.Columns[5].HeaderCell.Value = "製品型式";
            DataBaseDataGridView.Columns[6].HeaderCell.Value = "登録日";
            DataBaseDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            CategoryComboBox.Items.Clear();
            CategoryComboBox.Items.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? string.Empty);
            }
        }

        private void CategorySelect(object sender) {
            var selectedRadioButton = (RadioButton)sender;
            var tag = selectedRadioButton.Tag?.ToString() ?? string.Empty;

            // 動作マッピング用の辞書を定義
            var actionMap = new Dictionary<(int, string), System.Action>
            {
                { (1, "1"), ViewSubstrateRegistrationLog },
                { (1, "2"), ViewSubstrateStockLog },
                { (2, "1"), ViewProductRegistrationLog },
                { (2, "2"), ViewProductRegistrationAllTypesLog },
                { (2, "3"), ViewSerialLog },
                { (3, "1"), ViewReprintLog }
            };

            // Tag と RadioButtonFlg の組み合わせに応じた動作を実行
            if (actionMap.TryGetValue((ProductInfo.RadioButtonFlg, tag), out var action)) {
                action();
            }
            else {
                MessageBox.Show("無効な選択です。", "[CategorySelect]エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void HistoryTableFilter() {
            try {
                // DataGridView のデータソースから取得
                if (DataBaseDataGridView.DataSource is not System.Data.DataTable historyTable) {
                    return;
                }

                if (string.IsNullOrEmpty(FilterStringTextBox.Text) || CategoryComboBox.SelectedIndex == 0) {
                    historyTable.DefaultView.RowFilter = null;
                }
                else if (historyTable.Columns[_listColFilter[CategoryComboBox.SelectedIndex]]?.DataType == typeof(long)) {
                    historyTable.DefaultView.RowFilter = $"CONVERT({_listColFilter[CategoryComboBox.SelectedIndex]}, 'System.String') LIKE '%{FilterStringTextBox.Text}%'";
                }
                else if (CategoryComboBox.Text != "") {
                    historyTable.DefaultView.RowFilter = $"{_listColFilter[CategoryComboBox.SelectedIndex]} LIKE '*{FilterStringTextBox.Text}*'";
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 成績書作成
        private void GenerationReport() {
            try {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "ConfigReport.xlsx");
                using FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBookConfig = new ExcelPackage(fileStreamConfig);
                //既存ワークシートを取得（workBookはExcelWorkbookクラスオブジェト）
                var sheet = workBookConfig.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Value?.ToString() == ProductInfo.ProductModel) ?? throw new Exception($"Configに品目番号:[{ProductInfo.ProductModel}]が見つかりません。");
                var searchAddressResultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得
                var directoryPath = workSheetMain.Cells[searchAddressResultRow, 3].Value.ToString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(directoryPath)) { throw new Exception("Configのファイルパスが無効です。"); }
                if (!Directory.Exists(directoryPath)) { throw new FileNotFoundException($"指定されたフォルダが存在しません: {directoryPath}"); }

                var searchName = workSheetMain.Cells[searchAddressResultRow, 4].Value.ToString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(searchName)) { throw new Exception("Configのファイル名が無効です。"); }

                var filePaths = Directory.GetFiles(directoryPath, $"*{searchName}*", SearchOption.TopDirectoryOnly);
                var filePath = filePaths[0];
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var fileExtension = Path.GetExtension(filePath).ToLower(); // 開いたファイルの拡張子取得

                var sheetName = !string.IsNullOrEmpty(workSheetMain.Cells[searchAddressResultRow, 5].Value.ToString())
                    ? workSheetMain.Cells[searchAddressResultRow, 5].Value.ToString()
                    : throw new Exception("シート名がありません。");
                var productNumberRange = workSheetMain.Cells[searchAddressResultRow, 6].Value?.ToString() ?? workSheetMain.Cells[2, 6].Value.ToString();
                var orderNumberRange = workSheetMain.Cells[searchAddressResultRow, 7].Value?.ToString() ?? workSheetMain.Cells[2, 7].Value.ToString();
                var quantityRange = workSheetMain.Cells[searchAddressResultRow, 8].Value?.ToString() ?? workSheetMain.Cells[2, 8].Value.ToString();
                var serialFirstRange = workSheetMain.Cells[searchAddressResultRow, 9].Value?.ToString() ?? workSheetMain.Cells[2, 9].Value.ToString();
                var serialLastRange = workSheetMain.Cells[searchAddressResultRow, 10].Value?.ToString() ?? workSheetMain.Cells[2, 10].Value.ToString();
                var productModelRange = workSheetMain.Cells[searchAddressResultRow, 11].Value?.ToString();
                var saveDirectory = workSheetMain.Cells[searchAddressResultRow, 12].Value?.ToString();
                using FileStream fileStreamReport = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBookReport = new ExcelPackage(fileStreamReport);

                var selectRow = DataBaseDataGridView.CurrentCell.RowIndex;
                var productNumber = ExcelHelper.GetCellValue(DataBaseDataGridView, selectRow, 2);
                var orderNumber = ExcelHelper.GetCellValue(DataBaseDataGridView, selectRow, 1);
                var quantity = ExcelHelper.GetCellValue(DataBaseDataGridView, selectRow, 5);
                var serialFirst = ExcelHelper.GetCellValue(DataBaseDataGridView, selectRow, 10);
                var serialLast = ExcelHelper.GetCellValue(DataBaseDataGridView, selectRow, 11);

                // セルに値を挿入
                var workSheetTemp = sheet[sheetName];
                if (!string.IsNullOrEmpty(productNumber)) {
                    var productNumbers = productNumber.Split("-");
                    workSheetTemp.Cells[productNumberRange].Value = productNumbers[0];
                }
                if (!string.IsNullOrEmpty(orderNumberRange)) { workSheetTemp.Cells[orderNumberRange].Value = orderNumber; }
                if (!string.IsNullOrEmpty(quantityRange)) { workSheetTemp.Cells[quantityRange].Value = quantity; }
                if (!string.IsNullOrEmpty(serialFirstRange)) { workSheetTemp.Cells[serialFirstRange].Value = serialFirst; }
                if (!string.IsNullOrEmpty(serialLastRange)) { workSheetTemp.Cells[serialLastRange].Value = serialLast; }
                if (!string.IsNullOrEmpty(productModelRange)) { workSheetTemp.Cells[productModelRange].Value = ProductInfo.ProductModel; }

                // ダイアログで保存先を選択
                using SaveFileDialog saveFileDialog = new() {
                    Filter = $"Excel Files (*{fileExtension})|*{fileExtension}|All Files (*.*)|*.*",
                    FileName = $"{fileName} のコピー{productNumber}{fileExtension}",
                    Title = "保存先を選択してください",
                    InitialDirectory = saveDirectory
                };
                if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                    var outputPath = saveFileDialog.FileName;
                    var fileInfo = new FileInfo(outputPath);
                    workBookReport.SaveAs(fileInfo);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public static class ExcelHelper {
            //public static string GetCellValueOrDefault(ExcelWorksheet sheet, int rowIndex, int colIndex, int? defaultRow = null) {
            //    // 対象セルの値を取得
            //    var value = sheet.Cells[rowIndex, colIndex].Value.ToString();

            //    // 値が存在しない場合の処理
            //    if (string.IsNullOrEmpty(value)) {
            //        // デフォルト行が指定されていない場合は空文字を返す
            //        if (defaultRow is null or 0) {
            //            return string.Empty;
            //        }

            //        // デフォルト行の値を取得
            //        return sheet.Cells[defaultRow.Value, colIndex].Value.ToString() ?? string.Empty;
            //    }

            //    // 値が存在すればその値を返す
            //    return value;
            //}

            public static string GetCellValue(DataGridView gridView, int rowIndex, int colIndex) {
                return gridView.Rows[rowIndex].Cells[colIndex].Value?.ToString() ?? string.Empty;
            }
        }
        // リスト印刷
        private void GenerationList() {
            try {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "ConfigList.xlsx");
                using FileStream fileStream = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBook = new ExcelPackage(fileStream);
                //既存ワークシートを取得（workBookはExcelWorkbookクラスオブジェクト）
                var sheet = workBook.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Value?.ToString() == ProductInfo.ProductModel) ?? throw new Exception($"Configに品目番号:[{ProductInfo.ProductModel}]が見つかりません。");
                var resultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得
                var sheetName = workSheetMain.Cells[resultRow, 2].Value?.ToString();
                var productName = workSheetMain.Cells[resultRow, 3].Value?.ToString();
                var productNameRange = workSheetMain.Cells[resultRow, 4].Value?.ToString();
                var productNumberRange = workSheetMain.Cells[resultRow, 5].Value?.ToString();
                var orderNumberRange = workSheetMain.Cells[resultRow, 6].Value?.ToString();
                var regDateRange = workSheetMain.Cells[resultRow, 7].Value?.ToString();
                var productModel = workSheetMain.Cells[resultRow, 8].Value?.ToString();
                var productModelRange = workSheetMain.Cells[resultRow, 9].Value?.ToString();
                var quantityRange = workSheetMain.Cells[resultRow, 10].Value?.ToString();
                var serialFirstRange = workSheetMain.Cells[resultRow, 11].Value?.ToString();
                var serialLastRange = workSheetMain.Cells[resultRow, 12].Value?.ToString();
                var commentRange = workSheetMain.Cells[resultRow, 13].Value?.ToString();
                var qrCodeRange = workSheetMain.Cells[resultRow, 14].Value?.ToString();

                var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
                ProductInfo.ProductID = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells[0].Value);
                ProductInfo.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells[1].Value.ToString() ?? string.Empty;
                ProductInfo.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells[2].Value.ToString() ?? string.Empty;
                ProductInfo.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells[4].Value.ToString() ?? string.Empty;
                ProductInfo.Quantity = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells[5].Value);
                ProductInfo.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells[10].Value.ToString() ?? string.Empty;
                ProductInfo.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells[11].Value.ToString() ?? string.Empty;
                ProductInfo.Comment = DataBaseDataGridView.Rows[selectRow].Cells[13].Value.ToString() ?? string.Empty;
                ProductInfo.UsedSubstrate = DataBaseDataGridView.Rows[selectRow].Cells[14].Value.ToString() ?? string.Empty;

                var workSheetTemp = sheet[sheetName];
                if (!string.IsNullOrEmpty(productNameRange)) { workSheetTemp.Cells[productNameRange].Value = productName; }
                if (!string.IsNullOrEmpty(productNumberRange)) { workSheetTemp.Cells[productNumberRange].Value = ProductInfo.ProductNumber; }
                if (!string.IsNullOrEmpty(orderNumberRange)) { workSheetTemp.Cells[orderNumberRange].Value = ProductInfo.OrderNumber; }
                if (!string.IsNullOrEmpty(regDateRange)) { workSheetTemp.Cells[regDateRange].Value = DateTime.Now.ToShortDateString(); }
                if (!string.IsNullOrEmpty(productModelRange)) { workSheetTemp.Cells[productModelRange].Value = ProductInfo.ProductModel; }
                if (!string.IsNullOrEmpty(quantityRange)) { workSheetTemp.Cells[quantityRange].Value = ProductInfo.Quantity; }
                if (!string.IsNullOrEmpty(serialFirstRange)) { workSheetTemp.Cells[serialFirstRange].Value = ProductInfo.SerialFirst; }
                if (!string.IsNullOrEmpty(serialLastRange)) { workSheetTemp.Cells[serialLastRange].Value = ProductInfo.SerialLast; }
                if (!string.IsNullOrEmpty(commentRange)) { workSheetTemp.Cells[commentRange].Value = ProductInfo.Comment; }

                //string[] usedSubstrate =  ProductInfo.UsedSubstrate.Split(",");

                var input = ProductInfo.UsedSubstrate;

                // 正規表現で分割
                var matches = Regex.Matches(input, @"\[([^\]]+)\]([^,\r\n]+(?:,[^,\r\n]+)*)");

                List<string> array1 = [];
                List<string> array2 = [];

                foreach (Match match in matches) {
                    var bracketContent = match.Groups[1].Value; // 角括弧の中身
                    var commaSeparatedValues = match.Groups[2].Value; // カンマ区切りの文字列

                    var values = commaSeparatedValues.Split(',');
                    array2.AddRange(values);

                    // valuesの数だけbracketContentを繰り返して追加
                    array1.AddRange(Enumerable.Repeat(bracketContent, values.Length));
                }

                for (var i = 0; i <= array1.Count - 1; i++) {

                    var targetRow = resultRow; // 検索対象の行番号
                    var searchValue = $"{array1[i]}";
                    var foundColumn = 0;

                    var searchAddressResult2 = workSheetMain.Cells
                        .Where(x => x.Start.Row == targetRow) // 指定した行のセルのみを対象にする
                        .First(x => x.Value?.ToString() == searchValue);

                    if (searchAddressResult2 != null) {
                        // セルが見つかった場合の処理
                        foundColumn = searchAddressResult2.Start.Column;
                    }
                    if (foundColumn == 0) {
                        throw new Exception($"{array1[i]}が見つかりません。");
                    }

                    var mainCellValue = workSheetMain.Cells[resultRow, foundColumn + 1].Value.ToString();
                    var tempCellValue = workSheetTemp.Cells[mainCellValue].Value?.ToString();

                    if (mainCellValue != string.Empty) {
                        if (tempCellValue == string.Empty) {
                            workSheetTemp.Cells[mainCellValue].Value = $"{array2[i]}";
                        }
                        else {
                            workSheetTemp.Cells[mainCellValue].Value += $"    {array2[i]}";
                        }
                    }
                }

                // QRコード
                if (!string.IsNullOrEmpty(qrCodeRange)) {
                    BarcodeWriter<PixelData> qr = new() {
                        Format = BarcodeFormat.QR_CODE,
                        Options = new QrCodeEncodingOptions {
                            ErrorCorrection = ErrorCorrectionLevel.L,
                            CharacterSet = "Shift_JIS",
                            Width = 100,
                            Height = 100,
                        },
                        Renderer = new PixelDataRenderer {
                            Foreground = new(Color.Gray.ToArgb()),
                            Background = new(Color.White.ToArgb()),
                        },
                    };

                    var pixelData = qr.Write($"{ProductInfo.OrderNumber};{ProductInfo.ProductNumber};{productModel};{ProductInfo.Quantity};{ProductInfo.SerialFirst};{ProductInfo.SerialLast}");

                    // PixelData を Bitmap に変換
                    using var bitmap = new Bitmap(pixelData.Width, pixelData.Height, PixelFormat.Format32bppArgb);
                    var bmpData = bitmap.LockBits(new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                    System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bmpData.Scan0, pixelData.Pixels.Length);
                    bitmap.UnlockBits(bmpData);
                    using MemoryStream stream = new();
                    bitmap.Save(stream, ImageFormat.Bmp);

                    var image = workSheetTemp.Drawings.AddPicture("QR", bitmap);
                    image.SetPosition(workSheetTemp.Cells[qrCodeRange].Start.Row - 1, 0, workSheetTemp.Cells[qrCodeRange].Start.Column - 1, 0);
                }

                //引数に保存先パスを指定
                var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "temporarilyList.xlsx");
                var fileInfo = new FileInfo(temporarilyPath);
                workBook.SaveAs(fileInfo);

                // 印刷
                Microsoft.Office.Interop.Excel.Application xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                var xlBooks = xlApp.Workbooks;
                var xlBook = xlBooks.Open(temporarilyPath, ReadOnly: true);

                // ワークシート選択
                var xlSheets = xlBook.Sheets;
                Microsoft.Office.Interop.Excel.Worksheet xlSheet = xlSheets[sheetName];

                // ワークシート表示
                xlSheet.Activate();

                // ワークブックを閉じてExcelを終了
                //xlBook.Close(false);
                //xlApp.Quit();

                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // チェックシート印刷
        private void GenerationCheckSheet() {
            try {

                var dialog = new InputDialog1();
                var result = dialog.ShowDialog();

                var temperature = string.Empty;
                var humidity = string.Empty;

                if (result == DialogResult.OK) {
                    temperature = dialog.Temperature;
                    humidity = dialog.Humidity;
                }
                else {
                    return;
                }

                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "ConfigCheckSheet.xlsx");
                using FileStream fileStream = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBook = new ExcelPackage(fileStream);
                //既存ワークシートを取得（workBookはExcelWorkbookクラスオブジェクト）
                var sheet = workBook.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Value?.ToString() == ProductInfo.ProductModel) ?? throw new Exception($"Configに品目番号:[{ProductInfo.ProductModel}]が見つかりません。");
                var resultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得
                var orderNumberRange = workSheetMain.Cells[resultRow, 7].Value?.ToString();
                var quantityRange = workSheetMain.Cells[resultRow, 8].Value?.ToString();
                var serialFirstRange = workSheetMain.Cells[resultRow, 9].Value?.ToString();
                var serialLastRange = workSheetMain.Cells[resultRow, 10].Value?.ToString();
                var regDateYearRange = workSheetMain.Cells[resultRow, 11].Value?.ToString();
                var regDateMonthRange = workSheetMain.Cells[resultRow, 12].Value?.ToString();
                var regDateDayRange = workSheetMain.Cells[resultRow, 13].Value?.ToString();
                var regTemperatureRange = workSheetMain.Cells[resultRow, 14].Value?.ToString();
                var regHumidityRange = workSheetMain.Cells[resultRow, 15].Value?.ToString();

                const int StartColumn = 16;
                var sheetNames = Enumerable.Range(StartColumn, 20) // 無限の範囲
                    .Select(column => workSheetMain.Cells[resultRow, column].Value?.ToString())
                    .TakeWhile(sheetName => !string.IsNullOrWhiteSpace(sheetName)) // 空白でない間
                    .ToList();

                if (sheetNames.Count == 0) { throw new Exception("対象シートがありません。"); }

                var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
                ProductInfo.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells[1].Value.ToString() ?? string.Empty;
                ProductInfo.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells[5].Value?.ToString(), out var quantity) ? quantity : 0;
                ProductInfo.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells[4].Value.ToString() ?? string.Empty;
                ProductInfo.RegDate = DataBaseDataGridView.Rows[selectRow].Cells[7].Value.ToString() ?? string.Empty;
                ProductInfo.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells[10].Value.ToString() ?? string.Empty;
                ProductInfo.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells[11].Value.ToString() ?? string.Empty;

                var date = DateTime.Parse(ProductInfo.RegDate);

                foreach (var sheetName in sheetNames) {
                    var workSheetTemp = sheet[sheetName] ?? throw new Exception($"シート[{sheetName}]が見つかりません。");
                    if (!string.IsNullOrEmpty(orderNumberRange)) { workSheetTemp.Cells[orderNumberRange].Value = ProductInfo.OrderNumber; }
                    if (!string.IsNullOrEmpty(quantityRange)) { workSheetTemp.Cells[quantityRange].Value = ProductInfo.Quantity; }
                    if (!string.IsNullOrEmpty(serialFirstRange)) { workSheetTemp.Cells[serialFirstRange].Value = ProductInfo.SerialFirst; }
                    if (!string.IsNullOrEmpty(serialLastRange)) { workSheetTemp.Cells[serialLastRange].Value = ProductInfo.SerialLast; }
                    if (!string.IsNullOrEmpty(regDateYearRange)) { workSheetTemp.Cells[regDateYearRange].Value = date.Year; }
                    if (!string.IsNullOrEmpty(regDateMonthRange)) { workSheetTemp.Cells[regDateMonthRange].Value = date.Month; }
                    if (!string.IsNullOrEmpty(regDateDayRange)) { workSheetTemp.Cells[regDateDayRange].Value = date.Day; }
                    if (!string.IsNullOrEmpty(regTemperatureRange)) { workSheetTemp.Cells[regTemperatureRange].Value = temperature; }
                    if (!string.IsNullOrEmpty(regHumidityRange)) { workSheetTemp.Cells[regHumidityRange].Value = humidity; }
                }

                // 不要なシートを非表示にする
                var allSheetName = sheet
                    .Select(sh => sh.Name.ToString())
                    .ToList();

                var hiddenSheetNames = allSheetName.Except(sheetNames.Where(name => name != null).Cast<string>()).ToList();
                foreach (var sheetName in hiddenSheetNames) {
                    sheet[sheetName].Hidden = eWorkSheetHidden.Hidden;
                }
                sheet["Sheet1"].Hidden = eWorkSheetHidden.VeryHidden;

                //引数に保存先パスを指定
                var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "temporarilyCheckSheet.xlsx");
                var fileInfo = new FileInfo(temporarilyPath);
                workBook.SaveAs(fileInfo);

                // 印刷
                Microsoft.Office.Interop.Excel.Application xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                var xlBooks = xlApp.Workbooks;
                var xlBook = xlBooks.Open(temporarilyPath, ReadOnly: true);

                //// ワークシート選択
                //var xlSheets = xlBook.Sheets;
                //Microsoft.Office.Interop.Excel.Worksheet xlSheet = xlSheets[0];

                //// ワークシート表示
                //xlSheet.Activate();

                //// ワークブックを閉じてExcelを終了
                //xlBook.Close(false);
                //xlApp.Quit();

                //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void HistoryWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void GenerationReportButton_Click(object sender, EventArgs e) { GenerationReport(); }
        private void GenerationListButton_Click(object sender, EventArgs e) { GenerationList(); }
        private void GenerationCheckSheetButton_Click(object sender, EventArgs e) { GenerationCheckSheet(); }
        private void CategoryComboBox_SelectedIndexChanged(object sender, EventArgs e) { HistoryTableFilter(); }
        private void FilterStringTextBox_TextChanged(object sender, EventArgs e) { HistoryTableFilter(); }
        private void CategoryRadioButton_CheckedChanged(object sender, EventArgs e) { CategorySelect(sender); }
        private void StockCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateStockLog(); }
        private void AllSubstrateCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateStockLog(); }
    }
}
