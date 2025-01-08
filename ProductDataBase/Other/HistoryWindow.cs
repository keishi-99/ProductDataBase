using ClosedXML.Excel;
using System.Data;
using System.Data.SQLite;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Rendering;
using static ProductDatabase.MainWindow;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProductDatabase {
    public partial class HistoryWindow : Form {
        public ProductInformation ProductInfo { get; }

        private readonly List<string> _listColFilter = [];

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
            DataBaseDataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(DataBaseDataGridView.Font, FontStyle.Bold);
            DataBaseDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            DataBaseDataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataBaseDataGridView.ReadOnly = true;
            DataBaseDataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            DataBaseDataGridView.RowTemplate.DefaultCellStyle.Padding = new Padding(5);
            DataBaseDataGridView.RowTemplate.Height += 10;
        }

        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);

                CategoryRadioButton1.Checked = true;

                switch (ProductInfo.RadioButtonFlg) {
                    case 1:
                        if (ProductInfo.RegType == 0) { CategoryRadioButton2.Visible = false; }
                        CategoryRadioButton2.Text = "在庫";
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = false;
                        GenerationReportButton.Visible = false;
                        GenerationListButton.Visible = false;
                        break;
                    case 2:
                        CategoryRadioButton2.Text = "シリアル";
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = false;
                        break;
                    case 3:
                        CategoryRadioButton1.Visible = false;
                        CategoryRadioButton2.Visible = false;
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = false;
                        GenerationReportButton.Visible = false;
                        GenerationListButton.Visible = false;
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ViewSubstrateRegistrationLog() {
            StockCheckBox.Visible = false;
            AllSubstrateCheckBox.Visible = false;
            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new DataTable();

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
            GenerationListButton.Visible = true;

            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new DataTable();

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
        private void ViewReprintLog() {
            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new DataTable();

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
            var historyTable = new DataTable();

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

            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new DataTable();

            var query = $"""SELECT _rowid_, * FROM "{ProductInfo.ProductName}_Serial" WHERE ProductModel = @ProductModel ORDER BY _rowid_ DESC""";
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
            var actionMap = new Dictionary<(int, string), Action>
            {
                { (1, "1"), ViewSubstrateRegistrationLog },
                { (1, "2"), ViewSubstrateStockLog },
                { (2, "1"), ViewProductRegistrationLog },
                { (2, "2"), ViewSerialLog },
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
        private void HistoryTableFilter(object sender, EventArgs e) {
            try {
                // DataGridView のデータソースから取得
                if (DataBaseDataGridView.DataSource is not DataTable historyTable) {
                    return;
                }

                if (string.IsNullOrEmpty(FilterStringTextBox.Text) || CategoryComboBox.SelectedIndex == 0) {
                    historyTable.DefaultView.RowFilter = null;
                }
                else if (CategoryComboBox.Text == "ID") {
                    if (int.TryParse(FilterStringTextBox.Text, out var id)) {
                        historyTable.DefaultView.RowFilter = $"{_listColFilter[CategoryComboBox.SelectedIndex]} = {id}";
                    }
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
                using FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using XLWorkbook workBookConfig = new(fileStreamConfig);
                var workSheetMain = workBookConfig.Worksheet("Sheet1");

                // セル検索
                var cell = workSheetMain.Search(ProductInfo.ProductModel).FirstOrDefault() ?? throw new Exception($"Configに品目番号:[{ProductInfo.ProductModel}]が見つかりません。");
                var findRow = cell.Address.RowNumber;

                // ワークシートのセルから値を取得
                var directoryPath = workSheetMain.Cell(findRow, 3).GetString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(directoryPath)) { throw new Exception("Configのファイルパスが無効です。"); }
                if (!Directory.Exists(directoryPath)) { throw new FileNotFoundException($"指定されたフォルダが存在しません: {directoryPath}"); }

                var fileName = workSheetMain.Cell(findRow, 4).GetString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(fileName)) { throw new Exception("Configのファイル名が無効です。"); }

                var filePaths = Directory.GetFiles(directoryPath, $"*{fileName}*", SearchOption.TopDirectoryOnly);
                var filePath = filePaths[0];
                var fileExtension = Path.GetExtension(filePath).ToLower(); // 開いたファイルの拡張子取得

                var sheetName = !string.IsNullOrEmpty(workSheetMain.Cell(findRow, 5).GetString())
                    ? workSheetMain.Cell(findRow, 5).GetString()
                    : throw new Exception("シート名がありません。");
                var productNumberRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 6, 2);
                var orderNumberRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 7, 2);
                var quantityRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 8, 2);
                var serialFirstRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 9, 2);
                var serialLastRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 10, 2);
                var saveDirectory = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 11, 2);

                using FileStream fileStreamReport = new(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using XLWorkbook workBookReport = new(fileStreamReport);

                var selectRow = DataBaseDataGridView.CurrentCell.RowIndex;
                var productNumber = ExcelHelper.GetCellValue(DataBaseDataGridView, selectRow, 2);
                var orderNumber = ExcelHelper.GetCellValue(DataBaseDataGridView, selectRow, 1);
                var quantity = ExcelHelper.GetCellValue(DataBaseDataGridView, selectRow, 5);
                var serialFirst = ExcelHelper.GetCellValue(DataBaseDataGridView, selectRow, 10);
                var serialLast = ExcelHelper.GetCellValue(DataBaseDataGridView, selectRow, 11);

                // すべてのワークシートを改ページプレビューに設定
                foreach (var worksheet in workBookReport.Worksheets) {
                    worksheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
                }

                // セルに値を挿入
                var workSheetTemp = workBookReport.Worksheet(sheetName);
                if (!string.IsNullOrEmpty(productNumber)) {
                    var productNumbers = productNumber.Split("-");
                    workSheetTemp.Cell(productNumberRange).Value = productNumbers[0];
                }
                if (!string.IsNullOrEmpty(orderNumber)) {
                    workSheetTemp.Cell(orderNumberRange).Value = orderNumber;
                }
                if (!string.IsNullOrEmpty(quantity)) {
                    workSheetTemp.Cell(quantityRange).Value = quantity;
                }
                if (!string.IsNullOrEmpty(serialFirst)) {
                    workSheetTemp.Cell(serialFirstRange).Value = serialFirst;
                }
                if (!string.IsNullOrEmpty(serialLast)) {
                    workSheetTemp.Cell(serialLastRange).Value = serialLast;
                }

                // ダイアログで保存先を選択
                using SaveFileDialog saveFileDialog = new() {
                    Filter = $"Excel Files (*{fileExtension})|*{fileExtension}|All Files (*.*)|*.*",
                    FileName = $"{productNumber}{fileExtension}",
                    Title = "保存先を選択してください",
                    InitialDirectory = saveDirectory
                };
                if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                    var outputPath = saveFileDialog.FileName;
                    workBookReport.SaveAs(outputPath);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public static class ExcelHelper {
            public static string GetCellValueOrDefault(IXLWorksheet sheet, int rowIndex, int colIndex, int? defaultRow = null) {
                // 対象セルの値を取得
                var value = sheet.Cell(rowIndex, colIndex).GetString();

                // 値が存在しない場合の処理
                if (string.IsNullOrEmpty(value)) {
                    // デフォルト行が指定されていない場合は空文字を返す
                    if (defaultRow is null or 0) {
                        return string.Empty;
                    }

                    // デフォルト行の値を取得
                    return sheet.Cell(defaultRow.Value, colIndex).GetString();
                }

                // 値が存在すればその値を返す
                return value;
            }

            public static string GetCellValue(DataGridView gridView, int rowIndex, int colIndex) {
                return gridView.Rows[rowIndex].Cells[colIndex].Value?.ToString() ?? string.Empty;
            }
        }
        // リスト印刷
        private void GenerationList() {
            try {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "ConfigList.xlsx");
                using FileStream fileStream = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using XLWorkbook workBook = new(fileStream);
                var workSheetMain = workBook.Worksheet("Sheet1");

                // セル検索
                var productModelCell = workSheetMain.Search(ProductInfo.ProductModel).FirstOrDefault() ?? throw new Exception($"Configに品目番号:[{ProductInfo.ProductModel}]が見つかりません。");
                var findRow = productModelCell.Address.RowNumber;

                // ワークシートのセルから値を取得
                var sheetName = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 2, null);
                var productName = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 3, null);
                var productNameRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 4, null);
                var productNumberRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 5, null);
                var orderNumberRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 6, null);
                var regDateRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 7, null);
                var productModel = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 8, null);
                var productModelRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 9, null);
                var quantityRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 10, null);
                var serialFirstRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 11, null);
                var serialLastRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 12, null);
                var commentRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 13, null);
                var qrCodeRange = ExcelHelper.GetCellValueOrDefault(workSheetMain, findRow, 14, null);

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

                var workSheetTemp = workBook.Worksheet(sheetName);
                workSheetTemp.Cell(productNameRange).Value = productName;
                workSheetTemp.Cell(productNumberRange).Value = ProductInfo.ProductNumber;
                workSheetTemp.Cell(orderNumberRange).Value = ProductInfo.OrderNumber;
                workSheetTemp.Cell(regDateRange).Value = ProductInfo.RegDate;
                workSheetTemp.Cell(productModelRange).Value = ProductInfo.ProductModel;
                workSheetTemp.Cell(quantityRange).Value = ProductInfo.Quantity;
                workSheetTemp.Cell(serialFirstRange).Value = ProductInfo.SerialFirst;
                workSheetTemp.Cell(serialLastRange).Value = ProductInfo.SerialLast;
                workSheetTemp.Cell(commentRange).Value = ProductInfo.Comment;

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

                var findColumn = 0;
                for (var i = 0; i <= array1.Count - 1; i++) {

                    var searchRange = workSheetMain.Range(findRow, 1, findRow, 44);
                    var searchValue = $"{array1[i]}";
                    var foundCell = searchRange.CellsUsed(c => c.Value.ToString() == searchValue).FirstOrDefault();

                    if (foundCell != null) {
                        // セルが見つかった場合の処理
                        var foundRow = foundCell.Address.RowNumber;
                        var foundColumn = foundCell.Address.ColumnNumber;
                    }

                    foreach (var cell in workSheetMain.Search(array1[i])) {
                        if (cell.Address.RowNumber == findRow) {
                            findColumn = cell.Address.ColumnNumber;
                            break;
                        }
                    }

                    if (findColumn == 0) {
                        throw new Exception($"{array1[i]}が見つかりません。");
                    }

                    var mainCellValue = workSheetMain.Cell(findRow, findColumn + 1).Value.ToString();
                    var tempCellValue = workSheetTemp.Cell(mainCellValue).Value.ToString();

                    if (mainCellValue != string.Empty) {
                        if (tempCellValue == string.Empty) {
                            workSheetTemp.Cell(mainCellValue).Value = $"{array2[i]}";
                        }
                        else {
                            workSheetTemp.Cell(mainCellValue).Value += $"    {array2[i]}";
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
                    var bmpData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                    System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bmpData.Scan0, pixelData.Pixels.Length);
                    bitmap.UnlockBits(bmpData);
                    using MemoryStream stream = new();
                    bitmap.Save(stream, ImageFormat.Bmp);

                    var image = workSheetTemp.AddPicture(stream);
                    image.MoveTo(workSheetTemp.Cell(qrCodeRange));
                }

                //引数に保存先パスを指定
                var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "temporarily.xlsx");
                workBook.SaveAs(temporarilyPath);

                // 印刷
                Excel.Application xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                var xlBooks = xlApp.Workbooks;
                var xlBook = xlBooks.Open(temporarilyPath);

                // ワークシート選択
                var xlSheets = xlBook.Sheets;
                Excel.Worksheet xlSheet = xlSheets[sheetName];

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

        private void HistoryWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void GenerationReportButton_Click(object sender, EventArgs e) { GenerationReport(); }
        private void GenerationListButton_Click(object sender, EventArgs e) { GenerationList(); }
        private void FilterStringTextBox_TextChanged(object sender, EventArgs e) { HistoryTableFilter(sender, e); }
        private void CategoryRadioButton_CheckedChanged(object sender, EventArgs e) { CategorySelect(sender); }
        private void StockCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateStockLog(); }
        private void AllSubstrateCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateStockLog(); }
    }
}
