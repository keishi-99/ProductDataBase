using ProductDatabase.Other;
using System.Data.SQLite;
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

                if (string.IsNullOrEmpty(FilterStringTextBox.Text) || CategoryComboBox.SelectedIndex == -1) {
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
                var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
                ProductInfo.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells[1].Value.ToString() ?? string.Empty;
                ProductInfo.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells[2].Value.ToString() ?? string.Empty;
                ProductInfo.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells[4].Value.ToString() ?? string.Empty;
                ProductInfo.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells[5].Value?.ToString(), out var quantity) ? quantity : 0;
                ProductInfo.RegDate = DataBaseDataGridView.Rows[selectRow].Cells[7].Value.ToString() ?? string.Empty;
                ProductInfo.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells[10].Value.ToString() ?? string.Empty;
                ProductInfo.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells[11].Value.ToString() ?? string.Empty;

                CommonUtils.GenerateReport(ProductInfo);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // リスト印刷
        private void GenerationList() {
            try {
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

                CommonUtils.GenerateList(ProductInfo);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // チェックシート印刷
        private void GenerationCheckSheet() {
            try {
                var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
                ProductInfo.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells[1].Value.ToString() ?? string.Empty;
                ProductInfo.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells[4].Value.ToString() ?? string.Empty;
                ProductInfo.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells[5].Value?.ToString(), out var quantity) ? quantity : 0;
                ProductInfo.RegDate = DataBaseDataGridView.Rows[selectRow].Cells[7].Value.ToString() ?? string.Empty;
                ProductInfo.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells[10].Value.ToString() ?? string.Empty;
                ProductInfo.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells[11].Value.ToString() ?? string.Empty;

                CommonUtils.GenerateCheckSheet(ProductInfo);

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
