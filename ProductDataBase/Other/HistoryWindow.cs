using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class HistoryWindow : Form {
        public ProductInfomation ProductInfo { get; }

        private readonly List<string> _listColFilter = [];

        public HistoryWindow(ProductInfomation productInfo) {
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
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void ViewSubstrateRegistrationLog() {
            StockCheckBox.Visible = false;
            AllSubstrateCheckBox.Visible = false;
            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new DataTable();

            var query = $"""SELECT _rowid_, * FROM "Substrate_{ProductInfo.ProductName}" WHERE SubstrateModel = @SubstrateModel ORDER BY _rowid_ DESC""";
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
            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new DataTable();

            var query = $"""SELECT _rowid_, * FROM "Product_{ProductInfo.ProductName}" WHERE ProductModel = @ProductModel ORDER BY _rowid_ DESC""";
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
                        SELECT _rowid_, *
                         FROM "Stock_{ProductInfo.ProductName}"
                         WHERE 1=1{otherSubstrate}{inStock}
                         ORDER BY _rowid_ DESC
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

            DataBaseDataGridView.Columns[0].HeaderCell.Value = "ID";
            DataBaseDataGridView.Columns[1].HeaderCell.Value = "基板名";
            DataBaseDataGridView.Columns[2].HeaderCell.Value = "基板型式";
            DataBaseDataGridView.Columns[3].HeaderCell.Value = "製造番号";
            DataBaseDataGridView.Columns[4].HeaderCell.Value = "注文番号";
            DataBaseDataGridView.Columns[5].HeaderCell.Value = "残数";
            DataBaseDataGridView.Columns[6].HeaderCell.Value = "使用履歴";
            DataBaseDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            CategoryComboBox.Items.Clear();
            CategoryComboBox.Items.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? string.Empty);
            }
        }
        private void ViewSerialLog() {
            using SQLiteConnection con = new(GetConnectionRegistration());
            var historyTable = new DataTable();

            var query = $"""SELECT _rowid_, * FROM "Serial_{ProductInfo.ProductName}" WHERE ProductModel = @ProductModel ORDER BY _rowid_ DESC""";
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
                MessageBox.Show("無効な選択です。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MessageBox.Show(ex.Message);
            }
        }

        private void HistoryWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void FilterStringTextBox_TextChanged(object sender, EventArgs e) { HistoryTableFilter(sender, e); }
        private void CategoryRadioButton_CheckedChanged(object sender, EventArgs e) { CategorySelect(sender); }
        private void StockCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateStockLog(); }
        private void AllSubstrateCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateStockLog(); }
    }
}
