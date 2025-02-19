using ProductDatabase.Other;
using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class HistoryWindow : Form {
        public ProductInformation ProductInfo { get; }

        private System.Data.DataTable _historyTable = new();

        private readonly List<string> _listColFilter = [];
        private bool IsListPrint => ProductInfo.PrintType is 5 or 6;
        private bool IsCheckSheetPrint => ProductInfo.PrintType is 6 or 7;
        private string _tableName = string.Empty;

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

                編集モードToolStripMenuItem.Enabled = Auth.IsAdministrator;

                switch (ProductInfo.RadioButtonFlg) {
                    case 1:
                        if (ProductInfo.RegType == 0) { CategoryRadioButton2.Visible = false; }
                        CategoryRadioButton2.Text = "在庫";
                        CategoryRadioButton3.Visible = false;
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = true;
                        AllSubstrateStockCheckBox.Visible = false;
                        GenerateReportButton.Visible = false;
                        GenerateListButton.Visible = IsListPrint;
                        GenerateCheckSheetButton.Visible = IsCheckSheetPrint;
                        break;
                    case 2:
                        CategoryRadioButton2.Text = "全てのタイプ";
                        CategoryRadioButton3.Text = "シリアル";
                        if (ProductInfo.RegType == 0) { CategoryRadioButton3.Visible = false; }
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = false;
                        AllSubstrateStockCheckBox.Visible = false;
                        GenerateListButton.Visible = IsListPrint;
                        GenerateCheckSheetButton.Visible = IsCheckSheetPrint;
                        break;
                    case 3:
                        CategoryRadioButton1.Visible = false;
                        CategoryRadioButton2.Visible = false;
                        CategoryRadioButton3.Visible = false;
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = false;
                        AllSubstrateStockCheckBox.Visible = false;
                        GenerateReportButton.Visible = false;
                        GenerateListButton.Visible = IsListPrint;
                        GenerateCheckSheetButton.Visible = IsCheckSheetPrint;
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ViewSubstrateRegistrationLog() {
            _tableName = "Substrate";
            編集モードToolStripMenuItem.Enabled = Auth.IsAdministrator;
            StockCheckBox.Visible = false;
            AllSubstrateCheckBox.Visible = true;
            AllSubstrateStockCheckBox.Visible = false;
            CategoryRadioButton3.Visible = false;
            using SQLiteConnection con = new(GetConnectionRegistration());
            _historyTable = new System.Data.DataTable();

            var otherSubstrate = !AllSubstrateCheckBox.Checked ? " AND SubstrateModel = @SubstrateModel" : string.Empty;

            var query = $"""
                SELECT rowid, * FROM "{ProductInfo.ProductName}_Substrate"
                WHERE 1=1{otherSubstrate} ORDER BY rowid DESC
                """;
            using SQLiteCommand command = new(query, con);
            command.Parameters.AddWithValue("@SubstrateModel", ProductInfo.SubstrateModel);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(_historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = _historyTable;

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
            DataBaseDataGridView.Columns[6].HeaderCell.Value = "使用量";
            DataBaseDataGridView.Columns[7].HeaderCell.Value = "減少量";
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
            _tableName = "Product";
            編集モードToolStripMenuItem.Enabled = Auth.IsAdministrator;
            GenerateReportButton.Visible = true;
            GenerateListButton.Visible = IsListPrint;
            GenerateCheckSheetButton.Visible = IsCheckSheetPrint;

            using SQLiteConnection con = new(GetConnectionRegistration());
            _historyTable = new System.Data.DataTable();

            var query = $"""SELECT * FROM "{ProductInfo.ProductName}_Product" WHERE ProductModel = @ProductModel ORDER BY ID DESC""";
            using SQLiteCommand command = new(query, con);
            command.Parameters.AddWithValue("@ProductModel", ProductInfo.ProductModel);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(_historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = _historyTable;

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
            _tableName = "Product";
            編集モードToolStripMenuItem.Enabled = Auth.IsAdministrator;
            GenerateReportButton.Visible = true;
            GenerateListButton.Visible = IsListPrint;
            GenerateCheckSheetButton.Visible = IsCheckSheetPrint;

            using SQLiteConnection con = new(GetConnectionRegistration());
            _historyTable = new System.Data.DataTable();

            var query = $"""SELECT * FROM "{ProductInfo.ProductName}_Product" ORDER BY ID DESC""";
            using SQLiteCommand command = new(query, con);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(_historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = _historyTable;

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
            編集モードToolStripMenuItem.Enabled = false;
            using SQLiteConnection con = new(GetConnectionRegistration());
            _historyTable = new System.Data.DataTable();

            var query = $"""SELECT rowid, * FROM Reprint WHERE ProductModel = @ProductModel ORDER BY rowid DESC""";
            using SQLiteCommand command = new(query, con);
            command.Parameters.AddWithValue("@ProductModel", ProductInfo.ProductModel);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(_historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = _historyTable;

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
            編集モードToolStripMenuItem.Enabled = false;
            StockCheckBox.Visible = true;
            AllSubstrateCheckBox.Visible = false;
            AllSubstrateStockCheckBox.Visible = true;
            using SQLiteConnection con = new(GetConnectionRegistration());
            _historyTable = new System.Data.DataTable();

            var otherSubstrate = !AllSubstrateStockCheckBox.Checked ? " AND SubstrateModel = @SubstrateModel" : string.Empty;
            //var inStock = StockCheckBox.Checked ? " AND Stock > 0" : string.Empty;
            var inStock = StockCheckBox.Checked ? " AND Stock > 0" : string.Empty;

            //var query = $"""
            //            SELECT *
            //             FROM "{ProductInfo.ProductName}_StockView"
            //             WHERE 1=1{otherSubstrate}{inStock}
            //            """;
            var query = $"""
                        SELECT
                            SubstrateName,
                            SubstrateModel,
                            SubstrateNumber,
                            OrderNumber,
                            SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                        FROM {ProductInfo.ProductName}_Substrate
                        WHERE 1=1{otherSubstrate}
                        GROUP BY SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                        HAVING 1=1{inStock}
                        ORDER BY MIN(rowid);
                        """;

            using SQLiteCommand command = new(query, con);
            command.Parameters.AddWithValue("@SubstrateModel", ProductInfo.SubstrateModel);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(_historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = _historyTable;

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
            _tableName = "Serial";
            編集モードToolStripMenuItem.Enabled = Auth.IsAdministrator; ;
            GenerateReportButton.Visible = false;
            GenerateListButton.Visible = false;
            GenerateCheckSheetButton.Visible = false;

            using SQLiteConnection con = new(GetConnectionRegistration());
            _historyTable = new System.Data.DataTable();

            var query = $"""
                        SELECT
                            s.rowid,
                            s.Serial,
                            p.OrderNumber,
                            p.ProductNumber,
                            p.ProductType,
                            p.ProductModel,
                            p.RegDate,
                            s.usedID
                        FROM
                            "{ProductInfo.ProductName}_Serial" AS s
                        INNER JOIN
                            "{ProductInfo.ProductName}_Product" AS p
                        ON
                            s.UsedID = p.ID;
                        """;

            using SQLiteCommand command = new(query, con);
            // SQLiteDataAdapterのインスタンス化
            using SQLiteDataAdapter adapter = new(command);

            // データの取得とDataTableへの格納
            adapter.Fill(_historyTable);

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = _historyTable;

            _listColFilter.Clear();
            _listColFilter.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                var headerValue = DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty;
                if (headerValue != null) { _listColFilter.Add(headerValue); }
            }

            DataBaseDataGridView.Columns[0].HeaderCell.Value = "RowID";
            DataBaseDataGridView.Columns[1].HeaderCell.Value = "シリアル";
            DataBaseDataGridView.Columns[2].HeaderCell.Value = "注文番号";
            DataBaseDataGridView.Columns[3].HeaderCell.Value = "製造番号";
            DataBaseDataGridView.Columns[4].HeaderCell.Value = "製品名";
            DataBaseDataGridView.Columns[5].HeaderCell.Value = "製品型式";
            DataBaseDataGridView.Columns[6].HeaderCell.Value = "登録日";
            DataBaseDataGridView.Columns[7].HeaderCell.Value = "UsedID";
            DataBaseDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            CategoryComboBox.Items.Clear();
            CategoryComboBox.Items.Add("");
            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value.ToString() ?? string.Empty);
            }
        }

        private void EditMode() {
            編集モードToolStripMenuItem.Enabled = false;
            編集終了ToolStripMenuItem.Enabled = Auth.IsAdministrator;
            DataBaseDataGridView.AllowUserToDeleteRows = true;
            DataBaseDataGridView.ReadOnly = false;
            CategoryRadioButton1.Enabled = false;
            CategoryRadioButton2.Enabled = false;
            CategoryRadioButton3.Enabled = false;
            AllSubstrateCheckBox.Enabled = false;
            StockCheckBox.Enabled = false;
            AllSubstrateStockCheckBox.Enabled = false;
            GenerateCheckSheetButton.Enabled = false;
            GenerateListButton.Enabled = false;
            GenerateReportButton.Enabled = false;
            switch (_tableName) {
                case "Substrate":
                    DataBaseDataGridView.Columns["rowid"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SubstrateName"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SubstrateModel"].ReadOnly = true;
                    DataBaseDataGridView.Columns["UsedProductType"].ReadOnly = true;
                    DataBaseDataGridView.Columns["UsedProductNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["UsedOrderNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["UseID"].ReadOnly = true;
                    DataBaseDataGridView.Columns["rowid"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SubstrateName"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SubstrateModel"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["UsedProductType"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["UsedProductNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["UsedOrderNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["UseID"].DefaultCellStyle.ForeColor = Color.Red;
                    break;
                case "Product":
                    DataBaseDataGridView.Columns["ID"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductType"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductModel"].ReadOnly = true;
                    DataBaseDataGridView.Columns["Quantity"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SerialFirst"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SerialLast"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SerialLastNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["UsedSubstrate"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ID"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductType"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductModel"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["Quantity"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SerialFirst"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SerialLast"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SerialLastNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["UsedSubstrate"].DefaultCellStyle.ForeColor = Color.Red;
                    break;
                case "Serial":
                    DataBaseDataGridView.Columns["rowid"].ReadOnly = true;
                    DataBaseDataGridView.Columns["Serial"].ReadOnly = true;
                    DataBaseDataGridView.Columns["OrderNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductType"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductModel"].ReadOnly = true;
                    DataBaseDataGridView.Columns["RegDate"].ReadOnly = true;
                    DataBaseDataGridView.Columns["UsedID"].ReadOnly = true;
                    DataBaseDataGridView.Columns["rowid"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["Serial"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["OrderNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductType"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductModel"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["RegDate"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["UsedID"].DefaultCellStyle.ForeColor = Color.Red;
                    break;
            }
        }
        private void SaveRegistrationLog() {
            using SQLiteConnection con = new(GetConnectionRegistration());
            con.Open();
            var command = con.CreateCommand();

            switch (_tableName) {
                case "Substrate":
                    foreach (var row in _historyTable.GetChanges()?.Rows.OfType<DataRow>() ?? []) {
                        if (row.RowState == DataRowState.Modified) {
                            //    // UPDATE文の設定
                            command.CommandText = $"""
                            UPDATE "{ProductInfo.ProductName}_Substrate"
                            SET  
                                SubstrateNumber = @SubstrateNumber,
                                OrderNumber = @OrderNumber,
                                Increase = @Increase,
                                Decrease = @Decrease,
                                Defect = @Defect,
                                Revision = @Revision,
                                Person = @Person,
                                RegDate = @RegDate,
                                Comment = @Comment,
                                UseId = @UseId
                            WHERE
                                rowid = @rowid;
                            """;

                            command.Parameters.Clear(); // パラメータをクリア
                            command.Parameters.Add("@SubstrateNumber", DbType.String).Value = row["SubstrateNumber"];
                            command.Parameters.Add("@OrderNumber", DbType.String).Value = row["OrderNumber"];
                            command.Parameters.Add("@Increase", DbType.Int32).Value = row["Increase"];
                            command.Parameters.Add("@Decrease", DbType.Int32).Value = row["Decrease"];
                            command.Parameters.Add("@Defect", DbType.Int32).Value = row["Defect"];
                            command.Parameters.Add("@RegDate", DbType.String).Value = row["RegDate"];
                            command.Parameters.Add("@Person", DbType.String).Value = row["Person"];
                            command.Parameters.Add("@Comment", DbType.String).Value = row["Comment"];
                            command.Parameters.Add("@UseId", DbType.Int32).Value = row["UseId"];
                            command.Parameters.Add("@rowid", DbType.Int32).Value = row["rowid"];

                            command.Connection = con;
                            command.ExecuteNonQuery();
                            // ログ出力
                            CommonUtils.Logger.AppendLog($";[基板履歴編集:前];注文番号[{row["OrderNumber", DataRowVersion.Original]}];製造番号[{row["SubstrateNumber", DataRowVersion.Original]}];製品名[{ProductInfo.ProductName}];基板名[{row["SubstrateName", DataRowVersion.Original]}];型式[{row["SubstrateModel", DataRowVersion.Original]}];追加量[{row["Increase", DataRowVersion.Original]}];使用量[{row["Decrease", DataRowVersion.Original]}];減少量[{row["Defect", DataRowVersion.Original]}];登録日[{row["RegDate", DataRowVersion.Original]}];担当者[{row["Person", DataRowVersion.Original]}];rowID[{row["rowid", DataRowVersion.Original]}];Comment[{row["Comment", DataRowVersion.Original]}]");
                            CommonUtils.Logger.AppendLog($";[基板履歴編集:後];注文番号[{row["OrderNumber"]}];製造番号[{row["SubstrateNumber"]}];製品名[{ProductInfo.ProductName}];基板名[{row["SubstrateName"]}];型式[{row["SubstrateModel"]}];追加量[{row["Increase"]}];使用量[{row["Decrease"]}];減少量[{row["Defect"]}];登録日[{row["RegDate"]}];担当者[{row["Person"]}];rowID[{row["rowid"]}];Comment[{row["Comment"]}]");
                        }
                        else if (row.RowState == DataRowState.Deleted) // 削除行の処理
                        {
                            //DELETE文の設定
                            command.CommandText = $"""
                                DELETE FROM "{ProductInfo.ProductName}_Substrate"
                                WHERE rowid = @rowid;
                                """;
                            command.Parameters.Clear(); // パラメータをクリア
                            command.Parameters.Add("@rowid", DbType.Int32).Value = row["rowid", DataRowVersion.Original];

                            command.Connection = con;
                            command.ExecuteNonQuery();
                            // ログ出力
                            CommonUtils.Logger.AppendLog($";[基板履歴削除];注文番号[{row["OrderNumber", DataRowVersion.Original]}];製造番号[{row["SubstrateNumber", DataRowVersion.Original]}];製品名[{ProductInfo.ProductName}];基板名[{row["SubstrateName", DataRowVersion.Original]}];型式[{row["SubstrateModel", DataRowVersion.Original]}];追加量[{row["Increase", DataRowVersion.Original]}];使用量[{row["Decrease", DataRowVersion.Original]}];減少量[{row["Defect", DataRowVersion.Original]}];Revision[{row["Revision", DataRowVersion.Original]}];登録日[{row["RegDate", DataRowVersion.Original]}];担当者[{row["Person", DataRowVersion.Original]}];rowID[{row["rowid", DataRowVersion.Original]}]");
                        }
                    }
                    break;
                case "Product":
                    foreach (var row in _historyTable.GetChanges()?.Rows.OfType<DataRow>() ?? []) {
                        if (row.RowState == DataRowState.Modified) {
                            //    // UPDATE文の設定
                            command.CommandText = $"""
                            UPDATE "{ProductInfo.ProductName}_Product"
                            SET
                                ID = @ID,
                                OrderNumber = @OrderNumber,
                                ProductNumber = @ProductNumber,
                                Person = @Person,
                                RegDate = @RegDate,
                                Revision = @Revision,
                                RevisionGroup = @RevisionGroup,
                                RegDate = @RegDate,
                                Comment = @Comment
                            WHERE
                                ID = @ID;
                            """;

                            command.Parameters.Clear(); // パラメータをクリア
                            command.Parameters.Add("@ID", DbType.Int32).Value = row["ID"];
                            command.Parameters.Add("@OrderNumber", DbType.String).Value = row["OrderNumber"];
                            command.Parameters.Add("@ProductNumber", DbType.String).Value = row["ProductNumber"];
                            command.Parameters.Add("@Person", DbType.String).Value = row["Person"];
                            command.Parameters.Add("@RegDate", DbType.String).Value = row["RegDate"];
                            command.Parameters.Add("@Revision", DbType.String).Value = row["Revision"];
                            command.Parameters.Add("@RevisionGroup", DbType.String).Value = row["RevisionGroup"];
                            command.Parameters.Add("@Comment", DbType.String).Value = row["Comment"];

                            command.Connection = con;
                            command.ExecuteNonQuery();
                            // ログ出力
                            CommonUtils.Logger.AppendLog($";[製品履歴編集:前];注文番号[{row["OrderNumber", DataRowVersion.Original]}];製造番号[{row["ProductNumber", DataRowVersion.Original]}];製品名[{ProductInfo.ProductName}];タイプ[{row["ProductType", DataRowVersion.Original]}];型式[{row["ProductModel", DataRowVersion.Original]}];Revision[{row["Revision", DataRowVersion.Original]}];登録日[{row["RegDate", DataRowVersion.Original]}];担当者[{row["Person", DataRowVersion.Original]};ID[{row["ID"]}];Comment[{row["Comment", DataRowVersion.Original]}]");
                            CommonUtils.Logger.AppendLog($";[製品履歴編集:後];注文番号[{row["OrderNumber"]}];製造番号[{row["ProductNumber"]}];製品名[{ProductInfo.ProductName}];タイプ[{row["ProductType"]}];型式[{row["ProductModel"]}];Revision[{row["Revision"]}];登録日[{row["RegDate"]}];担当者[{row["Person"]};ID[{row["ID"]}];Comment[{row["Comment"]}]");
                        }
                        else if (row.RowState == DataRowState.Deleted) // 削除行の処理
                        {
                            //DELETE文の設定
                            command.CommandText = $"""
                                DELETE FROM "{ProductInfo.ProductName}_Product"
                                WHERE ID = @ID;
                                """;
                            command.Parameters.Clear(); // パラメータをクリア
                            command.Parameters.Add("@ID", DbType.Int32).Value = row["ID", DataRowVersion.Original];

                            command.Connection = con;
                            command.ExecuteNonQuery();
                            // ログ出力
                            CommonUtils.Logger.AppendLog($";[製品履歴削除];注文番号[{row["OrderNumber", DataRowVersion.Original]}];製造番号[{row["ProductNumber", DataRowVersion.Original]}];製品名[{ProductInfo.ProductName}];タイプ[{row["ProductType", DataRowVersion.Original]}];型式[{row["ProductModel", DataRowVersion.Original]}];ID[{row["ID", DataRowVersion.Original]}]");
                        }
                    }

                    break;
                case "Serial":
                    foreach (var row in _historyTable.GetChanges()?.Rows.OfType<DataRow>() ?? []) {
                        if (row.RowState == DataRowState.Modified) {
                            ////    // UPDATE文の設定
                            //command.CommandText = $"""
                            //UPDATE "{ProductInfo.ProductName}_Serial"
                            //SET
                            //    Serial = @Serial,
                            //    UsedID = @UsedID
                            //WHERE
                            //    rowid = @rowid;
                            //""";

                            //command.Parameters.Clear(); // パラメータをクリア
                            //command.Parameters.Add("@Serial", DbType.String).Value = row["Serial"];
                            //command.Parameters.Add("@UsedID", DbType.String).Value = row["UsedID"];
                            //command.Parameters.Add("@rowid", DbType.Int32).Value = row["rowid"];

                            //command.Connection = con;
                            //command.ExecuteNonQuery();
                            //// ログ出力
                            //CommonUtils.Logger.AppendLog($";[シリアル履歴編集];製品名[{ProductInfo.ProductName}];rowid[{row["rowid"]}];");
                        }
                        else if (row.RowState == DataRowState.Deleted) // 削除行の処理
                        {
                            //DELETE文の設定
                            command.CommandText = $"""
                                DELETE FROM "{ProductInfo.ProductName}_Serial"
                                WHERE rowid = @rowid;
                                """;
                            command.Parameters.Clear(); // パラメータをクリア
                            command.Parameters.Add("@rowid", DbType.Int32).Value = row["rowid", DataRowVersion.Original];

                            command.Connection = con;
                            command.ExecuteNonQuery();
                            // ログ出力
                            CommonUtils.Logger.AppendLog($";[シリアル履歴削除];製品名[{ProductInfo.ProductName}];Serial[{row["Serial", DataRowVersion.Original]}];rowid[{row["rowid", DataRowVersion.Original]}]");
                        }
                    }
                    break;
            }
            // バックアップ作成
            CommonUtils.BackupManager.CreateBackup();
            Close();
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
        private void GenerateReport() {
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
        private void GenerateList() {
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
        private void GenerateCheckSheet() {
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
        private void 編集ToolStripMenuItem_Click(object sender, EventArgs e) { EditMode(); }
        private void 編集終了ToolStripMenuItem_Click(object sender, EventArgs e) { SaveRegistrationLog(); }
        private void GenerateReportButton_Click(object sender, EventArgs e) { GenerateReport(); }
        private void GenerateListButton_Click(object sender, EventArgs e) { GenerateList(); }
        private void GenerateCheckSheetButton_Click(object sender, EventArgs e) { GenerateCheckSheet(); }
        private void CategoryComboBox_SelectedIndexChanged(object sender, EventArgs e) { HistoryTableFilter(); }
        private void FilterStringTextBox_TextChanged(object sender, EventArgs e) { HistoryTableFilter(); }
        private void CategoryRadioButton_CheckedChanged(object sender, EventArgs e) { CategorySelect(sender); }
        private void StockCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateStockLog(); }
        private void AllSubstrateCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateRegistrationLog(); }
        private void AllSubstrateStockCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateStockLog(); }
    }
}
