using Microsoft.Data.Sqlite;
using ProductDatabase.ExcelService;
using ProductDatabase.Other;
using System.Data;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class HistoryWindow : Form {
        public ProductInformation ProductInfo { get; }

        private readonly int _radioButtonNumber = 0;

        private System.Data.DataTable _historyTable = new();

        private readonly List<string> _listColFilter = [];
        private string _tableName = string.Empty;

        public HistoryWindow(ProductInformation productInfo, int radioButtonNumber) {
            InitializeComponent();
            ProductInfo = productInfo;
            _radioButtonNumber = radioButtonNumber;
            // 最大サイズをディスプレイサイズに合わせる
            if (Screen.PrimaryScreen is not null) {
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

                編集モードToolStripMenuItem.Enabled = Config.IsAdministrator;

                switch (_radioButtonNumber) {
                    case 1:
                        CategoryRadioButton2.Text = "在庫";
                        CategoryRadioButton3.Visible = false;
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = true;
                        GroupModelCheckBox.Visible = false;
                        ShowUsedSubstrateButton.Visible = false;
                        GenerateReportButton.Visible = false;
                        GenerateListButton.Visible = ProductInfo.IsListPrint;
                        GenerateCheckSheetButton.Visible = ProductInfo.IsCheckSheetPrint;
                        if (string.IsNullOrEmpty(ProductInfo.SubstrateModel)) { AllSubstrateCheckBox.Checked = true; }
                        break;
                    case 2:
                        CategoryRadioButton2.Text = "全てのタイプ";
                        CategoryRadioButton3.Text = "シリアル";
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = false;
                        GroupModelCheckBox.Visible = false;
                        break;
                    case 3:
                        CategoryRadioButton1.Visible = false;
                        CategoryRadioButton2.Visible = false;
                        CategoryRadioButton3.Visible = false;
                        StockCheckBox.Visible = false;
                        AllSubstrateCheckBox.Visible = false;
                        GroupModelCheckBox.Visible = false;
                        ShowUsedSubstrateButton.Visible = false;
                        GenerateReportButton.Visible = false;
                        GenerateListButton.Visible = false;
                        GenerateCheckSheetButton.Visible = false;
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadDataAndDisplay(string categoryName, string query, params (string name, object value)[] parameters) {
            using SqliteConnection con = new(GetConnectionRegistration());
            con.Open();

            using SqliteCommand cmd = new(query, con);
            cmd.Parameters.AddRange([.. parameters.Select(p => new SqliteParameter(p.name, p.value))]);
            using (var reader = cmd.ExecuteReader()) {

                _historyTable = new DataTable();
                _historyTable.Load(reader);

                DataBaseDataGridView.Columns.Clear();
                DataBaseDataGridView.DataSource = _historyTable;
            }

            _listColFilter.Clear();
            _listColFilter.Add("");

            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                _listColFilter.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty);
            }
            var headerTextMap = new Dictionary<string, Dictionary<string, string>>
            {
                {
                    "Substrate", new Dictionary<string, string>
                    {
                        { "ID", "ID" },
                        { "StockName", "在庫名" },
                        { "SubstrateName", "基板名" },
                        { "SubstrateModel", "基板型式" },
                        { "SubstrateNumber", "製造番号" },
                        { "OrderNumber", "注文番号" },
                        { "Increase", "追加数" },
                        { "Decrease", "使用数" },
                        { "Defect", "減少数" },
                        { "ProductType", "使用製品名" },
                        { "ProductNumber", "使用製番" },
                        { "OrderNumber1", "使用注番" },
                        { "Person", "担当者" },
                        { "RegDate", "登録日" },
                        { "Comment", "コメント" },
                        { "usedID", "UsedID" }
                    }
                },
                {
                    "SubstrateStock", new Dictionary<string, string>
                    {
                        { "StockName", "在庫名" },
                        { "SubstrateName", "基板名" },
                        { "SubstrateModel", "基板型式" },
                        { "SubstrateNumber", "製造番号" },
                        { "OrderNumber", "注文番号" },
                        { "Stock", "残数" }
                    }
                },
                {
                    "Product", new Dictionary<string, string>
                    {
                        { "ID", "ID" },
                        { "OrderNumber", "注文番号" },
                        { "ProductNumber", "製造番号" },
                        { "ProductName", "製品名" },
                        { "ProductType", "タイプ" },
                        { "ProductModel", "製品型式" },
                        { "Quantity", "数量" },
                        { "Person", "担当者" },
                        { "RegDate", "登録日" },
                        { "Revision", "Rev" },
                        { "RevisionGroup", "Group" },
                        { "SerialFirst", "シリアル先頭" },
                        { "SerialLast", "シリアル末尾" },
                        { "SerialLastNumber", "末番" },
                        { "Comment", "コメント" }
                    }
                },
                {
                    "Serial", new Dictionary<string, string>
                    {
                        { "ID", "ID" },
                        { "Serial", "シリアル" },
                        { "OrderNumber", "注文番号" },
                        { "ProductNumber", "製造番号" },
                        { "ProductName", "製品名" },
                        { "ProductType", "タイプ" },
                        { "ProductModel", "製品型式" },
                        { "RegDate", "登録日" },
                        { "usedID", "UsedID" }
                    }
                },
                {
                    "Reprint", new Dictionary<string, string>
                    {
                        { "ID", "ID" },
                        { "SerialPrintType", "印刷対象" },
                        { "ProductName", "製品名" },
                        { "OrderNumber", "注文番号" },
                        { "ProductNumber", "製造番号" },
                        { "ProductType", "タイプ" },
                        { "ProductModel", "製品型式" },
                        { "Quantity", "数量" },
                        { "Person", "担当者" },
                        { "RegDate", "登録日" },
                        { "Revision", "Rev" },
                        { "SerialFirst", "シリアル先頭" },
                        { "SerialLast", "シリアル末尾" },
                        { "Comment", "コメント" }
                    }
                }
            };

            if (headerTextMap.TryGetValue(categoryName, out var columnHeaders)) {
                foreach (var column in DataBaseDataGridView.Columns.Cast<DataGridViewColumn>()) {
                    if (columnHeaders.TryGetValue(column.Name, out var headerText)) {
                        column.HeaderCell.Value = headerText;
                    }
                }

                CategoryComboBox.Items.Clear();
                CategoryComboBox.Items.Add("");
                for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                    CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty);
                }
            }
            DataBaseDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }

        private void ViewSubstrateRegistrationLog() {
            if (CategoryRadioButton1.Checked) {
                _tableName = "Substrate";
                編集モードToolStripMenuItem.Enabled = Config.IsAdministrator;
                StockCheckBox.Visible = false;
                AllSubstrateCheckBox.Visible = true;
                GroupModelCheckBox.Visible = false;

                var substrateTableName = $"[V_SubstrateWithCategory]";
                var productTableName = $"[T_Product]";
                var substrateCategoryFilter = !string.IsNullOrEmpty(ProductInfo.CategoryName) ? " AND s.CategoryName = @CategoryName" : string.Empty;
                var stockFilter = !string.IsNullOrEmpty(ProductInfo.StockName) ? " AND s.StockName = @StockName" : string.Empty;
                var substrateFilter = !AllSubstrateCheckBox.Checked ? " AND s.SubstrateModel = @SubstrateModel" : string.Empty;

                var query =
                    $"""
                    SELECT
                        s.ID, s.stockName, s.SubstrateName, s.SubstrateModel, s.SubstrateNumber, s.OrderNumber,s.Increase, s.Decrease, s.Defect, p.ProductType, p.ProductNumber, p.OrderNumber, s.Person, s.RegDate, s.Comment, s.UseID
                    FROM
                        {substrateTableName} AS s
                    LEFT JOIN
                        {productTableName} AS p
                    ON
                        s.UseID = p.ID
                    WHERE
                        1=1{substrateCategoryFilter}{stockFilter}{substrateFilter}
                    ORDER BY
                        s.ID DESC
                    ;
                    """;

                LoadDataAndDisplay("Substrate", query, ("@CategoryName", ProductInfo.CategoryName), ("@StockName", ProductInfo.StockName), ("@SubstrateModel", ProductInfo.SubstrateModel));
            }
            if (CategoryRadioButton2.Checked) {
                _tableName = string.Empty;
                編集モードToolStripMenuItem.Enabled = false;
                StockCheckBox.Visible = true;
                AllSubstrateCheckBox.Visible = true;
                GroupModelCheckBox.Visible = true;

                var substrateTableName = $"[V_SubstrateWithCategory]";
                var substrateCategoryFilter = !string.IsNullOrEmpty(ProductInfo.CategoryName) ? " AND CategoryName = @CategoryName" : string.Empty;
                var stockFilter = !string.IsNullOrEmpty(ProductInfo.StockName) ? " AND StockName = @StockName" : string.Empty;
                var substrateFilter = !AllSubstrateCheckBox.Checked ? " AND SubstrateModel = @SubstrateModel" : string.Empty;
                var inStock = StockCheckBox.Checked ? " AND Stock > 0" : string.Empty;

                var selectClause = GroupModelCheckBox.Checked
                    ? "StockName, SubstrateName, SubstrateModel, SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock"
                    : "StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber, SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock";

                var groupByClause = GroupModelCheckBox.Checked
                    ? "SubstrateName, SubstrateModel"
                    : "SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber";

                var orderByClause = GroupModelCheckBox.Checked
                    ? "SubstrateModel"
                    : "MIN(ID) DESC";

                var query =
                    $"""
                    SELECT
                        {selectClause}
                    FROM
                        {substrateTableName}
                    WHERE
                        1=1{substrateCategoryFilter}{stockFilter}{substrateFilter}
                    GROUP BY
                        {groupByClause}
                    HAVING
                        1=1 {inStock}
                    ORDER BY
                        {orderByClause}
                    ;
                    """;

                LoadDataAndDisplay("SubstrateStock", query, ("@CategoryName", ProductInfo.CategoryName), ("@StockName", ProductInfo.StockName), ("@SubstrateModel", ProductInfo.SubstrateModel));
            }
        }

        private void ViewProductRegistration() {
            _tableName = "Product";
            編集モードToolStripMenuItem.Enabled = Config.IsAdministrator;
            GenerateReportButton.Visible = true;
            ShowUsedSubstrateButton.Visible = true;
            if (string.IsNullOrEmpty(ProductInfo.ProductModel)) {
                GenerateListButton.Visible = true;
                GenerateCheckSheetButton.Visible = true;
                CategoryRadioButton1.Visible = false;
                CategoryRadioButton2.Checked = true;
            }
            else {
                GenerateListButton.Visible = ProductInfo.IsListPrint;
                GenerateCheckSheetButton.Visible = ProductInfo.IsCheckSheetPrint;
            }

            var productTableName = $"[V_ProductWithCategory]";
            var productCategoryFilter = !string.IsNullOrEmpty(ProductInfo.CategoryName) ? " AND CategoryName = @CategoryName" : string.Empty;
            var productNameFilter = !string.IsNullOrEmpty(ProductInfo.ProductName) ? " AND ProductName = @ProductName" : string.Empty;
            var productModel = CategoryRadioButton1.Checked ? ProductInfo.ProductModel : string.Empty;
            var productModelFilter = !string.IsNullOrEmpty(productModel) ? " AND ProductModel = @ProductModel" : string.Empty;

            var query =
                $"""
                SELECT *
                FROM {productTableName}
                WHERE 1=1{productCategoryFilter}{productNameFilter}{productModelFilter}
                ORDER BY ID DESC;
                """;

            LoadDataAndDisplay("Product", query, ("@CategoryName", ProductInfo.CategoryName), ("@ProductName", ProductInfo.ProductName), ("@ProductModel", ProductInfo.ProductModel));
        }
        private void ViewSerialLog() {
            _tableName = "Serial";
            編集モードToolStripMenuItem.Enabled = Config.IsAdministrator; ;
            GenerateReportButton.Visible = false;
            ShowUsedSubstrateButton.Visible = false;
            GenerateListButton.Visible = false;
            GenerateCheckSheetButton.Visible = false;

            var serialTableName = $"[T_Serial]";
            var productTableName = $"[V_ProductWithCategory]";
            var substrateCategoryFilter = !string.IsNullOrEmpty(ProductInfo.CategoryName) ? " AND p.CategoryName = @CategoryName" : string.Empty;
            var productNameFilter = !string.IsNullOrEmpty(ProductInfo.ProductName) ? " AND s.ProductName = @ProductName" : string.Empty;

            var query =
                $"""
                SELECT
                    s.rowid,
                    s.Serial,
                    p.OrderNumber,
                    p.ProductNumber,
                    s.ProductName,
                    p.ProductType,
                    p.ProductModel,
                    p.RegDate,
                    s.usedID
                FROM
                    {serialTableName} AS s
                LEFT JOIN
                    {productTableName} AS p
                ON
                    s.UsedID = p.ID
                WHERE
                    1=1{substrateCategoryFilter}{productNameFilter}
                ORDER BY
                    s.rowid DESC
                ;
                """;

            LoadDataAndDisplay("Serial", query, ("@CategoryName", ProductInfo.CategoryName), ("@ProductName", ProductInfo.ProductName));
        }

        private void ViewReprintLog() {
            編集モードToolStripMenuItem.Enabled = false;

            var productModelFilter = !string.IsNullOrEmpty(ProductInfo.ProductModel) ? " AND ProductModel = @ProductModel" : string.Empty;

            var query =
                $"""
                SELECT
                    *
                FROM
                    T_Reprint
                WHERE
                    1=1{productModelFilter}
                ORDER BY
                    ID DESC
                ;
                """;

            LoadDataAndDisplay("Reprint", query, ("@ProductModel", ProductInfo.ProductModel));
        }

        private SqliteConnection? _editModeConnection; // 編集モード用の接続
        private SqliteTransaction? _editModeTransaction; // 編集モード用のトランザクション
        private void EditMode() {
            // 編集モード用の接続とトランザクションを開始
            try {
                _editModeConnection = new SqliteConnection(GetConnectionRegistration());
                _editModeConnection.Open();
                _editModeTransaction = _editModeConnection.BeginTransaction(); // トランザクション開始（ロック）
            } catch (SqliteException ex) {
                MessageBox.Show($"データベースがロックされています。: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            編集モードToolStripMenuItem.Enabled = false;
            編集終了ToolStripMenuItem.Enabled = Config.IsAdministrator;
            DataBaseDataGridView.AllowUserToDeleteRows = true;
            DataBaseDataGridView.ReadOnly = false;
            CategoryRadioButton1.Enabled = false;
            CategoryRadioButton2.Enabled = false;
            CategoryRadioButton3.Enabled = false;
            AllSubstrateCheckBox.Enabled = false;
            StockCheckBox.Enabled = false;
            GenerateCheckSheetButton.Enabled = false;
            GenerateListButton.Enabled = false;
            GenerateReportButton.Enabled = false;
            switch (_tableName) {
                case "Substrate":
                    DataBaseDataGridView.Columns["ID"].ReadOnly = true;
                    DataBaseDataGridView.Columns["StockName"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SubstrateName"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SubstrateModel"].ReadOnly = true;
                    DataBaseDataGridView.Columns["Decrease"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductType"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["OrderNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["UseID"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ID"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["StockName"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SubstrateName"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SubstrateModel"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["Decrease"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductType"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["OrderNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["UseID"].DefaultCellStyle.ForeColor = Color.Red;
                    break;
                case "Product":
                    DataBaseDataGridView.Columns["ID"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductName"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductType"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductModel"].ReadOnly = true;
                    DataBaseDataGridView.Columns["Quantity"].ReadOnly = true;
                    DataBaseDataGridView.Columns["Revision"].ReadOnly = true;
                    DataBaseDataGridView.Columns["RevisionGroup"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SerialFirst"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SerialLast"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SerialLastNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ID"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductName"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductType"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductModel"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["Quantity"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["Revision"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["RevisionGroup"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SerialFirst"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SerialLast"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SerialLastNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    break;
                case "Serial":
                    DataBaseDataGridView.Columns["rowid"].ReadOnly = true;
                    DataBaseDataGridView.Columns["Serial"].ReadOnly = true;
                    DataBaseDataGridView.Columns["OrderNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductName"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductType"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductModel"].ReadOnly = true;
                    DataBaseDataGridView.Columns["RegDate"].ReadOnly = true;
                    DataBaseDataGridView.Columns["UsedID"].ReadOnly = true;
                    DataBaseDataGridView.Columns["rowid"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["Serial"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["OrderNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductName"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductType"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductModel"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["RegDate"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["UsedID"].DefaultCellStyle.ForeColor = Color.Red;
                    break;
            }
        }
        private void SaveRegistrationLog() {
            try {
                // 編集モードで使用している接続を使用
                if (_editModeConnection is null) {
                    MessageBox.Show("編集モードが開始されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var command = _editModeConnection.CreateCommand(); // 編集モードの接続を使用

                switch (_tableName) {
                    case "Substrate":
                        foreach (var row in _historyTable.GetChanges()?.Rows.OfType<DataRow>() ?? []) {
                            if (row.RowState == DataRowState.Modified) {
                                // UPDATE文の設定
                                var substrateTableName = $"[T_Substrate]";
                                command.CommandText =
                                    $"""
                                    UPDATE
                                        {substrateTableName}
                                    SET
                                        SubstrateNumber = @SubstrateNumber,
                                        OrderNumber = @OrderNumber,
                                        Increase = @Increase,
                                        Decrease = @Decrease,
                                        Defect = @Defect,
                                        Person = @Person,
                                        RegDate = @RegDate,
                                        Comment = @Comment,
                                        UseId = @UseId
                                    WHERE
                                        ID = @ID
                                    ;
                                    """;

                                command.Parameters.Clear(); // パラメータをクリア
                                command.Parameters.Add("@SubstrateNumber", SqliteType.Text).Value = row["SubstrateNumber"];
                                command.Parameters.Add("@OrderNumber", SqliteType.Text).Value = row["OrderNumber"];
                                command.Parameters.Add("@Increase", SqliteType.Integer).Value = row["Increase"];
                                command.Parameters.Add("@Decrease", SqliteType.Integer).Value = row["Decrease"];
                                command.Parameters.Add("@Defect", SqliteType.Integer).Value = row["Defect"];
                                command.Parameters.Add("@RegDate", SqliteType.Text).Value = row["RegDate"];
                                command.Parameters.Add("@Person", SqliteType.Text).Value = row["Person"];
                                command.Parameters.Add("@Comment", SqliteType.Text).Value = row["Comment"];
                                command.Parameters.Add("@UseId", SqliteType.Integer).Value = row["UseId"];
                                command.Parameters.Add("@ID", SqliteType.Integer).Value = row["ID"];

                                command.ExecuteNonQuery();
                                // ログ出力
                                string[] logMessageArray = [
                                    $"[基板履歴編集:前]",
                                    $"[{ProductInfo.CategoryName}]",
                                    $"ID[{row["ID", DataRowVersion.Original]}]",
                                    $"注文番号[{row["OrderNumber", DataRowVersion.Original]}]",
                                    $"製造番号[{row["SubstrateNumber", DataRowVersion.Original]}]",
                                    $"ストック名[{row["StockName", DataRowVersion.Original]}]",
                                    $"基板名[{row["SubstrateName", DataRowVersion.Original]}]",
                                    $"型式[{row["SubstrateModel", DataRowVersion.Original]}]",
                                    $"追加数[{row["Increase", DataRowVersion.Original]}]",
                                    $"使用数[{row["Decrease", DataRowVersion.Original]}]",
                                    $"減少数[{row["Defect", DataRowVersion.Original]}]",
                                    $"[]",
                                    $"登録日[{row["RegDate", DataRowVersion.Original]}]",
                                    $"担当者[{row["Person", DataRowVersion.Original]}]",
                                    $"コメント[{row["Comment", DataRowVersion.Original]}]"
                                ];
                                CommonUtils.Logger.AppendLog(logMessageArray);

                                logMessageArray = [
                                    $"[基板履歴編集:後]",
                                    $"[{ProductInfo.CategoryName}]",
                                    $"ID[{row["ID"]}]",
                                    $"注文番号[{row["OrderNumber"]}]",
                                    $"製造番号[{row["SubstrateNumber"]}]",
                                    $"ストック名[{row["StockName"]}]",
                                    $"基板名[{row["SubstrateName"]}]",
                                    $"型式[{row["SubstrateModel"]}]",
                                    $"追加数[{row["Increase"]}]",
                                    $"使用数[{row["Decrease"]}]",
                                    $"減少数[{row["Defect"]}]",
                                    $"[]",
                                    $"登録日[{row["RegDate"]}]",
                                    $"担当者[{row["Person"]}]",
                                    $"コメント[{row["Comment"]}]"
                                ];
                                CommonUtils.Logger.AppendLog(logMessageArray);
                            }
                            else if (row.RowState == DataRowState.Deleted) // 削除行の処理
                            {
                                // DELETE文の設定

                                var substrateTableName = $"[T_Substrate]";
                                command.CommandText =
                                    $"""
                                    UPDATE
                                        {substrateTableName}
                                    SET
                                        StockName = NULL, SubstrateName = NULL, SubstrateModel = NULL, SubstrateNumber = NULL, OrderNumber = NULL, Increase = NULL, Decrease = NULL, Defect = NULL,
                                        Person = NULL, RegDate = NULL, Comment = NULL, UseID = NULL
                                    WHERE
                                        ID = @ID
                                    ;
                                    """;
                                command.Parameters.Clear(); // パラメータをクリア
                                command.Parameters.Add("@ID", SqliteType.Integer).Value = row["ID", DataRowVersion.Original];

                                command.ExecuteNonQuery();
                                // ログ出力
                                string[] logMessageArray = [
                                    $"[基板履歴削除]",
                                    $"[{ProductInfo.CategoryName}]",
                                    $"ID[{row["ID", DataRowVersion.Original]}]",
                                    $"注文番号[{row["OrderNumber", DataRowVersion.Original]}]",
                                    $"製造番号[{row["SubstrateNumber", DataRowVersion.Original]}]",
                                    $"ストック名[{row["StockName", DataRowVersion.Original]}]",
                                    $"基板名[{row["SubstrateName", DataRowVersion.Original]}]",
                                    $"型式[{row["SubstrateModel", DataRowVersion.Original]}]",
                                    $"追加数[{row["Increase", DataRowVersion.Original]}]",
                                    $"使用数[{row["Decrease", DataRowVersion.Original]}]",
                                    $"減少数[{row["Defect", DataRowVersion.Original]}]",
                                    $"[]",
                                    $"登録日[{row["RegDate", DataRowVersion.Original]}]",
                                    $"担当者[{row["Person", DataRowVersion.Original]}]",
                                    $"UseID[{ProductInfo.CategoryName}_{row["UseID", DataRowVersion.Original]}]",
                                ];
                                CommonUtils.Logger.AppendLog(logMessageArray);
                            }
                        }
                        break;
                    case "Product":
                        foreach (var row in _historyTable.GetChanges()?.Rows.OfType<DataRow>() ?? []) {
                            if (row.RowState == DataRowState.Modified) {
                                // UPDATE文の設定
                                var productTableName = $"[T_Product]";
                                command.CommandText =
                                    $"""
                                    UPDATE
                                        {productTableName}
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
                                        ID = @ID
                                    ;
                                    """;

                                command.Parameters.Clear(); // パラメータをクリア
                                command.Parameters.Add("@ID", SqliteType.Integer).Value = row["ID"];
                                command.Parameters.Add("@OrderNumber", SqliteType.Text).Value = row["OrderNumber"];
                                command.Parameters.Add("@ProductNumber", SqliteType.Text).Value = row["ProductNumber"];
                                command.Parameters.Add("@Person", SqliteType.Text).Value = row["Person"];
                                command.Parameters.Add("@RegDate", SqliteType.Text).Value = row["RegDate"];
                                command.Parameters.Add("@Revision", SqliteType.Text).Value = row["Revision"];
                                command.Parameters.Add("@RevisionGroup", SqliteType.Text).Value = row["RevisionGroup"];
                                command.Parameters.Add("@Comment", SqliteType.Text).Value = row["Comment"];

                                command.ExecuteNonQuery();
                                // ログ出力
                                string[] logMessageArray = [
                                    $"[製品履歴編集:前]",
                                    $"[{ProductInfo.CategoryName}]",
                                    $"ID[{row["ID", DataRowVersion.Original]}]",
                                    $"注文番号[{row["OrderNumber", DataRowVersion.Original]}]",
                                    $"製造番号[{row["ProductNumber", DataRowVersion.Original]}]",
                                    $"製品名[{row["ProductName", DataRowVersion.Original]}]",
                                    $"タイプ[{row["ProductType", DataRowVersion.Original]}]",
                                    $"型式[{row["ProductModel", DataRowVersion.Original]}]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"Revision[{row["Revision", DataRowVersion.Original]}]",
                                    $"登録日[{row["RegDate", DataRowVersion.Original]}]",
                                    $"担当者[{row["Person", DataRowVersion.Original]}]",
                                    $"コメント[{row["Comment", DataRowVersion.Original]}]"
                                ];
                                CommonUtils.Logger.AppendLog(logMessageArray);

                                logMessageArray = [
                                    $"[製品履歴編集:後]",
                                    $"[{ProductInfo.CategoryName}]",
                                    $"ID[{row["ID"]}]",
                                    $"注文番号[{row["OrderNumber"]}]",
                                    $"製造番号[{row["ProductNumber"]}]",
                                    $"製品名[{row["ProductName"]}]",
                                    $"タイプ[{row["ProductType"]}]",
                                    $"型式[{row["ProductModel"]}]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"Revision[{row["Revision"]}]",
                                    $"登録日[{row["RegDate"]}]",
                                    $"担当者[{row["Person"]}]",
                                    $"コメント[{row["Comment"]}]"
                                ];
                                CommonUtils.Logger.AppendLog(logMessageArray);
                            }
                            else if (row.RowState == DataRowState.Deleted) // 削除行の処理
                            {
                                // DELETE文の設定
                                var productTableName = $"[T_Product]";
                                command.CommandText =
                                    $"""
                                    UPDATE
                                        {productTableName}
                                    SET
                                        ProductName = NULL, OrderNumber = NULL, ProductNumber = NULL, ProductType = NULL, ProductModel = NULL, Quantity = NULL, Person = NULL, RegDate = NULL,
                                        Revision = NULL, RevisionGroup = NULL, SerialFirst = NULL, SerialLast = NULL, SerialLastNumber = NULL, Comment = NULL
                                    WHERE
                                        ID = @ID
                                    ;
                                    """;
                                command.Parameters.Clear(); // パラメータをクリア
                                command.Parameters.Add("@ID", SqliteType.Integer).Value = row["ID", DataRowVersion.Original];

                                command.ExecuteNonQuery();
                                // ログ出力
                                string[] logMessageArray = [
                                    $"[製品履歴削除]",
                                    $"[{ProductInfo.CategoryName}]",
                                    $"ID[{row["ID", DataRowVersion.Original]}]",
                                    $"注文番号[{row["OrderNumber", DataRowVersion.Original]}]",
                                    $"製造番号[{row["ProductNumber", DataRowVersion.Original]}]",
                                    $"製品名[{row["ProductName", DataRowVersion.Original]}]",
                                    $"タイプ[{row["ProductType", DataRowVersion.Original]}]",
                                    $"型式[{row["ProductModel", DataRowVersion.Original]}]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                ];
                                CommonUtils.Logger.AppendLog(logMessageArray);
                            }
                        }

                        break;
                    case "Serial":
                        foreach (var row in _historyTable.GetChanges()?.Rows.OfType<DataRow>() ?? []) {
                            if (row.RowState == DataRowState.Modified) {
                            }
                            else if (row.RowState == DataRowState.Deleted) // 削除行の処理
                            {
                                // DELETE文の設定
                                var serialTableName = $"[T_Serial]";
                                command.CommandText =
                                    $"""
                                    DELETE FROM
                                        {serialTableName}
                                    WHERE
                                        rowid = @rowid
                                    ;
                                    """;
                                command.Parameters.Clear(); // パラメータをクリア
                                command.Parameters.Add("@rowid", SqliteType.Integer).Value = row["rowid", DataRowVersion.Original];

                                command.ExecuteNonQuery();
                                // ログ出力
                                string[] logMessageArray = [
                                    $"[シリアル履歴削除]",
                                    $"[{ProductInfo.CategoryName}]",
                                    $"ID[{row["rowid", DataRowVersion.Original]}]",
                                    $"製品名[{row["ProductName", DataRowVersion.Original]}]",
                                    $"Serial[{row["Serial", DataRowVersion.Original]}]",
                                    $"UsedID[{row["UsedID", DataRowVersion.Original]}]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                    $"[]",
                                ];
                                CommonUtils.Logger.AppendLog(logMessageArray);
                            }
                        }
                        break;
                }
                // バックアップ作成
                CommonUtils.BackupManager.CreateBackup();

                // 編集モードのトランザクションをコミットしてロック解除
                if (_editModeTransaction is not null) {
                    _editModeTransaction.Commit();
                    _editModeTransaction.Dispose();
                    _editModeTransaction = null;
                }
                if (_editModeConnection is not null) {
                    _editModeConnection.Close();
                    _editModeConnection.Dispose();
                    _editModeConnection = null;
                }

                Close();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // エラー発生時のロールバック（トランザクションが存在する場合）
                if (_editModeTransaction is not null) {
                    _editModeTransaction.Rollback();
                    _editModeTransaction.Dispose();
                    _editModeTransaction = null;
                }
                if (_editModeConnection is not null) {
                    _editModeConnection.Close();
                    _editModeConnection.Dispose();
                    _editModeConnection = null;
                }
            }
        }

        private void CategorySelect(object sender) {
            var selectedRadioButton = (RadioButton)sender;
            if (!selectedRadioButton.Checked) { return; }
            var tag = selectedRadioButton.Tag?.ToString() ?? string.Empty;

            // 動作マッピング用の辞書を定義
            var actionMap = new Dictionary<(int, string), System.Action>
            {
                { (1, "1"), ViewSubstrateRegistrationLog },
                { (1, "2"), ViewSubstrateRegistrationLog },
                { (2, "1"), ViewProductRegistration },
                { (2, "2"), ViewProductRegistration },
                { (2, "3"), ViewSerialLog },
                { (3, "1"), ViewReprintLog }
            };

            // Tag と RadioButtonFlg の組み合わせに応じた動作を実行
            if (actionMap.TryGetValue((_radioButtonNumber, tag), out var action)) {
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

        // 使用基板表示
        private void ShowDataForm() {

            // 選択されたセルがない場合は処理を終了
            if (DataBaseDataGridView.CurrentCell is null) { return; }

            // 新しいフォームクラスを作成
            var dataForm = new Form {
                Text = "使用基板",
                ShowIcon = false, // アイコンを非表示
                Width = 800,
                Height = 400,
                MaximizeBox = false, // 最大化ボタンを非表示
                MinimizeBox = false, // 最小化ボタンを非表示
                StartPosition = FormStartPosition.CenterScreen // 画面中央に表示
            };

            // DataGridViewをフォームに追加
            var dataGridView = new DataGridView {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false, // 行の追加を禁止
                AllowUserToDeleteRows = false, // 行の削除を禁止
                ReadOnly = true, // 編集を禁止
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells // 列幅を自動設定
            };
            dataForm.Controls.Add(dataGridView);

            // データベース接続とSQLクエリの実行
            using SqliteConnection con = new(GetConnectionRegistration());
            {
                con.Open();
                using var cmd = con.CreateCommand();

                var substrateTableName = $"[T_Substrate]";
                cmd.CommandText =
                    $"""
                    SELECT
                        ID, SubstrateName, SubstrateModel, SubstrateNumber, Decrease
                    FROM
                        {substrateTableName}
                    WHERE
                        UseID = @ID
                    ORDER BY
                        SubstrateModel ASC
                    ;
                    """;

                // Sotck名とIDで特定
                //WHERE
                //    StockName = @StockName AND UseID = @ID

                var i = DataBaseDataGridView.SelectedCells[0].RowIndex;
                var id = Convert.ToInt32(DataBaseDataGridView.Rows[i].Cells["ID"].Value);
                //cmd.Parameters.Add("@StockName", SqliteType.Text).Value = ProductInfo.StockName;
                cmd.Parameters.Add("@ID", SqliteType.Integer).Value = id;

                using var reader = cmd.ExecuteReader();
                var dt = new DataTable();

                dt.Load(reader);
                // DataGridViewにデータを表示
                dataGridView.DataSource = dt;
            }
            // 新しいウィンドウを表示
            dataForm.ShowDialog();
        }

        // 成績書作成
        private void GenerateReport() {
            try {
                // --- 処理中カーソルに変更 ---
                Cursor.Current = Cursors.WaitCursor;

                // 選択されたセルがない場合は処理を終了
                if (DataBaseDataGridView.CurrentCell is null) { return; }
                var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
                ProductInfo.ProductID = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells["ID"].Value);
                ProductInfo.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
                ProductInfo.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
                ProductInfo.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
                ProductInfo.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
                ProductInfo.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
                ProductInfo.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
                ProductInfo.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;

                ExcelServiceClosedXml.ReportGeneratorClosedXml.GenerateReport(ProductInfo);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                // --- 通常カーソルに戻す ---
                Cursor.Current = Cursors.Default;
            }
        }
        // リスト作成
        private void GenerateList() {
            try {
                // --- 処理中カーソルに変更 ---
                Cursor.Current = Cursors.WaitCursor;

                // 選択されたセルがない場合は処理を終了
                if (DataBaseDataGridView.CurrentCell is null) { return; }
                var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
                ProductInfo.ProductID = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells["ID"].Value);
                ProductInfo.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
                ProductInfo.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
                ProductInfo.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
                ProductInfo.Quantity = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value);
                ProductInfo.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
                ProductInfo.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
                ProductInfo.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;
                ProductInfo.Comment = DataBaseDataGridView.Rows[selectRow].Cells["Comment"].Value.ToString() ?? string.Empty;

                ExcelServiceClosedXml.ListGeneratorClosedXml.GenerateList(ProductInfo);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                // --- 通常カーソルに戻す ---
                Cursor.Current = Cursors.Default;
            }
        }
        // チェックシート作成
        private void GenerateCheckSheet() {
            try {
                // --- 処理中カーソルに変更 ---
                Cursor.Current = Cursors.WaitCursor;

                // 選択されたセルがない場合は処理を終了
                if (DataBaseDataGridView.CurrentCell is null) { return; }
                var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
                ProductInfo.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
                ProductInfo.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
                ProductInfo.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
                ProductInfo.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
                ProductInfo.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
                ProductInfo.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
                ProductInfo.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;

                ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.GenerateCheckSheet(ProductInfo);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                // --- 通常カーソルに戻す ---
                Cursor.Current = Cursors.Default;
            }
        }

        private void ClosedEvents() {
            // 編集モードのトランザクションをコミットしてロック解除
            if (_editModeTransaction is not null) {
                _editModeTransaction.Commit();
                _editModeTransaction.Dispose();
                _editModeTransaction = null;
            }
            if (_editModeConnection is not null) {
                _editModeConnection.Close();
                _editModeConnection.Dispose();
                _editModeConnection = null;
            }
        }

        private void HistoryWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void 編集ToolStripMenuItem_Click(object sender, EventArgs e) { EditMode(); }
        private void 編集終了ToolStripMenuItem_Click(object sender, EventArgs e) { SaveRegistrationLog(); }
        private void ShowUsedSubstrateButton_Click(object sender, EventArgs e) { ShowDataForm(); }
        private void GenerateReportButton_Click(object sender, EventArgs e) { GenerateReport(); }
        private void GenerateListButton_Click(object sender, EventArgs e) { GenerateList(); }
        private void GenerateCheckSheetButton_Click(object sender, EventArgs e) { GenerateCheckSheet(); }
        private void CategoryComboBox_SelectedIndexChanged(object sender, EventArgs e) { HistoryTableFilter(); }
        private void FilterStringTextBox_TextChanged(object sender, EventArgs e) { HistoryTableFilter(); }
        private void CategoryRadioButton_CheckedChanged(object sender, EventArgs e) { CategorySelect(sender); }
        private void StockCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateRegistrationLog(); }
        private void AllSubstrateCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateRegistrationLog(); }
        private void GroupModelCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateRegistrationLog(); }
        private void HistoryWindow_FormClosed(object sender, FormClosedEventArgs e) { ClosedEvents(); }
    }
}
