using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.ExcelService;
using ProductDatabase.Other;
using System.Data;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class HistoryWindow : Form {

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;
        private readonly SubstrateMaster _substrateMaster;
        private readonly SubstrateRegisterWork _substrateRegisterWork;
        private readonly AppSettings _appSettings;

        private readonly int _radioButtonNumber = 0;

        private System.Data.DataTable _historyTable = new();

        private readonly List<string> _listColFilter = [];
        private string _tableName = string.Empty;

        private readonly Dictionary<string, Dictionary<string, string>> _headerTextMap = new() {
                {
                    "Substrate", new Dictionary<string, string>
                    {
                        { "ID", "ID" },
                        { "SubstrateID", "基板ID" },
                        { "ProductName", "製品名" },
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
                        { "usedID", "UsedID" },
                        { "CreatedAt", "作成日" },
                    }
                },
                {
                    "SubstrateStock", new Dictionary<string, string>
                    {
                        { "ProductName", "製品名" },
                        { "SubstrateID", "基板ID" },
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
                        { "ProductID", "製品ID" },
                        { "CategoryName", "カテゴリ" },
                        { "OrderNumber", "注文番号" },
                        { "ProductNumber", "製造番号" },
                        { "OLesNumber", "OLes番号" },
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
                        { "Comment", "コメント" },
                        { "CreatedAt", "作成日" },
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
                        { "Comment", "コメント" },
                        { "CreatedAt", "作成日" },
                    }
                }
            };

        public HistoryWindow(ProductMaster productMaster, ProductRegisterWork productRegisterWork, SubstrateMaster substrateMaster, SubstrateRegisterWork substrateRegisterWork, int radioButtonNumber, AppSettings appSettings) {
            InitializeComponent();

            _productMaster = productMaster;
            _productRegisterWork = productRegisterWork;
            _substrateMaster = substrateMaster;
            _substrateRegisterWork = substrateRegisterWork;
            _radioButtonNumber = radioButtonNumber;
            _appSettings = appSettings;

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
            Font = new System.Drawing.Font(_appSettings.FontName, _appSettings.FontSize);

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
                    GenerateListButton.Visible = _productMaster.IsListPrint;
                    GenerateCheckSheetButton.Visible = _productMaster.IsCheckSheetPrint;
                    if (string.IsNullOrEmpty(_substrateMaster.SubstrateModel)) { AllSubstrateCheckBox.Checked = true; }
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
        }

        private void LoadDataAndDisplay(string categoryName, string query, params (string name, object value)[] parameters) {
            using SqliteConnection con = new(GetConnectionRegistration());

            var dynamicParams = new DynamicParameters();
            foreach (var (name, value) in parameters) {
                dynamicParams.Add(name, value);
            }

            using var reader = con.ExecuteReader(query, dynamicParams);

            // DataTableに読み込み
            _historyTable = new DataTable();
            _historyTable.Load(reader);

            // DataGridViewにバインド
            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = _historyTable;

            _listColFilter.Clear();
            _listColFilter.Add("");

            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                _listColFilter.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty);
            }

            if (_headerTextMap.TryGetValue(categoryName, out var columnHeaders)) {
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

                var substrateCategoryFilter = !string.IsNullOrEmpty(_substrateMaster.CategoryName) ? " AND s.CategoryName = @CategoryName" : string.Empty;
                var stockFilter = !string.IsNullOrEmpty(_substrateMaster.ProductName) ? " AND s.ProductName = @ProductName" : string.Empty;
                var substrateIdFilter = !AllSubstrateCheckBox.Checked ? " AND s.SubstrateID = @SubstrateID" : string.Empty;

                var query =
                    $"""
                    SELECT
                        s.ID,
                        s.SubstrateID,
                        s.ProductName,
                        s.SubstrateName,
                        s.SubstrateModel,
                        s.OrderNumber,
                        s.SubstrateNumber,
                        s.Increase,
                        s.Decrease,
                        s.Defect,
                        p.ProductType,
                        p.ProductNumber,
                        p.OrderNumber,
                        s.Person,
                        s.RegDate,
                        s.Comment,
                        s.UseID,
                        s.CreatedAt
                    FROM
                        {Constants.VSubstrateTableName} AS s
                    LEFT JOIN
                        {Constants.VProductTableName} AS p
                    ON
                        s.UseID = p.ID
                    WHERE
                        s.IsDeleted = 0
                        {substrateCategoryFilter}
                        {stockFilter}
                        {substrateIdFilter}
                    ORDER BY
                        s.ID DESC
                    ;
                    """;

                LoadDataAndDisplay("Substrate",
                    query, ("@CategoryName",
                    _substrateMaster.CategoryName),
                    ("@ProductName", _substrateMaster.ProductName),
                    ("@SubstrateID", _substrateMaster.SubstrateID)
                );
            }
            if (CategoryRadioButton2.Checked) {
                _tableName = string.Empty;
                編集モードToolStripMenuItem.Enabled = false;
                在庫調整ToolStripMenuItem.Enabled = true;
                StockCheckBox.Visible = true;
                AllSubstrateCheckBox.Visible = true;
                GroupModelCheckBox.Visible = true;

                var substrateCategoryFilter = !string.IsNullOrEmpty(_substrateMaster.CategoryName) ? " AND CategoryName = @CategoryName" : string.Empty;
                var stockFilter = !string.IsNullOrEmpty(_substrateMaster.ProductName) ? " AND ProductName = @ProductName" : string.Empty;
                var substrateId = !AllSubstrateCheckBox.Checked ? _substrateMaster.SubstrateID : 0;
                var substrateIdFilter = (substrateId != 0) ? " AND SubstrateID = @SubstrateID" : string.Empty;
                var inStock = StockCheckBox.Checked ? " AND Stock > 0" : string.Empty;

                var selectClause = GroupModelCheckBox.Checked
                    ? "SubstrateID, ProductName, SubstrateName, SubstrateModel, SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock"
                    : "SubstrateID, ProductName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber, SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock";

                var groupByClause = GroupModelCheckBox.Checked
                    ? "SubstrateName, SubstrateID"
                    : "SubstrateName, SubstrateID, SubstrateNumber, OrderNumber";

                var orderByClause = GroupModelCheckBox.Checked
                    ? "SubstrateModel"
                    : "MIN(ID) DESC";

                var query =
                    $"""
                    SELECT
                        {selectClause}
                    FROM
                        {Constants.VSubstrateTableName}
                    WHERE
                        IsDeleted = 0
                        {substrateCategoryFilter}
                        {stockFilter}
                        {substrateIdFilter}
                    GROUP BY
                        {groupByClause}
                    HAVING
                        1=1 {inStock}
                    ORDER BY
                        {orderByClause}
                    ;
                    """;

                LoadDataAndDisplay("SubstrateStock",
                    query, ("@CategoryName",
                    _substrateMaster.CategoryName),
                    ("@ProductName",
                    _substrateMaster.ProductName),
                    ("@SubstrateID", _substrateMaster.SubstrateID)
                );
            }
        }

        private void ViewProductRegistration() {
            _tableName = "Product";
            編集モードToolStripMenuItem.Enabled = Config.IsAdministrator;
            GenerateReportButton.Visible = true;
            ShowUsedSubstrateButton.Visible = true;
            if (string.IsNullOrEmpty(_productMaster.ProductModel)) {
                GenerateListButton.Visible = true;
                GenerateCheckSheetButton.Visible = true;
                CategoryRadioButton1.Visible = false;
                CategoryRadioButton2.Checked = true;
            }
            else {
                GenerateListButton.Visible = _productMaster.IsListPrint;
                GenerateCheckSheetButton.Visible = _productMaster.IsCheckSheetPrint;
            }

            var productCategoryFilter = !string.IsNullOrEmpty(_productMaster.CategoryName) ? " AND CategoryName = @CategoryName" : string.Empty;
            var productNameFilter = !string.IsNullOrEmpty(_productMaster.ProductName) ? " AND ProductName = @ProductName" : string.Empty;
            var productId = CategoryRadioButton1.Checked ? _productMaster.ProductID : 0;
            var productIdFilter = (productId != 0) ? " AND ProductID = @ProductID" : string.Empty;

            var query =
                $"""
                SELECT 
                    ID,
                    ProductID,
                    CategoryName,
                    ProductName,
                    ProductType,
                    ProductModel,
                    OrderNumber,
                    ProductNumber,
                    OLesNumber,
                    Quantity,
                    SerialFirst,
                    SerialLast,
                    Revision,
                    RevisionGroup,
                    SerialLastNumber,
                    Person,
                    RegDate,
                    Comment,
                    CreatedAt
                FROM 
                    {Constants.VProductTableName}
                WHERE 
                    IsDeleted = 0
                    {productCategoryFilter}
                    {productNameFilter}
                    {productIdFilter}
                ORDER BY 
                    ID DESC;
                """;

            LoadDataAndDisplay("Product",
                query, ("@CategoryName",
                _productMaster.CategoryName),
                ("@ProductName", _productMaster.ProductName),
                ("@ProductID", _productMaster.ProductID)
            );
        }
        private void ViewSerialLog() {
            _tableName = "Serial";
            編集モードToolStripMenuItem.Enabled = Config.IsAdministrator;
            GenerateReportButton.Visible = false;
            ShowUsedSubstrateButton.Visible = false;
            GenerateListButton.Visible = false;
            GenerateCheckSheetButton.Visible = false;

            var substrateCategoryFilter = !string.IsNullOrEmpty(_productMaster.CategoryName) ? " AND (p.CategoryName = @CategoryName OR p.CategoryName IS NULL)" : string.Empty;
            var productNameFilter = !string.IsNullOrEmpty(_productMaster.ProductName) ? " AND s.ProductName = @ProductName" : string.Empty;
            var productId = CategoryRadioButton1.Checked ? _productMaster.ProductID : 0;
            var productIdFilter = (productId != 0) ? " AND ProductID = @ProductID" : string.Empty;

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
                    {Constants.TSerialTableName} AS s
                LEFT JOIN
                    {Constants.VProductTableName} AS p
                ON
                    s.UsedID = p.ID
                WHERE
                    1=1
                    {substrateCategoryFilter}
                    {productNameFilter}
                    {productIdFilter}
                ORDER BY
                    s.rowid DESC
                ;
                """;

            LoadDataAndDisplay("Serial",
                query,
                ("@CategoryName", _productMaster.CategoryName),
                ("@ProductName", _productMaster.ProductName),
                ("@ProductID", _productMaster.ProductID)
            );
        }

        private void ViewReprintLog() {
            編集モードToolStripMenuItem.Enabled = false;

            var productModelFilter = !string.IsNullOrEmpty(_productMaster.ProductModel) ? " AND ProductModel = @ProductModel" : string.Empty;

            var query =
                $"""
                SELECT
                    ID,
                    SerialPrintType,
                    ProductName,
                    OrderNumber,
                    ProductNumber,
                    ProductType,
                    ProductModel,
                    Quantity,
                    Person,
                    RegDate,
                    Revision,
                    SerialFirst,
                    SerialLast,
                    Comment,
                    CreatedAt
                FROM
                    T_Reprint
                WHERE
                    1=1
                    {productModelFilter}
                ORDER BY
                    ID DESC
                ;
                """;

            LoadDataAndDisplay("Reprint",
                query, ("@ProductModel",
                _productMaster.ProductModel)
            );
        }

        private SqliteConnection? _editModeConnection;
        private SqliteTransaction? _editModeTransaction;
        private void EditMode() {
            try {
                _editModeConnection = new SqliteConnection(GetConnectionRegistration());
                _editModeConnection.Open();
                _editModeTransaction = _editModeConnection.BeginTransaction();
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
                    DataBaseDataGridView.ReadOnly = true;
                    DataBaseDataGridView.DefaultCellStyle.ForeColor = Color.Red;
                    break;
                case "Product":
                    DataBaseDataGridView.Columns["ID"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductID"].ReadOnly = true;
                    DataBaseDataGridView.Columns["CategoryName"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductName"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductType"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ProductModel"].ReadOnly = true;
                    DataBaseDataGridView.Columns["Quantity"].ReadOnly = true;
                    DataBaseDataGridView.Columns["Person"].ReadOnly = true;
                    DataBaseDataGridView.Columns["RegDate"].ReadOnly = true;
                    DataBaseDataGridView.Columns["Revision"].ReadOnly = true;
                    DataBaseDataGridView.Columns["RevisionGroup"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SerialFirst"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SerialLast"].ReadOnly = true;
                    DataBaseDataGridView.Columns["SerialLastNumber"].ReadOnly = true;
                    DataBaseDataGridView.Columns["CreatedAt"].ReadOnly = true;
                    DataBaseDataGridView.Columns["ID"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductID"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["CategoryName"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductName"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductType"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["ProductModel"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["Quantity"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["Person"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["RegDate"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["Revision"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["RevisionGroup"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SerialFirst"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SerialLast"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["SerialLastNumber"].DefaultCellStyle.ForeColor = Color.Red;
                    DataBaseDataGridView.Columns["CreatedAt"].DefaultCellStyle.ForeColor = Color.Red;
                    break;
                case "Serial":
                    DataBaseDataGridView.ReadOnly = true;
                    DataBaseDataGridView.DefaultCellStyle.ForeColor = Color.Red;
                    break;
            }
        }

        private void SaveRegistrationLog() {
            List<string[]> pendingLogs = [];

            try {
                if (_editModeConnection is null || _editModeTransaction is null) {
                    MessageBox.Show("編集モードが開始されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                using var command = _editModeConnection.CreateCommand();
                command.Transaction = _editModeTransaction;

                var changes = _historyTable.GetChanges();
                if (changes == null) {
                    MessageBox.Show("変更がありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    CleanupEditMode();
                    Close();
                    return;
                }

                switch (_tableName) {
                    case "Substrate":
                        ProcessSubstrateChanges(changes, pendingLogs);
                        break;
                    case "Product":
                        ProcessProductChanges(changes, pendingLogs);
                        break;
                    case "Serial":
                        ProcessSerialChanges(changes, pendingLogs);
                        break;
                }

                // バックアップ作成
                CommonUtils.BackupManager.CreateBackup();

                _editModeTransaction.Commit();

                foreach (var log in pendingLogs) {
                    CommonUtils.Logger.AppendLog(log);
                }

                CleanupEditMode();
                Close();

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                try {
                    _editModeTransaction?.Rollback();
                } catch (Exception rollbackEx) {
                    MessageBox.Show($"ロールバックに失敗しました: {rollbackEx.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                CleanupEditMode();
            }
        }

        private void CleanupEditMode() {
            _editModeTransaction?.Dispose();
            _editModeTransaction = null;
            _editModeConnection?.Close();
            _editModeConnection?.Dispose();
            _editModeConnection = null;
        }

        private void ProcessSubstrateChanges(DataTable changes, List<string[]> pendingLogs) {
            foreach (var row in changes.Rows.OfType<DataRow>()) {
                if (row.RowState == DataRowState.Modified) {
                    LogSubstrateEdit(row, pendingLogs);
                    UpdateSubstrateRow(_editModeConnection!, row);
                }
                else if (row.RowState == DataRowState.Deleted) {
                    LogSubstrateDelete(row, pendingLogs);
                    DeleteSubstrateRow(_editModeConnection!, row);
                }
            }
        }

        private void LogSubstrateEdit(DataRow row, List<string[]> pendingLogs) {
            pendingLogs.Add([
                "[基板履歴編集:前]",
                $"[{_substrateMaster.CategoryName}]",
                $"ID[{GetValue(row, "ID", DataRowVersion.Original)}]",
                $"注文番号[{GetValue(row, "OrderNumber", DataRowVersion.Original)}]",
                $"製造番号[{GetValue(row, "SubstrateNumber", DataRowVersion.Original)}]",
                $"[]",
                $"製品名[{GetValue(row,"ProductName", DataRowVersion.Original)}]",
                $"基板名[{GetValue(row,"SubstrateName", DataRowVersion.Original)}]",
                $"型式[{GetValue(row, "SubstrateModel", DataRowVersion.Original)}]",
                $"追加数[{GetValue(row, "Increase", DataRowVersion.Original)}]",
                $"使用数[{GetValue(row, "Decrease", DataRowVersion.Original)}]",
                $"減少数[{GetValue(row, "Defect", DataRowVersion.Original)}]",
                $"[]",
                $"登録日[{GetValue(row, "RegDate", DataRowVersion.Original)}]",
                $"担当者[{GetValue(row, "Person", DataRowVersion.Original)}]",
                $"コメント[{GetValue(row, "Comment", DataRowVersion.Original)}]"
            ]);

            pendingLogs.Add([
                "[基板履歴編集:後]",
                $"[{_substrateMaster.CategoryName}]",
                $"ID[{GetValue(row, "ID")}]",
                $"注文番号[{GetValue(row,"OrderNumber")}]",
                $"製造番号[{GetValue(row,"SubstrateNumber")}]",
                $"[]",
                $"製品名[{GetValue(row, "ProductName")}]",
                $"基板名[{GetValue(row, "SubstrateName")}]",
                $"型式[{GetValue(row, "SubstrateModel")}]",
                $"追加数[{GetValue(row, "Increase")}]",
                $"使用数[{GetValue(row, "Decrease")}]",
                $"減少数[{GetValue(row, "Defect")}]",
                $"[]",
                $"登録日[{GetValue(row, "RegDate")}]",
                $"担当者[{GetValue(row, "Person")}]",
                $"コメント[{GetValue(row, "Comment")}]"
            ]);
        }
        private void UpdateSubstrateRow(IDbConnection connection, DataRow row) {
            var sql =
                $"""
                UPDATE {Constants.TSubstrateTableName}
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
                WHERE ID = @ID;
                """;

            connection.Execute(sql, new {
                SubstrateNumber = row["SubstrateNumber"],
                OrderNumber = row["OrderNumber"],
                Increase = row["Increase"],
                Decrease = row["Decrease"],
                Defect = row["Defect"],
                RegDate = row["RegDate"],
                Person = row["Person"],
                Comment = row["Comment"],
                UseId = row["UseId"],
                ID = row["ID"]
            }, _editModeTransaction);
        }

        private void LogSubstrateDelete(DataRow row, List<string[]> pendingLogs) {
            pendingLogs.Add([
                "[基板履歴削除]",
                $"[{_substrateMaster.CategoryName}]",
                $"ID[{GetValue(row, "ID", DataRowVersion.Original)}]",
                $"注文番号[{GetValue(row, "OrderNumber", DataRowVersion.Original)}]",
                $"製造番号[{GetValue(row, "SubstrateNumber", DataRowVersion.Original)}]",
                $"[]",
                $"製品名[{GetValue(row,"ProductName", DataRowVersion.Original)}]",
                $"基板名[{GetValue(row,"SubstrateName", DataRowVersion.Original)}]",
                $"型式[{GetValue(row, "SubstrateModel", DataRowVersion.Original)}]",
                $"追加数[{GetValue(row, "Increase", DataRowVersion.Original)}]",
                $"使用数[{GetValue(row, "Decrease", DataRowVersion.Original)}]",
                $"減少数[{GetValue(row, "Defect", DataRowVersion.Original)}]",
                $"[]",
                $"登録日[{GetValue(row, "RegDate", DataRowVersion.Original)}]",
                $"担当者[{GetValue(row, "Person", DataRowVersion.Original)}]",
                $"コメント[{GetValue(row, "Comment", DataRowVersion.Original)}]"
            ]);
        }
        private void DeleteSubstrateRow(IDbConnection connection, DataRow row) {
            var sql =
                $"""
                UPDATE {Constants.TSubstrateTableName}
                SET
                    IsDeleted = 1,
                    DeletedAt = datetime('now', 'localtime')
                WHERE ID = @ID;
                """;

            connection.Execute(sql, new {
                ID = row["ID", DataRowVersion.Original]
            }, _editModeTransaction);
        }

        private void ProcessProductChanges(DataTable changes, List<string[]> pendingLogs) {
            foreach (var row in changes.Rows.OfType<DataRow>()) {
                if (row.RowState == DataRowState.Modified) {
                    LogProductEdit(row, pendingLogs);
                    UpdateProductRow(_editModeConnection!, row);
                }
                else if (row.RowState == DataRowState.Deleted) {
                    LogProductDelete(row, pendingLogs);
                    LogProductSubstrateDelete(_editModeConnection!, row, pendingLogs);
                    LogProductSerialDelete(_editModeConnection!, row, pendingLogs);
                    DeleteProductRow(_editModeConnection!, row);
                    DeleteProductSubstrateRow(_editModeConnection!, row);
                    DeleteProductSerialRow(_editModeConnection!, row);
                }
            }
        }
        private void UpdateProductRow(IDbConnection connection, DataRow row) {
            var sql =
                $"""
                UPDATE {Constants.TProductTableName}
                SET
                    OrderNumber = @OrderNumber,
                    ProductNumber = @ProductNumber,
                    OLesNumber = @OLesNumber,
                    Person = @Person,
                    RegDate = @RegDate,
                    Revision = @Revision,
                    RevisionGroup = @RevisionGroup,
                    Comment = @Comment
                WHERE ID = @ID;
                """;

            connection.Execute(sql, new {
                ID = row["ID"],
                OrderNumber = row["OrderNumber"],
                ProductNumber = row["ProductNumber"],
                OLesNumber = row["OLesNumber"],
                Person = row["Person"],
                RegDate = row["RegDate"],
                Revision = row["Revision"],
                RevisionGroup = row["RevisionGroup"],
                Comment = row["Comment"]
            }, _editModeTransaction);
        }

        private void LogProductEdit(DataRow row, List<string[]> pendingLogs) {
            pendingLogs.Add([
                $"[製品履歴編集:前]",
                $"[{_productMaster.CategoryName}]",
                $"ID[{row["ID", DataRowVersion.Original]}]",
                $"注文番号[{row["OrderNumber", DataRowVersion.Original]}]",
                $"製造番号[{row["ProductNumber", DataRowVersion.Original]}]",
                $"OLes番号[{row["OLesNumber", DataRowVersion.Original]}]",
                $"製品名[{row["ProductName", DataRowVersion.Original]}]",
                $"タイプ[{row["ProductType", DataRowVersion.Original]}]",
                $"型式[{row["ProductModel", DataRowVersion.Original]}]",
                $"数量[{row["Quantity", DataRowVersion.Original]}]",
                $"シリアル先頭[{row["SerialFirst", DataRowVersion.Original]}]",
                $"シリアル末尾[{row["SerialLast", DataRowVersion.Original]}]",
                $"Revision[{row["Revision", DataRowVersion.Original]}]",
                $"登録日[{row["RegDate", DataRowVersion.Original]}]",
                $"担当者[{row["Person", DataRowVersion.Original]}]",
                $"コメント[{row["Comment", DataRowVersion.Original]}]"
            ]);

            pendingLogs.Add([
                $"[製品履歴編集:後]",
                $"[{_productMaster.CategoryName}]",
                $"ID[{GetValue(row,"ID")}]",
                $"注文番号[{GetValue(row,"OrderNumber")}]",
                $"製造番号[{GetValue(row,"ProductNumber")}]",
                $"OLes番号[{GetValue(row,"OLesNumber")}]",
                $"製品名[{GetValue(row, "ProductName")}]",
                $"タイプ[{GetValue(row, "ProductType")}]",
                $"型式[{GetValue(row, "ProductModel")}]",
                $"数量[{GetValue(row, "Quantity")}]",
                $"シリアル先頭[{GetValue(row, "SerialFirst")}]",
                $"シリアル末尾[{GetValue(row, "SerialLast")}]",
                $"Revision[{GetValue(row, "Revision")}]",
                $"登録日[{GetValue(row, "RegDate")}]",
                $"担当者[{GetValue(row, "Person")}]",
                $"コメント[{GetValue(row, "Comment")}]"
            ]);
        }
        private void LogProductDelete(DataRow row, List<string[]> pendingLogs) {
            pendingLogs.Add([
                "[製品履歴削除]",
                $"[{_productMaster.CategoryName}]",
                $"ID[{GetValue(row, "ID", DataRowVersion.Original)}]",
                $"注文番号[{GetValue(row, "OrderNumber", DataRowVersion.Original)}]",
                $"製造番号[{GetValue(row, "ProductNumber", DataRowVersion.Original)}]",
                $"OLes番号[{GetValue(row, "OLesNumber", DataRowVersion.Original)}]",
                $"製品名[{GetValue(row,"ProductName", DataRowVersion.Original)}]",
                $"タイプ[{GetValue(row,"ProductType", DataRowVersion.Original)}]",
                $"型式[{GetValue(row, "ProductModel", DataRowVersion.Original)}]",
                $"数量[{GetValue(row, "Quantity", DataRowVersion.Original)}]",
                $"シリアル先頭[{GetValue(row, "SerialFirst", DataRowVersion.Original)}]",
                $"シリアル末尾[{GetValue(row, "SerialLast", DataRowVersion.Original)}]",
                $"Revision[{GetValue(row, "Revision", DataRowVersion.Original)}]",
                $"登録日[{GetValue(row, "RegDate", DataRowVersion.Original)}]",
                $"担当者[{GetValue(row, "Person", DataRowVersion.Original)}]",
                $"コメント[{GetValue(row, "Comment", DataRowVersion.Original)}]"
            ]);
        }
        private void LogProductSubstrateDelete(IDbConnection connection, DataRow row, List<string[]> pendingLogs) {
            var sql =
                $"""
                SELECT
                    ID,
                    OrderNumber,
                    SubstrateNumber,
                    ProductName,
                    SubstrateName,
                    SubstrateModel,
                    Increase,
                    Decrease,
                    Defect,
                    RegDate,
                    Person,
                    Comment,
                    UseID
                FROM {Constants.VSubstrateTableName}
                WHERE UseID = @ID;
                """;

            var results = connection.Query(sql, new {
                ID = row["ID", DataRowVersion.Original]
            }, _editModeTransaction);

            foreach (var item in results) {
                pendingLogs.Add([
                    $"[製品削除に伴う基板削除]",
                    $"[{_substrateMaster.CategoryName}]",
                    $"ID[{item.ID}]",
                    $"注文番号[{item.OrderNumber}]",
                    $"製造番号[{item.SubstrateNumber}]",
                    $"[]",
                    $"製品名[{item.ProductName}]",
                    $"基板名[{item.SubstrateName}]",
                    $"型式[{item.SubstrateModel}]",
                    $"追加数[{item.Increase}]",
                    $"使用数[{item.Decrease}]",
                    $"減少数[{item.Defect}]",
                    $"登録日[{item.RegDate}]",
                    $"担当者[{item.Person}]",
                    $"コメント[{item.Comment}]",
                    $"UseID[{item.UseID}]",
                ]);
            }
        }
        private void LogProductSerialDelete(IDbConnection connection, DataRow row, List<string[]> pendingLogs) {
            var sql =
                $"""
                SELECT
                    rowid,
                    ProductName,
                    Serial,
                    UsedID
                FROM {Constants.TSerialTableName}
                WHERE UsedID = @ID;
                """;

            var results = connection.Query(sql, new {
                ID = row["ID", DataRowVersion.Original]
            }, _editModeTransaction);

            foreach (var item in results) {
                pendingLogs.Add([
                    $"[製品削除に伴うシリアル削除]",
                    $"[{_productMaster.CategoryName}]",
                    $"ID[{item.rowid}]",
                    $"製品名[{item.ProductName}]",
                    $"Serial[{item.Serial}]",
                    $"UsedID[{item.UsedID}]",
                    $"[]", $"[]", $"[]", $"[]",
                    $"[]", $"[]", $"[]", $"[]",
                    $"[]", $"[]",
                ]);
            }
        }

        private void DeleteProductRow(IDbConnection connection, DataRow row) {
            var sql =
                $"""
                UPDATE {Constants.TProductTableName}
                SET
                    IsDeleted = 1,
                    DeletedAt = datetime('now', 'localtime')
                WHERE ID = @ID;
                """;

            connection.Execute(sql, new {
                ID = row["ID", DataRowVersion.Original]
            }, _editModeTransaction);
        }
        private void DeleteProductSubstrateRow(IDbConnection connection, DataRow row) {
            var sql = $"""
        UPDATE {Constants.TSubstrateTableName}
        SET
            IsDeleted = 1,
            DeletedAt = datetime('now', 'localtime')
        WHERE UseID = @ID;
        """;

            connection.Execute(sql, new {
                ID = row["ID", DataRowVersion.Original]
            }, _editModeTransaction);
        }
        private void DeleteProductSerialRow(IDbConnection connection, DataRow row) {
            var sql =
                $"""
                DELETE FROM {Constants.TSerialTableName}
                WHERE UsedID = @ID;
                """;

            connection.Execute(sql, new {
                ID = row["ID", DataRowVersion.Original]
            }, _editModeTransaction);
        }

        private void ProcessSerialChanges(DataTable changes, List<string[]> pendingLogs) {
            foreach (var row in changes.Rows.OfType<DataRow>()) {
                if (row.RowState == DataRowState.Deleted) {
                    LogSerialDelete(row, pendingLogs);
                    DeleteSerialRow(_editModeConnection!, row);
                }
            }
        }
        private void LogSerialDelete(DataRow row, List<string[]> pendingLogs) {
            pendingLogs.Add([
            $"[シリアル履歴削除]",
                $"[{_productMaster.CategoryName}]",
                $"ID[{GetValue(row, "rowid", DataRowVersion.Original)}]",
                $"製品名[{GetValue(row, "ProductName", DataRowVersion.Original)}]",
                $"Serial[{GetValue(row, "Serial", DataRowVersion.Original)}]",
                $"UsedID[{GetValue(row, "UsedID", DataRowVersion.Original)}]",
                $"[]",
                $"[]",
                $"[]",
                $"[]",
                $"[]",
                $"[]",
                $"[]",
                $"[]",
                $"[]",
                $"[]",
            ]);
        }
        private void DeleteSerialRow(IDbConnection connection, DataRow row) {
            var sql =
                $"""
                DELETE FROM {Constants.TSerialTableName}
                WHERE rowid = @rowid;
                """;

            connection.Execute(sql, new {
                rowid = row["rowid", DataRowVersion.Original]
            }, _editModeTransaction);
        }

        private static string GetValue(DataRow row, string columnName, DataRowVersion version = DataRowVersion.Current) {
            var value = row[columnName, version];
            return value == DBNull.Value ? "" : value.ToString() ?? "";
        }

        private void CategorySelect(object sender) {
            var selectedRadioButton = (RadioButton)sender;
            if (!selectedRadioButton.Checked) { return; }
            var tag = selectedRadioButton.Tag?.ToString() ?? string.Empty;

            var actionMap = new Dictionary<(int, string), System.Action>
            {
                { (1, "1"), ViewSubstrateRegistrationLog },
                { (1, "2"), ViewSubstrateRegistrationLog },
                { (2, "1"), ViewProductRegistration },
                { (2, "2"), ViewProductRegistration },
                { (2, "3"), ViewSerialLog },
                { (3, "1"), ViewReprintLog }
            };

            if (actionMap.TryGetValue((_radioButtonNumber, tag), out var action)) {
                action();
            }
            else {
                MessageBox.Show("無効な選択です。", "[CategorySelect]エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void HistoryTableFilter() {
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
        }
        private void InventoryAdjustment() {
            var i = DataBaseDataGridView.SelectedCells[0].RowIndex;
            _substrateMaster.SubstrateID = Convert.ToInt32(DataBaseDataGridView.Rows[i].Cells["SubstrateID"].Value);
            _substrateRegisterWork.OrderNumber = DataBaseDataGridView.Rows[i].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
            _substrateRegisterWork.ProductNumber = DataBaseDataGridView.Rows[i].Cells["SubstrateNumber"].Value.ToString() ?? string.Empty;

            using SqliteConnection con = new(GetConnectionRegistration());
            var sql =
                $"""
                SELECT
                    SubstrateID,
                    SubstrateName,
                    SubstrateModel,
                    ProductName,
                    RegType,
                    CheckBox,
                    SerialPrintType
                FROM {Constants.SubstrateTableName}
                WHERE SubstrateID = @SubstrateID;
                """;

            var result = con.QueryFirstOrDefault(sql, new { _substrateMaster.SubstrateID });

            if (result != null) {
                _substrateMaster.SubstrateName = result.SubstrateName ?? string.Empty;
                _substrateMaster.SubstrateModel = result.SubstrateModel ?? string.Empty;
                _substrateMaster.ProductName = result.ProductName ?? string.Empty;
                _substrateMaster.RegType = Convert.ToInt32(result.RegType);
                _substrateMaster.CheckBin = Convert.ToInt32(result.CheckBox?.ToString() ?? "0", 2);
                _substrateMaster.SerialPrintType = Convert.ToInt32(result.SerialPrintType);
            }

            using SubstrateRegistrationWindow window = new(_substrateMaster, _substrateRegisterWork, _appSettings);
            window.ShowDialog(this);
            LoadEvents();
        }
        private void ShowUsedSubstrateDetails() {
            if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }

            var dataForm = new Form {
                Text = "使用基板",
                ShowIcon = false,
                Width = 800,
                Height = 400,
                MaximizeBox = false,
                MinimizeBox = false,
                StartPosition = FormStartPosition.CenterScreen
            };

            var dataGridView = new DataGridView {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            };

            dataForm.Controls.Add(dataGridView);

            using SqliteConnection con = new(GetConnectionRegistration());
            var sql =
                $"""
                SELECT
                    ID,
                    SubstrateName,
                    SubstrateModel,
                    SubstrateNumber,
                    Decrease
                FROM {Constants.VSubstrateTableName}
                WHERE UseID = @ID
                ORDER BY SubstrateModel ASC;
                """;

            var i = DataBaseDataGridView.SelectedCells[0].RowIndex;
            var id = Convert.ToInt32(DataBaseDataGridView.Rows[i].Cells["ID"].Value);

            var results = con.Query(sql, new { ID = id }).ToList();
            var dt = new DataTable();

            if (results.Count != 0) {
                // 動的オブジェクトからDataTableを作成
                var firstRow = results.First() as IDictionary<string, object>;
                foreach (var key in firstRow!.Keys) {
                    dt.Columns.Add(key);
                }

                foreach (var row in results) {
                    var dict = row as IDictionary<string, object>;
                    dt.Rows.Add([.. dict!.Values]);
                }
            }

            dataGridView.DataSource = dt;
            dataForm.ShowDialog();
        }
        private void GenerateReport() {
            try {
                Cursor.Current = Cursors.WaitCursor;

                if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
                var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
                _productRegisterWork.RowID = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells["ID"].Value);
                _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
                _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
                _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
                _productRegisterWork.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
                _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
                _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
                _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;

                ExcelServiceClosedXml.ReportGeneratorClosedXml.GenerateReport(_productMaster, _productRegisterWork);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                Cursor.Current = Cursors.Default;
            }
        }
        // リスト作成
        private void GenerateList() {
            try {
                Cursor.Current = Cursors.WaitCursor;

                if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
                var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
                _productRegisterWork.RowID = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells["ID"].Value);
                _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
                _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
                _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
                _productRegisterWork.Quantity = Convert.ToInt32(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value);
                _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
                _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
                _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;
                _productRegisterWork.Comment = DataBaseDataGridView.Rows[selectRow].Cells["Comment"].Value.ToString() ?? string.Empty;

                ExcelServiceClosedXml.ListGeneratorClosedXml.GenerateList(_productMaster, _productRegisterWork);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                Cursor.Current = Cursors.Default;
            }
        }
        // チェックシート作成
        private void GenerateCheckSheet() {
            try {
                Cursor.Current = Cursors.WaitCursor;

                if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
                var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
                _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
                _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
                _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
                _productRegisterWork.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
                _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
                _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
                _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;

                ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.GenerateCheckSheet(_productMaster, _productRegisterWork);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                Cursor.Current = Cursors.Default;
            }
        }

        private void ClosedEvents() {
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
        private void 在庫調整ToolStripMenuItem_Click(object sender, EventArgs e) { InventoryAdjustment(); }
        private void ShowUsedSubstrateButton_Click(object sender, EventArgs e) { ShowUsedSubstrateDetails(); }
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
