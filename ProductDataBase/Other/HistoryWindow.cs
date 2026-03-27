using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.ExcelService;
using ProductDatabase.Models;
using ProductDatabase.Other;
using System.Data;
using static ProductDatabase.Data.ProductRepository;

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

        // DataGridView の既定の背景色（編集モード終了時に復元する）
        private readonly Color _originalGridBackColor;

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
                        { "OLesSerial", "O-Lesシリアル" },
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
            // 編集モード終了時に復元するために既定の背景色を記録する
            _originalGridBackColor = DataBaseDataGridView.BackgroundColor;
        }

        // ロード時に初期UIを設定しラジオボタンモードに応じた表示制御を行う
        private void LoadEvents() {
            Font = new System.Drawing.Font(_appSettings.FontName, _appSettings.FontSize);

            CategoryRadioButton1.Checked = true;
            CategoryComboBox.SelectedIndex = 0;

            // 管理者のみ「編集」メニューを表示する
            編集ToolStripMenuItem.Visible = _appSettings.IsAdministrator;

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
                    // 再印刷履歴は編集対象外なので「編集」メニューを非表示にする
                    編集ToolStripMenuItem.Visible = false;
                    break;
            }
        }

        // クエリを実行してDataGridViewにバインドしカラムヘッダーを日本語に変換する
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

        // ラジオボタン選択に応じて基板登録履歴または在庫サマリーを表示する
        private void ViewSubstrateRegistrationLog() {
            if (CategoryRadioButton1.Checked) {
                _tableName = "Substrate";
                在庫調整ToolStripMenuItem.Enabled = false;
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

        // フィルター条件に基づいて製品登録履歴を取得しDataGridViewに表示する
        private void ViewProductRegistration() {
            _tableName = "Product";
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

        // フィルター条件に基づいてシリアル番号履歴を取得しDataGridViewに表示する
        private void ViewSerialLog() {
            _tableName = "Serial";
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
                    s.OLesSerial,
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

        // 製品型式でフィルタして再印刷履歴をDataGridViewに表示する
        private void ViewReprintLog() {
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

        // 基板履歴の編集前後の値を操作ログリストに追加する
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

        // 基板履歴テーブルの指定行を更新するUPDATE文を実行する
        private static void UpdateSubstrateRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
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
            }, transaction);
        }

        // 基板履歴の削除対象行の値を操作ログリストに追加する
        private void LogSubstrateDelete(DataRow row, List<string[]> pendingLogs) {
            pendingLogs.Add([
                "[基板履歴削除]",
                $"[{_substrateMaster.CategoryName}]",
                $"ID[{GetValue(row, "ID")}]",
                $"注文番号[{GetValue(row, "OrderNumber")}]",
                $"製造番号[{GetValue(row, "SubstrateNumber")}]",
                $"[]",
                $"製品名[{GetValue(row,"ProductName")}]",
                $"基板名[{GetValue(row,"SubstrateName")}]",
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

        // 基板履歴テーブルの指定行を論理削除（IsDeleted=1）する
        private static void DeleteSubstrateRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql =
                $"""
                UPDATE {Constants.TSubstrateTableName}
                SET
                    IsDeleted = 1,
                    DeletedAt = datetime('now', 'localtime')
                WHERE ID = @ID;
                """;

            connection.Execute(sql, new {
                ID = row["ID"]
            }, transaction);
        }

        // 製品履歴テーブルの編集可能フィールドをUPDATE文で更新する
        private static void UpdateProductRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
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
            }, transaction);
        }

        // 製品履歴の編集前後の値を操作ログリストに追加する
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

        // 製品履歴の削除対象行の値を操作ログリストに追加する
        private void LogProductDelete(DataRow row, List<string[]> pendingLogs) {
            pendingLogs.Add([
                "[製品履歴削除]",
                $"[{_productMaster.CategoryName}]",
                $"ID[{GetValue(row, "ID")}]",
                $"注文番号[{GetValue(row, "OrderNumber")}]",
                $"製造番号[{GetValue(row, "ProductNumber")}]",
                $"OLes番号[{GetValue(row, "OLesNumber")}]",
                $"製品名[{GetValue(row,"ProductName")}]",
                $"タイプ[{GetValue(row,"ProductType")}]",
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

        // 製品削除に連動して論理削除される基板履歴の内容を操作ログリストに追加する
        private void LogProductSubstrateDelete(IDbConnection connection, DataRow row, List<string[]> pendingLogs, IDbTransaction transaction) {
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
                ID = row["ID"]
            }, transaction);

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

        // 製品削除に連動して削除されるシリアル履歴の内容を操作ログリストに追加する
        private void LogProductSerialDelete(IDbConnection connection, DataRow row, List<string[]> pendingLogs, IDbTransaction transaction) {
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
                ID = row["ID"]
            }, transaction);

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

        // 製品履歴テーブルの指定行を論理削除（IsDeleted=1）する
        private static void DeleteProductRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql =
                $"""
                UPDATE {Constants.TProductTableName}
                SET
                    IsDeleted = 1,
                    DeletedAt = datetime('now', 'localtime')
                WHERE ID = @ID;
                """;

            connection.Execute(sql, new {
                ID = row["ID"]
            }, transaction);
        }

        // 製品削除に連動してUseIDが一致する基板履歴を論理削除する
        private static void DeleteProductSubstrateRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql = $"""
        UPDATE {Constants.TSubstrateTableName}
        SET
            IsDeleted = 1,
            DeletedAt = datetime('now', 'localtime')
        WHERE UseID = @ID;
        """;

            connection.Execute(sql, new {
                ID = row["ID"]
            }, transaction);
        }

        // 製品削除に連動してUsedIDが一致するシリアルを物理削除する
        private static void DeleteProductSerialRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql =
                $"""
                DELETE FROM {Constants.TSerialTableName}
                WHERE UsedID = @ID;
                """;

            connection.Execute(sql, new {
                ID = row["ID"]
            }, transaction);
        }

        // シリアル履歴の削除対象行の値を操作ログリストに追加する
        private void LogSerialDelete(DataRow row, List<string[]> pendingLogs) {
            pendingLogs.Add([
            $"[シリアル履歴削除]",
                $"[{_productMaster.CategoryName}]",
                $"ID[{GetValue(row, "rowid")}]",
                $"製品名[{GetValue(row, "ProductName")}]",
                $"Serial[{GetValue(row, "Serial")}]",
                $"UsedID[{GetValue(row, "UsedID")}]",
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

        // シリアルテーブルから指定rowidの行を物理削除する
        private static void DeleteSerialRow(IDbConnection connection, DataRow row, IDbTransaction transaction) {
            var sql =
                $"""
                DELETE FROM {Constants.TSerialTableName}
                WHERE rowid = @rowid;
                """;

            connection.Execute(sql, new {
                rowid = row["rowid"]
            }, transaction);
        }

        // DataRowから指定カラムの値を文字列で取得しDBNull の場合は空文字を返す
        private static string GetValue(DataRow row, string columnName, DataRowVersion version = DataRowVersion.Current) {
            var value = row[columnName, version];
            return value == DBNull.Value ? "" : value.ToString() ?? "";
        }

        // 選択行の製品履歴を編集ダイアログで編集してDBに保存する
        private async Task EditProductRecord() {
            if (DataBaseDataGridView.CurrentCell is null) { return; }
            var rowIndex = DataBaseDataGridView.CurrentCell.RowIndex;
            if (rowIndex < 0) { return; }

            var dgvRow = DataBaseDataGridView.Rows[rowIndex];

            using var dialog = new Other.HistoryEditDialog(dgvRow);
            if (dialog.ShowDialog(this) != DialogResult.OK) { return; }

            // DataRowを取得してBeginEdit/EndEditで変更を記録（Original/Currentが使えるようになる）
            if (dgvRow.DataBoundItem is not DataRowView drv) { return; }
            var row = drv.Row;

            // UIスレッドでDataRowの変更を記録（Task.Run前に実施）
            row.BeginEdit();
            row["OrderNumber"] = dialog.OrderNumber;
            row["ProductNumber"] = dialog.ProductNumber;
            row["OLesNumber"] = dialog.OLesNumber;
            row["Person"] = dialog.Person;
            row["Comment"] = dialog.Comment;
            row.EndEdit();

            List<string[]> pendingLogs = [];
            using var overlay = new LoadingOverlay(this);
            try {
                // DB操作をバックグラウンドスレッドで実行
                await Task.Run(() => {
                    using var con = new SqliteConnection(GetConnectionRegistration());
                    con.Open();
                    using var tx = con.BeginTransaction();

                    // Original=変更前, Current=変更後
                    LogProductEdit(row, pendingLogs);
                    UpdateProductRow(con, row, tx);

                    CommonUtils.BackupManager.CreateBackup();
                    tx.Commit();

                    foreach (var log in pendingLogs) {
                        CommonUtils.Logger.AppendLog(log);
                    }
                });

                RefreshCurrentView();

            } catch (Exception ex) {
                // エラー時はDataRow変更を取り消す（UIスレッドで実施）
                row.RejectChanges();
                MessageBox.Show(ex.Message, "編集エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 選択行をまとめて削除する（Substrate/Product/Serial対応、複数行選択可）
        private async Task DeleteSelectedRows() {
            // 選択行の収集（SelectedRowsが空の場合はCurrentCellの行を使用）
            var rowsToDelete = new List<DataRow>();
            var selectedRows = DataBaseDataGridView.SelectedRows;
            if (selectedRows.Count > 0) {
                foreach (DataGridViewRow dgvRow in selectedRows) {
                    if (dgvRow.DataBoundItem is DataRowView drv) { rowsToDelete.Add(drv.Row); }
                }
            }
            else {
                if (DataBaseDataGridView.CurrentCell is null) { return; }
                var idx = DataBaseDataGridView.CurrentCell.RowIndex;
                if (DataBaseDataGridView.Rows[idx].DataBoundItem is DataRowView drv) {
                    rowsToDelete.Add(drv.Row);
                }
            }
            if (rowsToDelete.Count == 0) { return; }

            var result = MessageBox.Show(
                $"{rowsToDelete.Count}件を削除しますか？\nこの操作は取り消せません。",
                "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result != DialogResult.Yes) { return; }

            // Task.Run でキャプチャするため変数をローカルにコピー
            var tableName = _tableName;
            List<string[]> pendingLogs = [];
            using var overlay = new LoadingOverlay(this);
            try {
                // DB操作をバックグラウンドスレッドで実行
                await Task.Run(() => {
                    using var con = new SqliteConnection(GetConnectionRegistration());
                    con.Open();
                    using var tx = con.BeginTransaction();

                    switch (tableName) {
                        case "Substrate":
                            foreach (var row in rowsToDelete) {
                                LogSubstrateDelete(row, pendingLogs);
                                DeleteSubstrateRow(con, row, tx);
                            }
                            break;
                        case "Product":
                            foreach (var row in rowsToDelete) {
                                LogProductDelete(row, pendingLogs);
                                LogProductSubstrateDelete(con, row, pendingLogs, tx);
                                LogProductSerialDelete(con, row, pendingLogs, tx);
                                DeleteProductRow(con, row, tx);
                                DeleteProductSubstrateRow(con, row, tx);
                                DeleteProductSerialRow(con, row, tx);
                            }
                            break;
                        case "Serial":
                            foreach (var row in rowsToDelete) {
                                LogSerialDelete(row, pendingLogs);
                                DeleteSerialRow(con, row, tx);
                            }
                            break;
                    }

                    CommonUtils.BackupManager.CreateBackup();
                    tx.Commit();

                    foreach (var log in pendingLogs) {
                        CommonUtils.Logger.AppendLog(log);
                    }
                });

                RefreshCurrentView();

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "削除エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 現在の_tableNameに対応するViewメソッドを呼び出して表示を更新する
        private void RefreshCurrentView() {
            switch (_tableName) {
                case "Substrate":
                    ViewSubstrateRegistrationLog();
                    break;
                case "Product":
                    ViewProductRegistration();
                    break;
                case "Serial":
                    ViewSerialLog();
                    break;
            }
        }

        // ラジオボタンのTagとモード番号の組み合わせで表示する履歴の種類を切り替える
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

        // 選択カラムとキーワードでDataGridViewのRowFilterを設定して表示を絞り込む
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

        // 選択行の基板情報を読み込み基板登録ウィンドウを開いて在庫を調整する
        private void InventoryAdjustment() {
            var i = DataBaseDataGridView.SelectedCells[0].RowIndex;
            _substrateMaster.SubstrateID = int.TryParse(DataBaseDataGridView.Rows[i].Cells["SubstrateID"].Value?.ToString(), out var subId) ? subId : 0;
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
                    Checkbox,
                    SerialPrintType
                FROM {Constants.SubstrateTableName}
                WHERE SubstrateID = @SubstrateID;
                """;

            var result = con.QueryFirstOrDefault(sql, new { _substrateMaster.SubstrateID });

            if (result != null) {
                _substrateMaster.SubstrateName = result.SubstrateName ?? string.Empty;
                _substrateMaster.SubstrateModel = result.SubstrateModel ?? string.Empty;
                _substrateMaster.ProductName = result.ProductName ?? string.Empty;
                _substrateMaster.RegType = int.TryParse(result.RegType?.ToString(), out int rt) ? rt : 0;
                _substrateMaster.CheckBin = Convert.ToInt32(result.Checkbox?.ToString() ?? "0", 2); // バイナリ文字列変換のため維持
                _substrateMaster.SerialPrintType = int.TryParse(result.SerialPrintType?.ToString(), out int spt) ? spt : 0;
            }

            using SubstrateRegistrationWindow window = new(_substrateMaster, _substrateRegisterWork, _appSettings);
            window.ShowDialog(this);
            LoadEvents();
        }

        // 選択行の製品IDに紐づく使用基板一覧をサブウィンドウに表示する
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
            var id = int.TryParse(DataBaseDataGridView.Rows[i].Cells["ID"].Value?.ToString(), out var idVal) ? idVal : 0;

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

        // 選択行の製品情報をセットしてExcel製造報告書を生成する
        private async Task GenerateReport() {
            if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
            var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
            _productRegisterWork.RowID = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["ID"].Value?.ToString(), out var rid1) ? rid1 : 0;
            _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
            _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
            _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
            _productRegisterWork.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
            _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
            _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
            _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;

            // [UIスレッド] 前処理: Config読み込み + ファイル選択 + 保存先選択
            (string templateFilePath, string savePath, ExcelServiceClosedXml.ReportGeneratorClosedXml.ReportConfigClosedXml config)? prepared;
            try {
                prepared = ExcelServiceClosedXml.ReportGeneratorClosedXml.PrepareReport(
                    _productMaster.ProductModel, _productRegisterWork.ProductNumber);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{nameof(GenerateReport)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (prepared is null) return; // キャンセル

            GenerateReportButton.Enabled = false;
            Exception? taskException = null;
            using (var overlay = new LoadingOverlay(this)) {
                try {
                    await CommonUtils.RunOnStaThreadAsync(() =>
                        ExcelServiceClosedXml.ReportGeneratorClosedXml.ExecuteReport(
                            _productMaster, _productRegisterWork,
                            prepared.Value.templateFilePath,
                            prepared.Value.savePath,
                            prepared.Value.config));
                } catch (Exception ex) {
                    taskException = ex;
                }
            }

            GenerateReportButton.Enabled = true;
            if (taskException is not null) {
                MessageBox.Show(taskException.Message, $"[{nameof(GenerateReport)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else {
                MessageBox.Show("成績書が正常に生成されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 選択行の製品情報をセットしてExcel製品一覧を生成する
        private async Task GenerateList() {
            if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
            var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
            _productRegisterWork.RowID = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["ID"].Value?.ToString(), out var rid2) ? rid2 : 0;
            _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
            _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
            _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
            _productRegisterWork.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
            _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
            _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
            _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;
            _productRegisterWork.Comment = DataBaseDataGridView.Rows[selectRow].Cells["Comment"].Value.ToString() ?? string.Empty;

            GenerateListButton.Enabled = false;
            Exception? taskException = null;
            using (var overlay = new LoadingOverlay(this)) {
                try {
                    await CommonUtils.RunOnStaThreadAsync(() =>
                        ExcelServiceClosedXml.ListGeneratorClosedXml.GenerateList(
                            _productMaster, _productRegisterWork));
                } catch (Exception ex) {
                    taskException = ex;
                }
            }

            GenerateListButton.Enabled = true;
            if (taskException is not null) {
                MessageBox.Show(taskException.Message, $"[{nameof(GenerateList)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 選択行の製品情報をセットしてExcelチェックシートを生成する
        private async Task GenerateCheckSheet() {
            if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
            var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
            _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value.ToString() ?? string.Empty;
            _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value.ToString() ?? string.Empty;
            _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value.ToString() ?? string.Empty;
            _productRegisterWork.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
            _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value.ToString() ?? string.Empty;
            _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value.ToString() ?? string.Empty;
            _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value.ToString() ?? string.Empty;

            // [UIスレッド] 前処理: Config読み込み + InputDialog（温度・湿度）
            (ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.CheckSheetConfigData configData, string temperature, string humidity)? prepared;
            try {
                prepared = ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.PrepareCheckSheet(_productMaster);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{nameof(GenerateCheckSheet)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (prepared is null) return; // キャンセル

            GenerateCheckSheetButton.Enabled = false;
            Exception? taskException = null;
            using (var overlay = new LoadingOverlay(this)) {
                try {
                    await CommonUtils.RunOnStaThreadAsync(() =>
                        ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.ExecuteCheckSheet(
                            _productMaster, _productRegisterWork,
                            prepared.Value.configData,
                            prepared.Value.temperature,
                            prepared.Value.humidity));
                } catch (Exception ex) {
                    taskException = ex;
                }
            }

            GenerateCheckSheetButton.Enabled = true;
            if (taskException is not null) {
                MessageBox.Show(taskException.Message, $"[{nameof(GenerateCheckSheet)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 編集モードを開始する：ContextMenuStrip を有効化して枠線を赤に変更する
        private void EnterEditMode() {
            DataBaseDataGridView.ContextMenuStrip = EditContextMenuStrip;
            // DataGridView の背景色を変えて編集モード中であることを視覚的に示す
            DataBaseDataGridView.BackgroundColor = Color.MistyRose;
            DataBaseDataGridView.EnableHeadersVisualStyles = false;
            DataBaseDataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.LightCoral;
            編集開始ToolStripMenuItem.Visible = false;
            編集終了ToolStripMenuItem.Visible = true;
            SetBottomControlsEnabled(false);
        }

        // 編集モードを終了する：ContextMenuStrip を無効化して外観を元に戻す
        private void ExitEditMode() {
            DataBaseDataGridView.ContextMenuStrip = null;
            DataBaseDataGridView.BackgroundColor = _originalGridBackColor;
            DataBaseDataGridView.EnableHeadersVisualStyles = true;
            編集開始ToolStripMenuItem.Visible = true;
            編集終了ToolStripMenuItem.Visible = false;
            SetBottomControlsEnabled(true);
        }

        // 下部のボタン・チェックボックスの有効/無効を切り替える
        private void SetBottomControlsEnabled(bool enabled) {
            GenerateReportButton.Enabled = enabled;
            GenerateListButton.Enabled = enabled;
            GenerateCheckSheetButton.Enabled = enabled;
            ShowUsedSubstrateButton.Enabled = enabled;
            AllSubstrateCheckBox.Enabled = enabled;
            StockCheckBox.Enabled = enabled;
            GroupModelCheckBox.Enabled = enabled;
        }

        // 右クリックメニューを開く前にテーブル種別・行選択に応じて項目を制御する
        private void EditContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e) {
            // 右クリックした位置の行インデックスを取得する
            var clientPoint = DataBaseDataGridView.PointToClient(Cursor.Position);
            var hit = DataBaseDataGridView.HitTest(clientPoint.X, clientPoint.Y);

            // ヘッダー行や空白領域ではメニューを表示しない
            if (hit.RowIndex < 0) { e.Cancel = true; return; }

            // 右クリックした行が選択されていなければ単一選択に切り替える
            if (!DataBaseDataGridView.Rows[hit.RowIndex].Selected) {
                DataBaseDataGridView.CurrentCell = DataBaseDataGridView[0, hit.RowIndex];
            }

            // テーブル種別に応じてメニュー項目の有効/無効を設定する
            EditContextMenuItem.Enabled = (_tableName == "Product");
            DeleteContextMenuItem.Enabled = (_tableName != string.Empty);
        }

        private void HistoryWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void 編集開始ToolStripMenuItem_Click(object sender, EventArgs e) { EnterEditMode(); }
        private void 編集終了ToolStripMenuItem_Click(object sender, EventArgs e) { ExitEditMode(); }
        private async void EditContextMenuItem_Click(object sender, EventArgs e) { await EditProductRecord(); }
        private async void DeleteContextMenuItem_Click(object sender, EventArgs e) { await DeleteSelectedRows(); }
        private void 在庫調整ToolStripMenuItem_Click(object sender, EventArgs e) { InventoryAdjustment(); }
        private void ShowUsedSubstrateButton_Click(object sender, EventArgs e) { ShowUsedSubstrateDetails(); }
        private async void GenerateReportButton_Click(object sender, EventArgs e) { await GenerateReport(); }
        private async void GenerateListButton_Click(object sender, EventArgs e) { await GenerateList(); }
        private async void GenerateCheckSheetButton_Click(object sender, EventArgs e) { await GenerateCheckSheet(); }
        private void CategoryComboBox_SelectedIndexChanged(object sender, EventArgs e) { HistoryTableFilter(); }
        private void FilterStringTextBox_TextChanged(object sender, EventArgs e) { HistoryTableFilter(); }
        private void CategoryRadioButton_CheckedChanged(object sender, EventArgs e) { CategorySelect(sender); }
        private void StockCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateRegistrationLog(); }
        private void AllSubstrateCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateRegistrationLog(); }
        private void GroupModelCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateRegistrationLog(); }
    }
}
