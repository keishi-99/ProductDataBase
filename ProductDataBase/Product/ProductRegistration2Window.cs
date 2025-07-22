using ProductDatabase.Other;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using ZXing;
using ZXing.Windows.Compatibility;
using static ProductDatabase.MainWindow;
using static ProductDatabase.Other.CommonUtils;

namespace ProductDatabase {
    public partial class ProductRegistration2Window : Form {

        public ProductPrintSettings ProductPrintSettings { get; set; } = new ProductPrintSettings();
        public LabelPageSettings LabelPageSettings => ProductPrintSettings.LabelPageSettings ?? new LabelPageSettings();
        public BarcodePageSettings BarcodePageSettings => ProductPrintSettings.BarcodePageSettings ?? new BarcodePageSettings();
        public LabelLayoutSettings LabelLayoutSettings => ProductPrintSettings.LabelLayoutSettings ?? new LabelLayoutSettings();
        public BarcodeLayoutSettings BarcodeLayoutSettings => ProductPrintSettings.BarcodeLayoutSettings ?? new BarcodeLayoutSettings();

        public string printSettingPath = string.Empty;

        public ProductInformation ProductInfo { get; set; } = new ProductInformation();

        private string[] _useSubstrate = [];

        private int _remainingCount = 0;
        private int _labelProNSerial;
        private int _labelProNumLabelsToPrint;

        private const float MmPerInch = 25.4f;
        private const float PointsPerInch = 72.0f;

        private int _pageCount = 1;
        private System.Drawing.Printing.PrintAction _printAction;
        private string _printerName = string.Empty;

        private string _serialType = string.Empty;
        private int _serialLastNumber;
        private readonly List<string> _strSerial = [];
        private readonly List<string> _checkBoxNames = [
                        "Substrate1CheckBox", "Substrate2CheckBox", "Substrate3CheckBox", "Substrate4CheckBox","Substrate5CheckBox",
                        "Substrate6CheckBox", "Substrate7CheckBox", "Substrate8CheckBox", "Substrate9CheckBox","Substrate10CheckBox",
                        "Substrate11CheckBox", "Substrate12CheckBox", "Substrate13CheckBox", "Substrate14CheckBox","Substrate15CheckBox"
                        ];
        private readonly List<string> _dataGridViewNames = [
                        "Substrate1DataGridView", "Substrate2DataGridView", "Substrate3DataGridView", "Substrate4DataGridView","Substrate5DataGridView",
                        "Substrate6DataGridView", "Substrate7DataGridView", "Substrate8DataGridView", "Substrate9DataGridView","Substrate10DataGridView",
                        "Substrate11DataGridView", "Substrate12DataGridView", "Substrate13DataGridView", "Substrate14DataGridView","Substrate15DataGridView"
                        ];

        private SQLiteConnection? _sqLiteConnection; // 編集モード用の接続
        private SQLiteTransaction? _sqLiteTransaction; // 編集モード用のトランザクション

        public ProductRegistration2Window() {
            InitializeComponent();
        }

        private void LoadEvents() {
            try {
                _sqLiteConnection = new SQLiteConnection(GetConnectionRegistration());
                _sqLiteConnection.Open();
                _sqLiteTransaction = _sqLiteConnection.BeginTransaction(); // トランザクション開始（ロック）

                SetFont();
                InitializeUIControls();

                if (ProductInfo.IsSerialGeneration) {
                    SetSerialNumbers();
                }

                switch (ProductInfo.RegType) {
                    case 2:
                    case 4:
                    case 3:
                    case 9:
                        if (ProductInfo.RegType == 9) {
                            using ServiceForm window = new(ServiceInfo);
                            if (window.ShowDialog(this) != DialogResult.OK) {
                                return;
                            }
                            ServiceInfo = window.ServiceInfo;
                        }
                        LoadSubstrateData(_sqLiteConnection);
                        break;
                    default:
                        HideAllControls();
                        break;
                }

                ConfigurePrintSettings();

            } catch (SQLiteException ex) {
                MessageBox.Show($"データベースがロックされています。: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }
        private void ClosingEvents() {
            // 編集モードのトランザクションをコミットしてロック解除
            if (_sqLiteTransaction != null) {
                _sqLiteTransaction.Dispose();
                _sqLiteTransaction = null;
            }
            if (_sqLiteConnection != null) {
                _sqLiteConnection.Close();
                _sqLiteConnection.Dispose();
                _sqLiteConnection = null;
            }
        }
        private void SetFont() {
            Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);
        }
        private void InitializeUIControls() {
            GenerateReportButton.Enabled = false;
            RegisterButton.Enabled = true;
            _useSubstrate = ProductInfo.UseSubstrate.Split(",");
            Array.Sort(_useSubstrate);
        }
        private void HideAllControls() {
            for (var i = 0; i <= 14; i++) {
                if (Controls[_checkBoxNames[i]] is CheckBox objCbx) {
                    objCbx.Visible = false;
                }

                if (Controls[_dataGridViewNames[i]] is DataGridView objDgv) {
                    objDgv.Visible = false;
                }
            }
        }
        private void SetSerialNumbers() {
            _labelProNSerial = ProductInfo.SerialFirstNumber;
            _serialLastNumber = ProductInfo.SerialFirstNumber + ProductInfo.Quantity - 1;
        }
        private void LoadSubstrateData(SQLiteConnection connection) {
            // サービス向け登録の場合は、サービス情報を使用する
            var isServiceRegistration = ProductInfo.RegType == 9;
            var useSubstrate = (isServiceRegistration ? ServiceInfo.ServiceUseSubstrate : _useSubstrate)
                ?? throw new Exception("ArrUseSubstrateがnullです。");

            var shortageSubstrateName = string.Empty;

            for (var i = 0; i < useSubstrate.Length; i++) {
                var objCbx = Controls[_checkBoxNames[i]] as CheckBox;
                var objDgv = Controls[_dataGridViewNames[i]] as DataGridView;

                var selectedRows = ProductInfo.SubstrateDataTable.Select($"SubstrateModel = '{useSubstrate[i]}'");
                var substrateName = string.Empty;
                foreach (var row in selectedRows) {
                    substrateName = row["SubstrateName"].ToString() ?? string.Empty;
                }

                SetupCheckBox(objCbx, i, substrateName, useSubstrate);
                SetupDataGridView(objDgv);

                if (FetchAndDisplaySubstrateData(connection, objDgv, i)) {
                    shortageSubstrateName += $"[{substrateName}]{Environment.NewLine}";
                }
            }

            if (!string.IsNullOrEmpty(shortageSubstrateName)) {
                Activate();
                MessageBox.Show($"在庫が足りません。{Environment.NewLine}{shortageSubstrateName}", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private static void SetupCheckBox(CheckBox? objCbx, int index, string substrateName, string[] useSubstrate) {
            if (objCbx != null) {
                objCbx.Enabled = true;
                objCbx.Checked = true;
                objCbx.Text = $"{substrateName} - {useSubstrate[index]}";
            }
        }
        private void SetupDataGridView(DataGridView? objDgv) {
            if (objDgv != null) {
                objDgv.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
                objDgv.Columns[2].ReadOnly = false;
                objDgv.Columns[3].ReadOnly = false;
                SetDataGridViewSize(objDgv);
            }
        }
        private void SetDataGridViewSize(DataGridView objDgv) {
            switch (Font.Size) {
                case 9:
                    objDgv.RowTemplate.Height = 24;
                    objDgv.Columns[0].Width = 130;
                    objDgv.Columns[1].Width = 35;
                    objDgv.Columns[2].Width = 35;
                    objDgv.Columns[3].Width = 24;
                    break;
                case 12:
                    objDgv.RowTemplate.Height = 24;
                    objDgv.Columns[0].Width = 200;
                    objDgv.Columns[1].Width = 50;
                    objDgv.Columns[2].Width = 50;
                    objDgv.Columns[3].Width = 24;
                    break;
                case 14:
                    objDgv.RowTemplate.Height = 30;
                    objDgv.Columns[0].Width = 240;
                    objDgv.Columns[1].Width = 60;
                    objDgv.Columns[2].Width = 60;
                    objDgv.Columns[3].Width = 30;
                    break;
            }
        }
        private bool FetchAndDisplaySubstrateData(SQLiteConnection connection, DataGridView? objDgv, int index) {
            // サービス向け登録の場合は、サービス情報を使用する
            var isServiceRegistration = ProductInfo.RegType == 9;
            var categoryName = (isServiceRegistration ? ServiceInfo.ServiceCategoryName : ProductInfo.CategoryName)
                ?? throw new Exception("CategoryNameがnullです。");
            var substrateTableName = $"[{categoryName}_Substrate]";
            var stockName = (isServiceRegistration ? ServiceInfo.ServiceStockName : ProductInfo.StockName)
                ?? throw new Exception("StockNameがnullです。");
            var useSubstrate = (isServiceRegistration ? ServiceInfo.ServiceUseSubstrate : _useSubstrate)
                ?? throw new Exception("ArrUseSubstrateがnullです。");

            var intQuantity = ProductInfo.Quantity;

            var commandText =
                $"""
                SELECT
                    SubstrateName,
                    SubstrateNumber,
                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                FROM
                    {substrateTableName}
                WHERE
                    StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber NOTNULL
                GROUP BY
                    SubstrateName, SubstrateModel, SubstrateNumber
                HAVING
                    Stock > 0
                ORDER BY
                    MIN(ID)
                ;
                """;

            using var dr = ExecuteReader(connection, commandText, ("@StockName", stockName), ("@SubstrateModel", useSubstrate[index]));

            while (dr.Read()) {
                var strSubstrateNumber = $"{dr["SubstrateNumber"]}";
                var intStock = Convert.ToInt32(dr["Stock"]);
                var substrateName = useSubstrate[index];
                substrateName = $"{dr["SubstrateName"]}";

                objDgv?.Rows.Add(strSubstrateNumber, intStock);

                if (objDgv != null) {
                    if (intQuantity >= intStock) {
                        intQuantity -= intStock;
                        objDgv.Rows[^1].Cells[2].Value = intStock;
                        objDgv.Rows[^1].Cells[3].Value = true;
                    }
                    else {
                        if (intQuantity == 0) {
                            objDgv.Rows[^1].Cells[2].Value = 0;
                        }
                        else {
                            objDgv.Rows[^1].Cells[2].Value = intQuantity;
                            objDgv.Rows[^1].Cells[3].Value = true;
                            intQuantity = 0;
                        }
                    }
                }
            }
            return intQuantity > 0;
        }

        // 印刷UI設定
        private void ConfigurePrintSettings() {
            SubstrateListPrintButton.Visible = ProductInfo.IsListPrint;
            CheckSheetPrintButton.Visible = ProductInfo.IsCheckSheetPrint;

            SerialPrintPositionLabel.Visible = ProductInfo.IsLabelPrint;
            SerialPrintPositionNumericUpDown.Visible = ProductInfo.IsLabelPrint;

            BarcodePrintPositionLabel.Visible = ProductInfo.IsBarcodePrint;
            BarcodePrintPositionNumericUpDown.Visible = ProductInfo.IsBarcodePrint;

            シリアルラベル印刷プレビューToolStripMenuItem.Enabled = ProductInfo.IsLabelPrint;
            シリアルラベル印刷設定ToolStripMenuItem.Enabled = ProductInfo.IsLabelPrint;
            バーコード印刷プレビューToolStripMenuItem.Enabled = ProductInfo.IsBarcodePrint;
            バーコード印刷設定ToolStripMenuItem.Enabled = ProductInfo.IsBarcodePrint;

            if (ProductInfo.IsSerialGeneration) {
                LoadSettings();
            }
        }
        private void LoadSettings() {
            try {
                ProductPrintSettings = new ProductPrintSettings();
                printSettingPath = Path.Combine(Environment.CurrentDirectory, "config", "Product", ProductInfo.CategoryName, ProductInfo.ProductName, $"PrintConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.json");
                if (!File.Exists(printSettingPath)) {
                    throw new DirectoryNotFoundException($"印刷用設定ファイルがありません。");
                }
                var jsonString = File.ReadAllText(printSettingPath);
                ProductPrintSettings = System.Text.Json.JsonSerializer.Deserialize<ProductPrintSettings>(jsonString) ?? new ProductPrintSettings();
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 登録処理
        private void RegisterCheck() {
            try {
                _strSerial.Clear();

                if (_sqLiteConnection == null || _sqLiteTransaction == null) {
                    throw new InvalidOperationException("編集モード用の接続が初期化されていません。");
                }

                if (!NumberCheck(_sqLiteConnection) || !QuantityCheck()) {
                    return;
                }
                if (ProductInfo.IsSerialGeneration) {
                    SerialCheck(_sqLiteConnection);
                    GenerateSerialCodes();
                }

                DisableControls();

                Registration(_sqLiteConnection, _sqLiteTransaction);

                LogRegistration(ProductInfo);
                BackupManager.CreateBackup();

                // 登録チェック
                RegistrationCheck(_sqLiteConnection);

                // 登録完了メッセージ
                MessageBox.Show("登録しました。");

                // 印刷処理
                HandleLabelPrinting();
                HandleBarcodePrinting();

                HandlePostRegistration();
                GenerateReportButton.Enabled = true;

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Registration(SQLiteConnection connection, SQLiteTransaction transaction) {
            try {
                InsertProduct(connection);

                switch (ProductInfo.RegType) {
                    case 0:
                        break;
                    case 1:
                        if (ProductInfo.IsSerialGeneration) {
                            InsertSerial(connection);
                        }
                        break;
                    case 2:
                    case 3:
                    case 4:
                    case 9:
                        if (ProductInfo.IsSerialGeneration) {
                            InsertSerial(connection);
                        }
                        RegisterSubstrate(connection);
                        break;
                    default:
                        throw new Exception("RegType unknown");
                }
                transaction.Commit();

            } catch (Exception) {
                // エラーが発生した場合はトランザクションをロールバック
                transaction.Rollback();
                MessageBox.Show("登録に失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void InsertProduct(SQLiteConnection connection) {
            var productTableName = $"[{ProductInfo.CategoryName}_Product]";
            var commandText =
                $"""
                INSERT INTO {productTableName}
                    (ProductName, OrderNumber, ProductNumber, ProductType, ProductModel,
                    Quantity, Person, RegDate, Revision, RevisionGroup,
                    SerialFirst, SerialLast, SerialLastNumber, Comment)
                VALUES
                    (@ProductName, @OrderNumber, @ProductNumber, @ProductType, @ProductModel,
                    @Quantity, @Person, @RegDate, @Revision, @RevisionGroup,
                    @SerialFirst, @SerialLast, @SerialLastNumber, @Comment)
                ;
                """;

            ExecuteNonQuery(connection, commandText,
                ("@ProductName", ProductInfo.ProductName),
                ("@OrderNumber", ProductInfo.OrderNumber),
                ("@ProductNumber", ProductInfo.ProductNumber),
                ("@ProductType", ProductInfo.ProductType),
                ("@ProductModel", ProductInfo.ProductModel),
                ("@Quantity", ProductInfo.Quantity),
                ("@Person", ProductInfo.Person),
                ("@RegDate", ProductInfo.RegDate),
                ("@Revision", ProductInfo.Revision),
                ("@RevisionGroup", ProductInfo.RevisionGroup),
                ("@SerialFirst", ProductInfo.SerialFirst),
                ("@SerialLast", ProductInfo.SerialLast),
                ("@SerialLastNumber", ProductInfo.IsSerialGeneration ? _serialLastNumber : DBNull.Value),
                ("@Comment", ProductInfo.Comment)
                );

            ProductInfo.ProductID = Convert.ToInt32(ExecuteScalar(connection, $"SELECT MAX(ID) FROM {productTableName};"));
        }
        private void InsertSerial(SQLiteConnection connection) {
            var serialTableName = $"[{ProductInfo.CategoryName}_Serial]";
            var commandText =
                $"""
                INSERT INTO {serialTableName}
                    (Serial, UsedID, ProductName)
                VALUES
                    (@Serial, @productRowId, @ProductName)
                ;
                """;

            foreach (var serial in _strSerial) {
                ExecuteNonQuery(connection, commandText,
                    ("@Serial", serial),
                    ("@productRowId", ProductInfo.ProductID),
                    ("@ProductName", ProductInfo.ProductName)
                );
            }
        }
        private void RegisterSubstrate(SQLiteConnection connection) {
            // サービス向け登録の場合は、サービス情報を使用する
            var isServiceRegistration = ProductInfo.RegType == 9;
            var categoryName = (isServiceRegistration ? ServiceInfo.ServiceCategoryName : ProductInfo.CategoryName)
                ?? throw new Exception("CategoryNameがnullです。");
            var stockName = (isServiceRegistration ? ServiceInfo.ServiceStockName : ProductInfo.StockName)
                ?? throw new Exception("StockNameがnullです。");
            var useSubstrate = (isServiceRegistration ? ServiceInfo.ServiceUseSubstrate : _useSubstrate)
                ?? throw new Exception("ArrUseSubstrateがnullです。");
            int? useID = isServiceRegistration ? null : ProductInfo.ProductID;

            for (var i = 0; i < useSubstrate.Length; i++) {
                if (!(Controls[_checkBoxNames[i]] as CheckBox)?.Checked ?? true) {
                    continue;
                }

                if (Controls[_dataGridViewNames[i]] as DataGridView is not DataGridView objDgv) {
                    throw new Exception("DataGridViewがnullです。");
                }

                foreach (DataGridViewRow row in objDgv.Rows) {
                    if (row.Cells[3].Value is not bool isChecked || !isChecked) {
                        continue;
                    }

                    var substrateNumber = row.Cells[0].Value.ToString() ?? string.Empty;
                    var useValue = Convert.ToInt32(row.Cells[2].Value);
                    var (substrateName, substrateModel, orderNumber) = GetSubstrateInfo(connection, i, categoryName, stockName, substrateNumber, useSubstrate);
                    InsertSubstrate(connection, categoryName, stockName, substrateName, substrateModel, substrateNumber, orderNumber, useValue, useID);
                }
            }
        }
        private static (string substrateName, string substrateModel, string orderNumber) GetSubstrateInfo(SQLiteConnection connection, int index, string categoryName, string stockName, string substrateNumber, string[] useSubstrate) {
                var substrateTableName = $"[{categoryName}_Substrate]";
            var commandText =
                $"""
                SELECT
                    SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber, SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                FROM
                    {substrateTableName}
                WHERE
                    StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                GROUP BY
                    OrderNumber
                ORDER
                    BY ID ASC LIMIT 1
                ;
                """;

            using var dr = ExecuteReader(connection, commandText,
                ("@StockName", stockName),
                ("@SubstrateModel", useSubstrate[index]),
                ("@SubstrateNumber", substrateNumber)
            );

            return dr.Read()
                ? ($"{dr["SubstrateName"]}", $"{dr["SubstrateModel"]}", $"{dr["OrderNumber"]}")
                : (string.Empty, string.Empty, string.Empty);
        }
        private void InsertSubstrate(SQLiteConnection connection, string categoryName, string stockName, string substrateName, string substrateModel, string substrateNumber, string orderNumber, int useValue, int? useID) {
            var substrateTableName = $"[{categoryName}_Substrate]";
            var commandText =
                $"""
                INSERT INTO {substrateTableName}
                    (StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber,
                     Decrease, Person, RegDate, Comment, UseID)
                VALUES
                    (@StockName, @SubstrateName, @SubstrateModel, @SubstrateNumber, @OrderNumber,
                     @Decrease, @Person, @RegDate, @Comment, @UseID)
                ;
                """;

            ExecuteNonQuery(connection, commandText,
                ("@StockName", stockName),
                ("@SubstrateName", substrateName),
                ("@SubstrateModel", substrateModel),
                ("@SubstrateNumber", substrateNumber),
                ("@OrderNumber", orderNumber),
                ("@Decrease", 0 - useValue),
                ("@Person", ProductInfo.Person),
                ("@RegDate", ProductInfo.RegDate),
                ("@Comment", ProductInfo.Comment),
                ("@UseID", useID)
                );
        }

        private static void ExecuteNonQuery(SQLiteConnection connection, string commandText, params (string, object?)[] parameters) {
            using var command = connection.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                // 空の文字列の場合にNULLを設定
                var sqlValue = value is string strValue && string.IsNullOrEmpty(strValue) ? DBNull.Value : value ?? DBNull.Value;
                command.Parameters.Add(name, DbType.String).Value = sqlValue;
            }

            command.ExecuteNonQuery();
        }
        private static object ExecuteScalar(SQLiteConnection connection, string commandText, params (string, object)[] parameters) {
            using var command = connection.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                command.Parameters.Add(name, DbType.String).Value = value ?? DBNull.Value;
            }

            return command.ExecuteScalar() ?? 0;
        }
        private static SQLiteDataReader ExecuteReader(SQLiteConnection connection, string commandText, params (string, object)[] parameters) {
            using var command = connection.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                command.Parameters.Add(name, DbType.String).Value = value ?? DBNull.Value;
            }

            return command.ExecuteReader();
        }

        // 登録チェック
        private void RegistrationCheck(SQLiteConnection connection) {

            var productTableName = $"[{ProductInfo.CategoryName}_Product]";
            var commandText = $@"SELECT * FROM {productTableName} WHERE Id = @Id;";

            using var dr = ExecuteReader(connection, commandText, ("@Id", ProductInfo.ProductID));

            if (dr.HasRows && dr.Read()) {
                // 1行のデータが存在する場合の処理
                if (dr.Read()) {
                    // 2行以上データが存在する場合の処理
                    throw new Exception("登録IDが複数存在します。");
                }
                else {
                    // 1行のみデータが存在する場合の処理
                }
            }
            else {
                // データが存在しない場合の処理
                throw new Exception("登録に失敗しました。IDが見つかりません。");
            }
        }
        // ログ出力
        private static void LogRegistration(ProductInformation productInfo) {
            string[] logMessageArray = [
                $"[製品登録]",
                $"[{productInfo.CategoryName}]",
                $"ID[{productInfo.ProductID}]",
                $"注文番号[{productInfo.OrderNumber}]",
                $"製造番号[{productInfo.ProductNumber}]",
                $"製品名[{productInfo.ProductName}]",
                $"タイプ[{productInfo.ProductType}]",
                $"型式[{productInfo.ProductModel}]",
                $"数量[{productInfo.Quantity}]",
                $"シリアル先頭[{productInfo.SerialFirst}]",
                $"シリアル末尾[{productInfo.SerialLast}]",
                $"Revision[{productInfo.Revision}]",
                $"登録日[{productInfo.RegDate}]",
                $"担当者[{productInfo.Person}]",
                $"コメント[{productInfo.Comment}]"
            ];
            CommonUtils.Logger.AppendLog(logMessageArray);
        }

        private bool NumberCheck(SQLiteConnection connection) {
            var productModel = string.Empty;

            if (!string.IsNullOrEmpty(ProductInfo.ProductNumber)) {
                // 製番が新規かチェック
                var productTableName = $"[{ProductInfo.CategoryName}_Product]";
                var commandText =
                    $"""
                    SELECT
                        *
                    FROM
                        {productTableName}
                    WHERE
                        ProductName = @ProductName AND ProductNumber = @ProductNumber
                    ORDER BY
                        ID ASC LIMIT 1;
                    """;


                using var dr = ExecuteReader(connection, commandText,
                    ("@ProductName", ProductInfo.ProductName),
                    ("@ProductNumber", ProductInfo.ProductNumber)
                    );
                while (dr.Read()) {
                    productModel = $"{dr["ProductModel"]}";
                }

                if (productModel != string.Empty) {
                    if (productModel == ProductInfo.ProductModel) {
                        Activate();
                        var result = MessageBox.Show($"製番[{ProductInfo.ProductNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                        if (result == DialogResult.No) {
                            return false;
                        }
                    }
                    else {
                        Activate();
                        MessageBox.Show($"[{ProductInfo.ProductNumber}]は[{productModel}]として登録があります。確認してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }

            productModel = string.Empty;

            if (!string.IsNullOrEmpty(ProductInfo.OrderNumber)) {
                // 注文番号が新規かチェック
                var productTableName = $"[{ProductInfo.CategoryName}_Product]";
                var commandText =
                    $"""
                    SELECT
                        *
                    FROM
                        {productTableName}
                    WHERE
                        ProductName = @ProductName AND OrderNumber = @OrderNumber
                    ORDER BY
                        ID ASC LIMIT 1;

                    """;

                using var dr = ExecuteReader(connection, commandText,
                    ("@ProductName", ProductInfo.ProductName),
                    ("@OrderNumber", ProductInfo.OrderNumber)
                    );
                while (dr.Read()) {
                    productModel = $"{dr["ProductModel"]}";
                }

                if (productModel != string.Empty) {
                    if (productModel == ProductInfo.ProductModel) {
                        Activate();
                        var result = MessageBox.Show($"注文番号[{ProductInfo.OrderNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                        if (result == DialogResult.No) {
                            return false;
                        }
                    }
                    else {
                        Activate();
                        var result = MessageBox.Show($"[{ProductInfo.OrderNumber}]は[{productModel}]として登録があります。登録しますか？", "", MessageBoxButtons.YesNo);
                        if (result == DialogResult.No) {
                            return false;
                        }
                    }
                }
            }

            return true;
        }
        private bool QuantityCheck() {
            try {
                switch (ProductInfo.RegType) {
                    case 0:
                    case 1:
                        break;
                    case 2:
                    case 3:
                    case 4:
                    case 9:
                        if (_useSubstrate == null) {
                            throw new Exception("ArrUseSubstrateが空です");
                        }
                        for (var i = 0; i <= _useSubstrate.GetUpperBound(0); i++) {

                            var objCbx = Controls[_checkBoxNames[i]] as System.Windows.Forms.CheckBox ?? throw new Exception("objCbxがnullです。");

                            var objDgv = Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objDgvがnullです。");
                            objDgv.Columns[2].ReadOnly = false;
                            objDgv.Columns[3].ReadOnly = false;

                            if (objCbx.Checked) {
                                var intQuantityCheck = ProductInfo.Quantity;
                                var dgvRowCnt = objDgv.Rows.Count;

                                for (var j = 0; j < dgvRowCnt; j++) {
                                    var boolCbx = objDgv.Rows[j].Cells[3].Value != null && (bool)objDgv.Rows[j].Cells[3].Value;
                                    if (boolCbx) {
                                        var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value.ToString());
                                        if (objDgv.Rows[j].Cells[2].Value == null) {
                                            throw new Exception("使用数が入力されていません。");
                                        }
                                        var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value.ToString());
                                        if (useValue <= 0) {
                                            throw new Exception("使用数が0以下になっています。");
                                        }

                                        if (stockValue < useValue) {
                                            throw new Exception("在庫より多い数量が入力されています。");
                                        }
                                        intQuantityCheck -= useValue;
                                    }
                                }

                                if (intQuantityCheck != 0) {
                                    throw new Exception("入力された数量の合計が必要数と一致しません。");
                                }
                            }
                        }
                        break;
                    default:
                        break;
                }
                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private void SerialCheck(SQLiteConnection connection) {
            if (ProductInfo.IsLabelPrint) {
                for (var i = 0; i < ProductInfo.Quantity; i++) {
                    _serialType = "Label";
                    _strSerial.Add(GenerateCode(ProductInfo.SerialFirstNumber + i));
                }
            }
            else if (ProductInfo.IsBarcodePrint) {
                for (var i = 0; i < ProductInfo.Quantity; i++) {
                    _serialType = "Barcode";
                    _strSerial.Add(GenerateCode(ProductInfo.SerialFirstNumber + i));
                }
            }
            else if (ProductInfo.IsSerialGeneration) {
                for (var i = 0; i < ProductInfo.Quantity; i++) {
                    _serialType = "Label";
                    _strSerial.Add(GenerateCode(ProductInfo.SerialFirstNumber + i));
                }
            }
            else {
                throw new Exception("PrintType unknown");
            }

            List<string> strSerialDuplication = [];

            using var cmd = connection.CreateCommand();
            var serialTableName = $"[{ProductInfo.CategoryName}_Serial]";
            var productTableName = $"[{ProductInfo.CategoryName}_Product]";
            cmd.CommandText =
                $"""
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
                    {serialTableName} AS s
                LEFT JOIN
                    {productTableName} AS p
                ON
                    s.UsedID = p.ID
                WHERE
                    s.ProductName = @ProductName
                AND
                    s.Serial IN ({string.Join(",", _strSerial.Select((_, i) => $"@Serial{i}"))})
                ;
                """;
            cmd.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;
            _strSerial
                .Select((serial, i) => new { ParamName = $"@Serial{i}", Value = serial.Trim() })
                .ToList()
                .ForEach(p => cmd.Parameters.Add(p.ParamName, System.Data.DbType.String).Value = p.Value);

            using var dr = cmd.ExecuteReader();
            while (dr.Read()) {
                strSerialDuplication.Add($"{dr["Serial"]}");
            }

            if (strSerialDuplication.Count > 0) {
                var strSQLDuplication = string.Join($"{Environment.NewLine}", strSerialDuplication);
                throw new Exception($"{strSQLDuplication}{Environment.NewLine}は既に使用されているシリアルです。");
            }
        }
        private void HandleLabelPrinting() {
            if (ProductInfo.IsLabelPrint) {
                MessageBox.Show("シリアルラベルを印刷します。");
                _serialType = "Label";

                if (!Print(true)) {
                    throw new OperationCanceledException("キャンセルしました。");
                }
            }
        }
        private void HandleBarcodePrinting() {
            if (ProductInfo.IsBarcodePrint) {
                MessageBox.Show("バーコードラベルを印刷します。");
                _serialType = "Barcode";

                if (!Print(true)) {
                    throw new OperationCanceledException("キャンセルしました。");
                }
            }
        }
        private void DisableControls() {
            RegisterButton.Enabled = false;
            SerialPrintPositionNumericUpDown.Enabled = false;
            BarcodePrintPositionNumericUpDown.Enabled = false;
        }
        private void GenerateSerialCodes() {
            _serialType = ProductInfo.IsBarcodePrint ? "Barcode" : "Label";
            ProductInfo.SerialFirst = GenerateCode(ProductInfo.SerialFirstNumber);
            ProductInfo.SerialLast = GenerateCode(_serialLastNumber);
        }
        private void HandlePostRegistration() {
            foreach (Control control in Controls) {
                switch (control) {
                    case DataGridView dgv:
                        dgv.Enabled = false;
                        break;
                    case System.Windows.Forms.CheckBox chk:
                        chk.Enabled = false;
                        break;
                }
            }
            SubstrateListPrintButton.Enabled = ProductInfo.IsListPrint;
            CheckSheetPrintButton.Enabled = ProductInfo.IsCheckSheetPrint;
        }

        // サービス向け用処理
        public class ServiceInformation {
            public DataTable ServiceDataTable { get; } = new();
            public string ServiceCategoryName { get; set; } = string.Empty;
            public string ServiceProductName { get; set; } = string.Empty;
            public string ServiceStockName { get; set; } = string.Empty;
            public string ServiceProductType { get; set; } = string.Empty;
            public string ServiceProductModel { get; set; } = string.Empty;
            public string[] ServiceUseSubstrate { get; set; } = [];
        }
        public ServiceInformation ServiceInfo { get; set; } = new();

        // 印刷処理
        private bool Print(bool print) {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();

                pd.BeginPrint += (sender, e) => _printAction = e.PrintAction;
                pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

                _printerName = pd.PrinterSettings.PrinterName;
                _labelProNumLabelsToPrint = ProductInfo.Quantity;
                _pageCount = 1;

                switch (print) {
                    case true:
                        //PrintDialogクラスの作成
                        var pdlg = new PrintDialog {
                            Document = pd
                        };
                        if (pdlg.ShowDialog() == DialogResult.OK) {
                            // ローディング画面の表示
                            using var loadingForm = new LoadingForm();
                            // 別スレッドで印刷処理を実行
                            Task.Run(() => {
                                try {
                                    pd.Print();
                                } finally {
                                    // 印刷が終了したらローディング画面を閉じる
                                    loadingForm.Invoke(new Action(() => loadingForm.Close()));
                                }
                            });

                            // ローディング画面をモーダルとして表示
                            loadingForm.ShowDialog();
                        }
                        else {
                            return false;
                        }
                        return true;
                    case false:
                        //PrintPreviewDialogオブジェクトの作成
                        var ppd = new PrintPreviewDialog();
                        ppd.Shown += (sender, e) => {
                            var tool = (ToolStrip)ppd.Controls[1];
                            tool.Items[0].Visible = false;
                            if (sender is Form form) {
                                form.WindowState = FormWindowState.Maximized;
                            }
                        };
                        ppd.PrintPreviewControl.Zoom = 3;
                        ppd.Document = pd;
                        ppd.ShowDialog();

                        return true;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            var headerString = string.Empty;
            var headerFont = SystemFonts.DefaultFont;
            var copiesPerLabel = 0;

            var labelCountX = 0;
            var labelCountY = 0;
            double labelWidth = 0;
            double labelHeight = 0;
            double marginX = 0;
            double marginY = 0;
            double intervalX = 0;
            double intervalY = 0;
            double headerPositionX = 0;
            double headerPositionY = 0;
            var startLine = 0;
            // プレビューかどうかの判定
            var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;
            try {
                if (e.Graphics == null) {
                    throw new Exception("e.Graphicsがnullです。");
                }

                var dpiX = e.Graphics.DpiX;
                var dpiY = e.Graphics.DpiY;
                e.Graphics.PageUnit = GraphicsUnit.Millimeter;

                switch (_serialType) {
                    case "Label":
                        labelCountX = LabelPageSettings.LabelsPerColumn;
                        labelCountY = LabelPageSettings.LabelsPerRow;
                        labelWidth = LabelPageSettings.LabelWidth;
                        labelHeight = LabelPageSettings.LabelHeight;
                        marginX = LabelPageSettings.MarginX;
                        marginY = LabelPageSettings.MarginY;
                        intervalX = LabelPageSettings.IntervalX;
                        intervalY = LabelPageSettings.IntervalY;
                        headerPositionX = LabelPageSettings.HeaderPositionX;
                        headerPositionY = LabelPageSettings.HeaderPositionY;
                        headerString = ConvertHeaderString(LabelPageSettings.HeaderTextFormat);
                        headerFont = LabelPageSettings.HeaderFont;
                        copiesPerLabel = LabelLayoutSettings.CopiesPerLabel;
                        startLine = (int)SerialPrintPositionNumericUpDown.Value - 1;
                        break;
                    case "Barcode":
                        labelCountX = BarcodePageSettings.LabelsPerColumn;
                        labelCountY = BarcodePageSettings.LabelsPerRow;
                        labelWidth = BarcodePageSettings.LabelWidth;
                        labelHeight = BarcodePageSettings.LabelHeight;
                        marginX = BarcodePageSettings.MarginX;
                        marginY = BarcodePageSettings.MarginY;
                        intervalX = BarcodePageSettings.IntervalX;
                        intervalY = BarcodePageSettings.IntervalY;
                        headerPositionX = BarcodePageSettings.HeaderPositionX;
                        headerPositionY = BarcodePageSettings.HeaderPositionY;
                        headerString = ConvertHeaderString(BarcodePageSettings.HeaderTextFormat);
                        headerFont = BarcodePageSettings.HeaderFont;
                        copiesPerLabel = BarcodeLayoutSettings.CopiesPerLabel;
                        startLine = (int)BarcodePrintPositionNumericUpDown.Value - 1;
                        break;
                    default:
                        break;
                }

                if (labelCountX == 0 || labelCountY == 0 || copiesPerLabel == 0) {
                    throw new Exception("印刷設定が異常です。");
                }

                // ハードマージンをミリメートルに変換
                var hardMarginX = 0f;
                var hardMarginY = 0f;
                if (!isPreview) {
                    (hardMarginX, hardMarginY) = _printerName switch {
                        //"TOSHIBA Universal Printer 2" => (5.6f, 1.9f),
                        //"検査課　複合機" => (5.6f, 1.9f),
                        _ => (e.PageSettings.HardMarginX * MmPerInch / 100, e.PageSettings.HardMarginY * MmPerInch / 100)
                    };
                }

                if (_pageCount == 1) {
                    _remainingCount = copiesPerLabel;
                    _labelProNSerial = ProductInfo.SerialFirstNumber;
                }
                if (_pageCount >= 2) {
                    startLine = 0;
                }

                // 最初のページのみオフセットを調整
                var verticalOffset = _pageCount == 1 ? startLine * (intervalY + labelHeight) : 0;
                // ヘッダーの描画
                e.Graphics.DrawString(headerString, headerFont, Brushes.Gray, (float)headerPositionX, (float)(verticalOffset + headerPositionY - hardMarginY));

                for (var y = startLine; y < labelCountY; y++) {
                    for (var x = 0; x < labelCountX; x++) {
                        var posX = (float)(marginX - hardMarginX + (x * (intervalX + labelWidth)));
                        var posY = (float)(marginY - hardMarginY + (y * (intervalY + labelHeight)));

                        // タイプ4で残り1の場合、最後のラベルに下線をつける
                        var fontUnderline = ProductInfo.IsUnderlinePrint && _remainingCount == 1;

                        // シリアル生成、PrintTypeが9かつ最終行の場合は型式下4桁、それ以外はシリアルを生成
                        string generatedCode;
                        if (!ProductInfo.IsLast4Digits || _remainingCount != 1) {
                            generatedCode = GenerateCode(_labelProNSerial); // シリアルコードを生成
                        }
                        else {
                            generatedCode = ProductInfo.ProductModel[^4..]; // 型式の下4桁を使用
                        }

                        using var labelImage = MakeLabelImage(generatedCode, fontUnderline, dpiX, dpiY);
                        e.Graphics.DrawImage(labelImage, posX, posY, (float)labelWidth, (float)labelHeight);

                        _remainingCount--;
                        if (_remainingCount == 0) {
                            _labelProNSerial++;
                            _labelProNumLabelsToPrint--;
                            //印刷するラベルがなくなった場合の処理
                            if (_labelProNumLabelsToPrint <= 0) {
                                // 最終行の行番号を表示
                                var sf = new StringFormat {
                                    Alignment = StringAlignment.Near,
                                    LineAlignment = StringAlignment.Center
                                };
                                var layoutRect = new RectangleF(0, posY, 0, (float)labelHeight);
                                var rowNumber = (y + 1).ToString();
                                e.Graphics.DrawString(rowNumber, headerFont, Brushes.Black, layoutRect, sf);

                                e.HasMorePages = false;
                                _pageCount = 1;
                                _labelProNumLabelsToPrint = 0;
                                return;
                            }
                            _remainingCount = copiesPerLabel;
                        }
                    }
                }

                if (_labelProNumLabelsToPrint > 0) {
                    _pageCount++;
                    e.HasMorePages = true;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        private string ConvertHeaderString(string s) {
            s = s.Replace("%P", ProductInfo.ProductName)
                 .Replace("%T", ProductInfo.ProductModel)
                 .Replace("%D", DateTime.Today.ToShortDateString())
                 .Replace("%M", ProductInfo.ProductNumber)
                 .Replace("%O", ProductInfo.OrderNumber)
                 .Replace("%N", ProductInfo.Quantity.ToString())
                 .Replace("%U", ProductInfo.Person);
            return s;
        }
        private string GenerateCode(int serialCode) {
            var monthCode = DateTime.Parse(ProductInfo.RegDate).ToString("MM");

            monthCode = monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => monthCode
            };

            var outputCode = _serialType switch {
                "Label" => LabelLayoutSettings.TextFormat ?? string.Empty,
                "Barcode" => BarcodeLayoutSettings.TextFormat ?? string.Empty,
                _ => throw new Exception("_serialType unknown")
            };

            outputCode = outputCode.Replace("%Y", DateTime.Parse(ProductInfo.RegDate).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(ProductInfo.RegDate).ToString("MM"))
                                    .Replace("%T", ProductInfo.Initial)
                                    .Replace("%R", ProductInfo.Revision)
                                    .Replace("%M", string.IsNullOrEmpty(monthCode) ? string.Empty : monthCode[^1..])
                                    .Replace("%S", Convert.ToInt32(serialCode).ToString($"D{ProductInfo.SerialDigit}"));
            return outputCode;
        }
        private Bitmap MakeLabelImage(string text, bool fontUnderline, float dpiX, float dpiY) {

            // プレビューモードかどうかを判定
            var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;

            // --- 2. Bitmapの生成とGraphicsの準備 ---
            (var labelWidth, var labelHeight) = _serialType switch {
                "Label" => (LabelPageSettings.LabelWidth, LabelPageSettings.LabelHeight),
                "Barcode" => (BarcodePageSettings.LabelWidth, BarcodePageSettings.LabelHeight),
                _ => throw new InvalidOperationException("Unsupported serial type.")
            };

            // ビットマップのサイズをピクセル単位で計算
            var pixelWidth = (int)(labelWidth / MmPerInch * dpiX);
            var pixelHeight = (int)(labelHeight / MmPerInch * dpiY);

            var labelImage = new Bitmap(pixelWidth, pixelHeight);
            labelImage.SetResolution(dpiX, dpiY);

            // 'using'ステートメントでGraphicsオブジェクトを確実に破棄
            using (var g = Graphics.FromImage(labelImage)) {
                // すべての描画操作をミリメートル単位で行うように設定
                g.PageUnit = GraphicsUnit.Millimeter;

                // 高品質な描画設定
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

                // --- 3. 印刷タイプに応じた描画処理 ---

                switch (_serialType) {
                    case "Label":
                        DrawLabel(g, text, fontUnderline);
                        break;

                    case "Barcode":
                        DrawBarcode(g, text);
                        break;
                }

                // --- 4. プレビュー用の枠を描画 ---

                if (isPreview) {
                    // 0.1mmの黒いペンで枠線を描画
                    using var p = new Pen(Color.Black, 0.1f);
                    g.DrawRectangle(p, 0, 0, (float)labelWidth - 0.1f, (float)labelHeight - 0.1f);
                }
            }

            return labelImage;
        }
        private void DrawLabel(Graphics g, string text, bool fontUnderline) {

            // フォントサイズはポイント単位でそのまま使用
            var fontName = LabelLayoutSettings.TextFont.Name;
            var fontSize = LabelLayoutSettings.TextFont.SizeInPoints;
            var style = fontUnderline ? FontStyle.Underline : FontStyle.Regular;

            using var textFont = new Font(fontName, fontSize, style);
            // テキストの配置設定
            using var sf = new StringFormat {
                Alignment = LabelLayoutSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near,
                LineAlignment = LabelLayoutSettings.AlignTextCenterY ? StringAlignment.Center : StringAlignment.Near
            };
            // 描画領域をミリメートル単位で計算
            var textPosX = LabelLayoutSettings.AlignTextCenterX ? 0f : (float)LabelLayoutSettings.TextPositionX;
            var textPosY = LabelLayoutSettings.AlignTextCenterY ? 0f : (float)LabelLayoutSettings.TextPositionY;

            var pageWidth = (float)LabelPageSettings.LabelWidth;
            var pageHeight = (float)LabelPageSettings.LabelHeight;

            var layoutRect = new RectangleF(textPosX, textPosY, pageWidth - textPosX, pageHeight - textPosY);

            g.DrawString(text, textFont, Brushes.Black, layoutRect, sf);
        }
        private void DrawBarcode(Graphics g, string text) {

            var pageWidth = (float)BarcodePageSettings.LabelWidth;
            var pageHeight = (float)BarcodePageSettings.LabelHeight;

            // --- テキストの描画 ---
            // フォントサイズはポイント単位でそのまま使用
            var fontName = BarcodeLayoutSettings.TextFont.Name;
            var fontSize = BarcodeLayoutSettings.TextFont.SizeInPoints;

            using (var textFont = new Font(fontName, fontSize, FontStyle.Regular)) {
                // テキストの配置設定
                using var sf = new StringFormat {
                    Alignment = BarcodeLayoutSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near
                };

                var textPosX = BarcodeLayoutSettings.AlignTextCenterX ? 0f : (float)BarcodeLayoutSettings.TextPositionX;
                var textPosY = (float)BarcodeLayoutSettings.TextPositionY;

                var layoutRectString = new RectangleF(textPosX, textPosY, pageWidth - textPosX, pageHeight - textPosY);

                g.DrawString(text, textFont, Brushes.Black, layoutRectString, sf);
            }

            // --- バーコードの描画 ---

            // ZXingはピクセル単位で画像を生成するため、mmからpixelへの変換が必要
            // GraphicsオブジェクトのDPI（g.DpiX）を使用して計算する
            var barcodePixelWidth = (int)(BarcodeLayoutSettings.BarcodeWidth / MmPerInch * g.DpiX);
            var barcodePixelHeight = (int)(BarcodeLayoutSettings.BarcodeHeight / MmPerInch * g.DpiY);

            var writer = new BarcodeWriter<Bitmap> {
                Format = BarcodeFormat.CODE_128,
                Options = new ZXing.Common.EncodingOptions {
                    Height = barcodePixelHeight,
                    Width = barcodePixelWidth,
                    PureBarcode = true // テキストを含まないバーコードのみを生成
                },
                Renderer = new BitmapRenderer()
            };

            using var barcodeBitmap = writer.Write(text);
            var barcodePosX = (float)BarcodeLayoutSettings.BarcodePositionX;
            var barcodePosY = (float)BarcodeLayoutSettings.BarcodePositionY;
            var barcodeWidth = (float)BarcodeLayoutSettings.BarcodeWidth;
            var barcodeHeight = (float)BarcodeLayoutSettings.BarcodeHeight;

            // X座標を中央に調整 (ミリメートル単位で計算)
            if (BarcodeLayoutSettings.AlignBarcodeCenterX) {
                barcodePosX = (pageWidth / 2f) - (barcodeWidth / 2f);
            }

            var layoutRectBarcode = new RectangleF(barcodePosX, barcodePosY, barcodeWidth, barcodeHeight);
            // g.DrawImageでミリメートル単位の座標とサイズを指定して描画
            g.DrawImage(barcodeBitmap, layoutRectBarcode);
        }
        // チェックボックスイベント
        private void CheckBox_CheckedChanged(object sender, EventArgs e) {
            var checkBox = (System.Windows.Forms.CheckBox)sender;
            DataGridView dataGridView = new();

            switch (checkBox.Name) {
                case "Substrate1CheckBox":
                    dataGridView = Substrate1DataGridView;
                    break;
                case "Substrate2CheckBox":
                    dataGridView = Substrate2DataGridView;
                    break;
                case "Substrate3CheckBox":
                    dataGridView = Substrate3DataGridView;
                    break;
                case "Substrate4CheckBox":
                    dataGridView = Substrate4DataGridView;
                    break;
                case "Substrate5CheckBox":
                    dataGridView = Substrate5DataGridView;
                    break;
                case "Substrate6CheckBox":
                    dataGridView = Substrate6DataGridView;
                    break;
                case "Substrate7CheckBox":
                    dataGridView = Substrate7DataGridView;
                    break;
                case "Substrate8CheckBox":
                    dataGridView = Substrate8DataGridView;
                    break;
                case "Substrate9CheckBox":
                    dataGridView = Substrate9DataGridView;
                    break;
                case "Substrate10CheckBox":
                    dataGridView = Substrate10DataGridView;
                    break;
                case "Substrate11CheckBox":
                    dataGridView = Substrate11DataGridView;
                    break;
                case "Substrate12CheckBox":
                    dataGridView = Substrate12DataGridView;
                    break;
                case "Substrate13CheckBox":
                    dataGridView = Substrate13DataGridView;
                    break;
                case "Substrate14CheckBox":
                    dataGridView = Substrate14DataGridView;
                    break;
                case "Substrate15CheckBox":
                    dataGridView = Substrate15DataGridView;
                    break;
                default:
                    break;
            }

            dataGridView.Enabled = checkBox.Checked;
            dataGridView.Visible = checkBox.Checked;
            checkBox.ForeColor = checkBox.Checked ? Color.Black : Color.Red;

            if (!checkBox.Checked && !ProductInfo.IsRegType9) {
                MessageBox.Show("チェックがない場合在庫から引き落とされなくなります。", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        // 成績書作成
        private void GenerateReport() {
            try {
                //CommonUtils.GenerateReport(ProductInfo);
                ReportGenerator.GenerateReport(ProductInfo);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // リスト作成
        private void GenerateList() {
            try {
                ListGenerator.GenerateList(ProductInfo);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // チェックシート作成
        private void GenerateCheckSheet() {
            try {
                CheckSheetGenerator.GenerateCheckSheetNPOI(ProductInfo);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 取得情報表示
        private void ShowInfo() {
            var items = new Dictionary<string, string>
                {
                    {"ProductName", $"{ProductInfo.ProductName}"},
                    {"ProductModel", $"{ProductInfo.ProductModel}"},
                    {"StockName", $"{ProductInfo.StockName}"},
                    {"ProductType", $"{ProductInfo.ProductType}"},
                    {"OrderNumber", $"{ProductInfo.OrderNumber}"},
                    {"ProductNumber", $"{ProductInfo.ProductNumber}"},
                    {"Revision", $"{ProductInfo.Revision}"},
                    {"RegDate", $"{ProductInfo.RegDate}"},
                    {"Person", $"{ProductInfo.Person}"},
                    {"Quantity", $"{ProductInfo.Quantity}"},
                    {"SerialFirstNumber", $"{ProductInfo.SerialFirstNumber}"},
                    {"SerialLastNumber", $"{_serialLastNumber}"},
                    {"Initial", $"{ProductInfo.Initial}"},
                    {"RegType", $"{ProductInfo.RegType}"},
                    {"PrintType", $"{ProductInfo.PrintType}"},
                    {"SerialDigit", $"{ProductInfo.SerialDigit}"}
                };

            var form = new Form {
                Text = "取得情報",
                Width = 300,
                Height = 400,
                AutoSize = true,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                ShowInTaskbar = false,
                MaximizeBox = false,
                MinimizeBox = false
            };

            // ListView作成
            var listView = new ListView {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Font = new Font("PlemolJP", ProductInfo.FontSize),
            };

            // 列の追加
            listView.Columns.Add("", 0);   // 値列の幅（調整可）
            listView.Columns.Add("項目", 200, HorizontalAlignment.Right);  // 項目列の幅
            listView.Columns.Add("値", 360);   // 値列の幅（調整可）

            // データを追加
            foreach (var kvp in items) {
                //var item = new ListViewItem(kvp.Key);
                var item = new ListViewItem("");  // ダミー1列目
                item.SubItems.Add(kvp.Key);
                item.SubItems.Add(kvp.Value);
                listView.Items.Add(item);
            }
            form.Controls.Add(listView);

            // フォームのイベントハンドラ
            form.Shown += (_, _) => {
                listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            };
            form.ShowDialog();
        }

        private void ProductRegistration2Window_Load(object sender, EventArgs e) {
            LoadEvents();
        }
        private void ProductRegistration2Window_FormClosing(object sender, FormClosingEventArgs e) {
            ClosingEvents();
        }
        private void RegisterButton_Click(object sender, EventArgs e) {
            RegisterCheck();
        }
        private void CloseButton_Click(object sender, EventArgs e) {
            Close();
        }
        private void SubstrateCheckBox_CheckedChanged(object sender, EventArgs e) {
            CheckBox_CheckedChanged(sender, e);
        }
        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e) {
            Close();
        }
        private void シリアルラベル印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Label";
            Print(false);
        }
        private void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Barcode";
            Print(false);
        }
        private void シリアルラベル印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductPrintSettingsWindow ls = new() {
                ProductInfo = ProductInfo
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void バーコード印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductBarcodeSettingsWindow ls = new() {
                ProductInfo = ProductInfo
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            ShowInfo();
        }
        private void GenerateReportButton_Click(object sender, EventArgs e) {
            GenerateReport();
        }
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) {
            GenerateList();
        }
        private void CheckSheetPrintButton_Click(object sender, EventArgs e) {
            GenerateCheckSheet();
        }

    }
}
