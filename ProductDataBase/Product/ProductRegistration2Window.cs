using GenCode128;
using ProductDatabase.Other;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class ProductRegistration2Window : Form {

        public ProductPrintSettings ProductPrintSettings { get; set; } = new ProductPrintSettings();
        public string printSettingPath = string.Empty;

        public ProductInformation ProductInfo { get; set; } = new ProductInformation();

        private string[] _useSubstrate = [];

        private int _remainingCount = 0;
        private int _labelProNSerial;
        private int _labelProNumLabelsToPrint;

        private int _pageCount = 1;
        private System.Drawing.Printing.PrintAction _printAction;

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

        public ProductRegistration2Window() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);

                GenerateReportButton.Enabled = false;
                RegisterButton.Enabled = true;
                _useSubstrate = ProductInfo.UseSubstrate.Split(",");
                Array.Sort(_useSubstrate);

                if (ProductInfo.IsSerialGeneration) {
                    _labelProNSerial = ProductInfo.SerialFirstNumber;
                    _serialLastNumber = ProductInfo.SerialFirstNumber + ProductInfo.Quantity - 1;
                }

                var strQuantity = string.Empty;
                switch (ProductInfo.RegType) {
                    case 2:
                    case 4:
                        for (var i = 0; i <= _useSubstrate.GetUpperBound(0); i++) {
                            var objCbx = Controls[_checkBoxNames[i]] as System.Windows.Forms.CheckBox;

                            if (objCbx != null) {
                                objCbx.Enabled = true;
                                objCbx.Checked = true;
                            }

                            var objDgv = Controls[_dataGridViewNames[i]] as DataGridView;
                            if (objDgv != null) {
                                objDgv.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
                                objDgv.Columns[2].ReadOnly = false;
                                objDgv.Columns[3].ReadOnly = false;
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

                            using SQLiteConnection con = new(GetConnectionRegistration());
                            con.Open();

                            using var cmd = con.CreateCommand();
                            // 使用基板表示
                            var selectedRows = ProductInfo.SubstrateDataTable.Select($"SubstrateModel = '{_useSubstrate[i]}'");
                            foreach (var row in selectedRows) {
                                var productName = row["SubstrateName"].ToString() ?? throw new Exception("ProductName is null");
                                if (objCbx != null) {
                                    objCbx.Text = $"{productName} - {_useSubstrate[i]}";
                                }
                            }

                            // 在庫テーブルからデータ取得
                            var intQuantity = ProductInfo.Quantity;
                            var substrateName = string.Empty;
                            //cmd.CommandText = $"""SELECT SubstrateNumber, Stock, SubstrateName FROM "{ProductInfo.StockName}_StockView" WHERE Stock > 0 AND SubstrateModel = @SubstrateModel""";
                            cmd.CommandText = $"""
                                SELECT
                                    SubstrateName,
                                    SubstrateNumber,
                                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                                FROM
                                    {ProductInfo.CategoryName}_Substrate
                                WHERE
                                    StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber NOTNULL
                                GROUP BY
                                    SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                                HAVING
                                    Stock > 0
                                ORDER BY
                                    MIN(ID);
                                """;
                            cmd.Parameters.Add("@StockName", DbType.String).Value = ProductInfo.StockName;
                            cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                            using var dr = cmd.ExecuteReader();
                            while (dr.Read()) {
                                var strSubstrateNumber = $"{dr["SubstrateNumber"]}";
                                var intStock = Convert.ToInt32(dr["Stock"]);
                                if (!string.IsNullOrEmpty($"{dr["substrateName"]}")) { substrateName = $"{dr["substrateName"]}"; }
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
                            if (intQuantity > 0) {
                                strQuantity += $"[{substrateName}]{Environment.NewLine}";
                            }
                        }

                        if (!string.IsNullOrEmpty(strQuantity)) {
                            Activate();
                            MessageBox.Show($"在庫が足りません。{Environment.NewLine}{strQuantity}", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }

                        break;
                    case 3:
                        if (_useSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (var i = 0; i <= _useSubstrate.GetUpperBound(0); i++) {
                            var objCbx = Controls[_checkBoxNames[i]] as System.Windows.Forms.CheckBox;

                            if (objCbx != null) {
                                objCbx.Enabled = true;
                                objCbx.Checked = true;
                            }

                            var objDgv = Controls[_dataGridViewNames[i]] as DataGridView;
                            if (objDgv != null) {
                                objDgv.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
                                objDgv.Columns[2].ReadOnly = false;
                                objDgv.Columns[3].ReadOnly = false;
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

                            using SQLiteConnection con = new(GetConnectionRegistration());
                            con.Open();
                            using var cmd = con.CreateCommand();

                            // 使用基板表示
                            var selectedRows = ProductInfo.SubstrateDataTable.Select($"SubstrateModel = '{_useSubstrate[i]}'");
                            foreach (var row in selectedRows) {
                                var productName = row["SubstrateName"].ToString() ?? throw new Exception("ProductName is null");
                                if (objCbx != null) {
                                    objCbx.Text = $"{productName} - {_useSubstrate[i]}";
                                }
                            }

                            //cmd.CommandText = $"""SELECT * FROM "{ProductInfo.StockName}_StockView" WHERE Stock > 0 AND SubstrateModel = @SubstrateModel""";
                            cmd.CommandText = $"""
                                SELECT
                                    SubstrateName,
                                    SubstrateModel,
                                    SubstrateNumber,
                                    OrderNumber,
                                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                                FROM
                                    {ProductInfo.CategoryName}_Substrate
                                WHERE
                                    StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber NOTNULL
                                GROUP BY
                                    SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                                HAVING
                                    Stock > 0
                                ORDER BY
                                    MIN(ID);
                                """;
                            cmd.Parameters.Add("@StockName", DbType.String).Value = ProductInfo.StockName;
                            cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                            using var dr = cmd.ExecuteReader();
                            while (dr.Read()) {
                                var substrateName = string.Empty;
                                var strSubstrateNumber = $"{dr["SubstrateNumber"]}";
                                var intStock = Convert.ToInt32(dr["Stock"]);
                                if (!string.IsNullOrEmpty($"{dr["substrateName"]}")) { substrateName = $"{dr["substrateName"]}"; }
                                objDgv?.Rows.Add(strSubstrateNumber, intStock);

                                var j = 0;
                                var strOrderNumber = $"{dr["OrderNumber"]}";
                                if (strOrderNumber == ProductInfo.OrderNumber) {
                                    if (objDgv != null) {
                                        var intQuantity = ProductInfo.Quantity;
                                        var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value);
                                        var useValue = intQuantity;
                                        objDgv.Rows[j].Cells[2].Value = ProductInfo.Quantity;
                                        objDgv.Rows[j].Cells[3].Value = true;
                                        // 必要数量分割り当てられたかチェック
                                        if (intQuantity > stockValue) {
                                            strQuantity += $"[{strOrderNumber}][{substrateName}]{Environment.NewLine}";
                                        }
                                    }
                                }
                                j++;
                            }
                        }

                        if (!string.IsNullOrEmpty(strQuantity)) {
                            Activate();
                            MessageBox.Show($"[{ProductInfo.OrderNumber}]の在庫が足りません。{Environment.NewLine}{strQuantity}", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        break;
                    case 9:
                        ServiceLoad();
                        break;
                    default:
                        for (var i = 0; i <= 14; i++) {
                            if (Controls[_checkBoxNames[i]] is CheckBox objCbx) {
                                objCbx.Visible = false;
                            }

                            if (Controls[_dataGridViewNames[i]] is DataGridView objDgv) {
                                objDgv.Visible = false;
                            }
                        }
                        break;
                }
                // 印刷UI設定
                ConfigurePrintSettings();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }
        // 印刷UI設定
        private void ConfigurePrintSettings() {
            SubstrateListPrintButton.Visible = ProductInfo.IsListPrint;
            CheckSheetPrintButton.Visible = ProductInfo.IsCheckSheetPrint;

            SerialPrintPostionLabel.Visible = ProductInfo.IsLabelPrint;
            SerialPrintPostionNumericUpDown.Visible = ProductInfo.IsLabelPrint;

            BarcodePrintPostionLabel.Visible = ProductInfo.IsBarcodePrint;
            BarcodePrintPostionNumericUpDown.Visible = ProductInfo.IsBarcodePrint;

            シリアルラベル印刷プレビューToolStripMenuItem.Enabled = ProductInfo.IsLabelPrint;
            シリアルラベル印刷設定ToolStripMenuItem.Enabled = ProductInfo.IsLabelPrint;
            バーコード印刷プレビューToolStripMenuItem.Enabled = ProductInfo.IsBarcodePrint;
            バーコード印刷設定ToolStripMenuItem.Enabled = ProductInfo.IsBarcodePrint;

            if (ProductInfo.IsLabelPrint || ProductInfo.IsBarcodePrint) {
                LoadSettings();
            }
        }
        private void LoadSettings() {
            try {
                ProductPrintSettings = new ProductPrintSettings();
                printSettingPath = Path.Combine(Environment.CurrentDirectory, "config", "Product", ProductInfo.CategoryName, ProductInfo.ProductName, $"PrintConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.json");
                if (!File.Exists(printSettingPath)) { throw new DirectoryNotFoundException($"印刷用設定ファイルがありません。"); }
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

                if (!NumberCheck() || !QuantityCheck()) { return; }
                if (ProductInfo.IsSerialGeneration) {
                    SerialCheck();
                    GenerateSerialCodes();
                }

                DisableControls();

                //if (!Registration()) {
                //    throw new Exception("登録失敗しました。");
                //}
                Registration();

                //HandleLabelPrinting();
                //HandleBarcodePrinting();

                MessageBox.Show("登録完了");

                HandlePostRegistration();
                GenerateReportButton.Enabled = true;
                //if (RequiresClosing) {
                //    Close();
                //}
                //else {
                //    HandlePostRegistration();
                //}
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Registration() {
            using var con = new SQLiteConnection(GetConnectionRegistration());
            con.Open();
            using var transaction = con.BeginTransaction();

            try {
                switch (ProductInfo.RegType) {
                    case 0:
                        RegisterProductWithoutSerial(con);
                        break;
                    case 1:
                        RegisterProductWithSerial(con);
                        break;
                    case 2:
                    case 3:
                    case 4:
                        RegisterProductWithSubstrate(con);
                        break;
                    case 9:
                        ServiceReg(con);
                        break;
                }

                HandleLabelPrinting();
                HandleBarcodePrinting();
                transaction.Commit();

                LogRegistration(ProductInfo);
                CommonUtils.BackupManager.CreateBackup();

                // 登録チェック
                RegistrationCheck();

            } catch (Exception) {
                if (transaction.Connection != null) { //接続が開いているか確認する。
                    transaction.Rollback();
                }
                throw;
            }
        }
        // 登録
        private void RegisterProductWithoutSerial(SQLiteConnection connection) {
            var command = connection.CreateCommand();
            command.CommandText = $"""
                INSERT INTO "{ProductInfo.CategoryName}_Product"
                    (ProductName, OrderNumber, ProductNumber, ProductType, ProductModel, Quantity,
                    Person, RegDate, Revision, RevisionGroup, Comment)
                VALUES
                    (@ProductName, @OrderNumber, @ProductNumber, @ProductType, @ProductModel, @Quantity,
                    @Person, @RegDate, @Revision, @RevisionGroup, @Comment)
                """;

            AddProductParameters(command, ProductInfo);
            command.ExecuteNonQuery();

            // 最終行取得
            command.CommandText = $"""SELECT MAX(ID) FROM "{ProductInfo.CategoryName}_Product";""";
            ProductInfo.ProductID = Convert.ToInt32(command.ExecuteScalar());
        }
        // 登録+シリアル
        private void RegisterProductWithSerial(SQLiteConnection connection) {
            var command = connection.CreateCommand();
            command.CommandText = $"""
                INSERT INTO "{ProductInfo.CategoryName}_Product"
                    (ProductName, OrderNumber, ProductNumber, ProductType, ProductModel, Quantity,
                    Person, RegDate, Revision, RevisionGroup, SerialFirst, SerialLast, SerialLastNumber, Comment)
                VALUES
                    (@ProductName, @OrderNumber, @ProductNumber, @ProductType, @ProductModel, @Quantity,
                    @Person, @RegDate, @Revision, @RevisionGroup, @SerialFirst, @SerialLast, @SerialLastNumber, @Comment)
                """;

            AddProductParameters(command, ProductInfo);
            command.Parameters.Add("@SerialFirst", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialFirst) ? DBNull.Value : ProductInfo.SerialFirst;
            command.Parameters.Add("@SerialLast", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialLast) ? DBNull.Value : ProductInfo.SerialLast;
            command.Parameters.Add("@SerialLastNumber", DbType.String).Value = ProductInfo.IsSerialGeneration ? _serialLastNumber : DBNull.Value;
            command.ExecuteNonQuery();

            // 最終行取得
            command.CommandText = $"""SELECT MAX(ID) FROM "{ProductInfo.CategoryName}_Product";""";
            ProductInfo.ProductID = Convert.ToInt32(command.ExecuteScalar());

            if (ProductInfo.IsSerialGeneration) {
                foreach (var serial in _strSerial) {
                    command.CommandText = $"""
                        INSERT INTO "{ProductInfo.CategoryName}_Serial"
                            (Serial, UsedID, ProductName)
                        VALUES
                            (@Serial, @ProductId, @ProductName)
                        """;
                    command.Parameters.Clear();
                    command.Parameters.Add("@Serial", DbType.String).Value = serial;
                    command.Parameters.Add("@ProductId", DbType.Int32).Value = ProductInfo.ProductID;
                    command.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;
                    command.ExecuteNonQuery();
                }
            }
        }
        // 登録+シリアル+使用基板
        private void RegisterProductWithSubstrate(SQLiteConnection connection) {
            if (_useSubstrate == null) { throw new Exception("ArrUseSubstrateがnullです。"); }

            var command = connection.CreateCommand();

            // 一時テーブルの内容を削除
            command.CommandText = """DELETE FROM "TempSubstrateReduction";""";
            command.ExecuteNonQuery();

            // ... (基板関連の処理) ...
            for (var i = 0; i < _useSubstrate.Length; i++) {

                var objCbx = Controls[_checkBoxNames[i]] as System.Windows.Forms.CheckBox ?? throw new Exception("objCbxがnullです。");

                if (objCbx.Checked) {
                    var objDgv = Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objCbxがnullです。");
                    var dgvRowCnt = objDgv.Rows.Count;

                    for (var j = 0; j < dgvRowCnt; j++) {
                        var boolCbx = Convert.ToBoolean(objDgv.Rows[j].Cells[3].Value);
                        if (boolCbx) {
                            var substrateName = string.Empty;
                            var substrateModel = string.Empty;
                            var orderNum = string.Empty;
                            var substrateNum = objDgv.Rows[j].Cells[0].Value.ToString() ?? string.Empty;
                            var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);

                            command.CommandText = $"""
                                SELECT
                                    SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                                FROM
                                    {ProductInfo.CategoryName}_Substrate
                                WHERE
                                    StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                                GROUP BY
                                    SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                                ORDER BY
                                    MIN(ID);
                                """;
                            command.Parameters.Add("@StockName", DbType.String).Value = ProductInfo.StockName;
                            command.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                            command.Parameters.Add("@SubstrateNumber", DbType.String).Value = substrateNum;

                            using var dr = command.ExecuteReader();
                            while (dr.Read()) {
                                substrateName = $"{dr["SubstrateName"]}";
                                substrateModel = $"{dr["SubstrateModel"]}";
                                orderNum = $"{dr["OrderNumber"]}";
                            }

                            // 一時テーブルに登録
                            command = connection.CreateCommand();
                            command.CommandText =
                                $"""
                                INSERT INTO "TempSubstrateReduction"
                                    (StockName,SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,
                                    Decrease,UsedProductType,UsedProductNumber,UsedOrderNumber,Person,RegDate,Comment)
                                VALUES
                                    (@StockName,@SubstrateName,@SubstrateModel,@SubstrateNumber,@OrderNumber,
                                    @Decrease,@UsedProductType,@UsedProductNumber,@UsedOrderNumber,@Person,@RegDate,@Comment)
                                """;

                            command.Parameters.Add("@StockName", DbType.String).Value = ProductInfo.StockName;
                            command.Parameters.Add("@SubstrateName", DbType.String).Value = string.IsNullOrWhiteSpace(substrateName) ? DBNull.Value : substrateName;
                            command.Parameters.Add("@SubstrateModel", DbType.String).Value = string.IsNullOrWhiteSpace(substrateModel) ? DBNull.Value : substrateModel;
                            command.Parameters.Add("@SubstrateNumber", DbType.String).Value = objDgv.Rows[j].Cells[0].Value;
                            command.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(orderNum) ? DBNull.Value : orderNum;
                            command.Parameters.Add("@Decrease", DbType.String).Value = 0 - useValue;
                            command.Parameters.Add("@UsedProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                            command.Parameters.Add("@UsedProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                            command.Parameters.Add("@UsedOrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                            command.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                            command.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                            command.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                            command.ExecuteNonQuery();
                        }
                    }
                }
            }

            // 製品テーブルに追加
            command.CommandText = $"""
                INSERT INTO "{ProductInfo.CategoryName}_Product"
                    (ProductName, OrderNumber, ProductNumber, ProductType, ProductModel, Quantity,
                    Person, RegDate, Revision, RevisionGroup, SerialFirst, SerialLast, SerialLastNumber, Comment)
                VALUES
                    (@ProductName, @OrderNumber, @ProductNumber, @ProductType, @ProductModel, @Quantity,
                    @Person, @RegDate, @Revision, @RevisionGroup, @SerialFirst, @SerialLast, @SerialLastNumber, @Comment)
                """;
            AddProductParameters(command, ProductInfo);
            command.Parameters.Add("@SerialFirst", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialFirst) ? DBNull.Value : ProductInfo.SerialFirst;
            command.Parameters.Add("@SerialLast", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialLast) ? DBNull.Value : ProductInfo.SerialLast;
            command.Parameters.Add("@SerialLastNumber", DbType.String).Value = ProductInfo.IsSerialGeneration ? _serialLastNumber : DBNull.Value;

            command.ExecuteNonQuery();

            // 最終行取得
            command.CommandText = $"""SELECT MAX(ID) FROM "{ProductInfo.CategoryName}_Product";""";
            ProductInfo.ProductID = Convert.ToInt32(command.ExecuteScalar());

            // 一時テーブルから基板テーブルにコピー
            command.CommandText =
                $"""
                INSERT INTO "{ProductInfo.CategoryName}_Substrate"
                    (StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber,
                    Decrease, UsedProductType, UsedProductNumber, UsedOrderNumber, Person, RegDate, Comment, UseID)
                SELECT
                    tsr.StockName, tsr.SubstrateName, tsr.SubstrateModel, tsr.SubstrateNumber, tsr.OrderNumber,
                    tsr.Decrease, tsr.UsedProductType, tsr.UsedProductNumber, tsr.UsedOrderNumber, tsr.Person, tsr.RegDate, tsr.Comment, @productRowId
                FROM
                    TempSubstrateReduction tsr
                """;
            command.Parameters.Add("@productRowId", DbType.String).Value = ProductInfo.ProductID;
            command.ExecuteNonQuery();

            if (ProductInfo.IsSerialGeneration) {
                foreach (var serial in _strSerial) {
                    command.CommandText = $"""
                        INSERT INTO "{ProductInfo.CategoryName}_Serial"
                            (Serial, UsedID, ProductName)
                        VALUES
                            (@Serial, @ProductId, @ProductName)
                        """;
                    command.Parameters.Clear();
                    command.Parameters.Add("@Serial", DbType.String).Value = serial;
                    command.Parameters.Add("@ProductId", DbType.Int32).Value = ProductInfo.ProductID;
                    command.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;
                    command.ExecuteNonQuery();
                }
            }
        }
        private static void AddProductParameters(SQLiteCommand command, ProductInformation productInfo) {
            command.Parameters.Add("@ProductName", DbType.String).Value = string.IsNullOrWhiteSpace(productInfo.ProductName) ? DBNull.Value : productInfo.ProductName;
            command.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(productInfo.OrderNumber) ? DBNull.Value : productInfo.OrderNumber;
            command.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(productInfo.ProductNumber) ? DBNull.Value : productInfo.ProductNumber;
            command.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(productInfo.ProductType) ? DBNull.Value : productInfo.ProductType;
            command.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(productInfo.ProductModel) ? DBNull.Value : productInfo.ProductModel;
            command.Parameters.Add("@Quantity", DbType.String).Value = productInfo.Quantity;
            command.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(productInfo.Person) ? DBNull.Value : productInfo.Person;
            command.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(productInfo.RegDate) ? DBNull.Value : productInfo.RegDate;
            command.Parameters.Add("@Revision", DbType.String).Value = string.IsNullOrWhiteSpace(productInfo.Revision) ? DBNull.Value : productInfo.Revision;
            command.Parameters.Add("@RevisionGroup", DbType.String).Value = productInfo.RevisionGroup;
            command.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(productInfo.Comment) ? DBNull.Value : productInfo.Comment;
        }
        // 登録チェック
        private void RegistrationCheck() {
            using SQLiteConnection con = new(GetConnectionRegistration());
            con.Open();

            using var cmd = con.CreateCommand();

            cmd.CommandText = $"""SELECT * FROM {ProductInfo.CategoryName}_Product WHERE Id = @Id;""";
            cmd.Parameters.Add("@Id", DbType.String).Value = ProductInfo.ProductID;

            using var dr = cmd.ExecuteReader();

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
                throw new Exception("登録IDが見つかりません。");
            }
        }
        // ログ出力
        private static void LogRegistration(ProductInformation productInfo) {
            string[] logMessageArray = [
                $"[製品登録]",
                $"ID[{productInfo.CategoryName}_{productInfo.ProductID}]",
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

        private bool NumberCheck() {
            using SQLiteConnection con = new(GetConnectionRegistration());
            con.Open();

            var productModel = string.Empty;

            if (!string.IsNullOrEmpty(ProductInfo.ProductNumber)) {
                // 製番が新規かチェック
                using (var cmd = con.CreateCommand()) {
                    cmd.CommandText = $"""SELECT * FROM "{ProductInfo.CategoryName}_Product" WHERE ProductName = @ProductName AND ProductNumber = @ProductNumber ORDER BY "ID" ASC LIMIT 1""";
                    cmd.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;
                    cmd.Parameters.Add("@ProductNumber", DbType.String).Value = ProductInfo.ProductNumber;

                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        productModel = $"{dr["ProductModel"]}";
                    }
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
                using (var cmd = con.CreateCommand()) {
                    cmd.CommandText = $"""SELECT * FROM "{ProductInfo.CategoryName}_Product" WHERE ProductName = @ProductName AND OrderNumber = @OrderNumber ORDER BY "ID" ASC LIMIT 1""";
                    cmd.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;
                    cmd.Parameters.Add("@OrderNumber", DbType.String).Value = ProductInfo.OrderNumber;

                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        productModel = $"{dr["ProductModel"]}";
                    }
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
                        //MessageBox.Show($"[{ProductInfo.OrderNumber}]は[{productModel}]として登録があります。確認してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //return false;
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
                        if (_useSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
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
        private void SerialCheck() {
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
            else { throw new Exception("PrintType unknown"); }

            List<string> strSerialDuplication = [];
            using (SQLiteConnection con = new(GetConnectionRegistration())) {
                con.Open();

                using var cmd = con.CreateCommand();
                cmd.CommandText = $"""
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
                        "{ProductInfo.CategoryName}_Serial" AS s
                    LEFT JOIN
                        "{ProductInfo.CategoryName}_Product" AS p
                    ON
                        s.UsedID = p.ID
                    WHERE
                        s.ProductName = @ProductName
                    AND
                        s.Serial IN ({string.Join(",", _strSerial.Select((_, i) => $"@Serial{i}"))})
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

                if (!StartPrint()) {
                    throw new OperationCanceledException("キャンセルしました。");
                }
            }
        }
        private void HandleBarcodePrinting() {
            if (ProductInfo.IsBarcodePrint) {
                MessageBox.Show("バーコードラベルを印刷します。");
                _serialType = "Barcode";

                if (!StartPrint()) {
                    throw new OperationCanceledException("キャンセルしました。");
                }
            }
        }
        private void DisableControls() {
            RegisterButton.Enabled = false;
            SerialPrintPostionNumericUpDown.Enabled = false;
            BarcodePrintPostionNumericUpDown.Enabled = false;
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
            public string[] ServiveUseSubstrate { get; set; } = [];
        }
        public ServiceInformation ServiceInfo { get; set; } = new();
        private void ServiceLoad() {
            using ServiceForm window = new(ServiceInfo);
            var result = window.ShowDialog(this);
            if (result != DialogResult.OK) { Close(); }

            ServiceInfo = window.ServiceInfo;
            var serviceCategoryName = ServiceInfo.ServiceCategoryName;
            var serviceStockName = ServiceInfo.ServiceStockName;
            var serviveUseSubstrate = ServiceInfo.ServiveUseSubstrate;

            for (var i = 0; i <= serviveUseSubstrate.GetUpperBound(0); i++) {
                var objCbx = Controls[_checkBoxNames[i]] as System.Windows.Forms.CheckBox;

                if (objCbx != null) {
                    objCbx.Enabled = true;
                    objCbx.Checked = true;
                }

                var objDgv = Controls[_dataGridViewNames[i]] as DataGridView;
                if (objDgv != null) {
                    objDgv.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
                    objDgv.Columns[2].ReadOnly = false;
                    objDgv.Columns[3].ReadOnly = false;
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

                using SQLiteConnection con = new(GetConnectionRegistration());
                con.Open();

                using var cmd = con.CreateCommand();
                // 使用基板表示

                var selectedRows = ProductInfo.SubstrateDataTable.Select($"SubstrateModel = '{serviveUseSubstrate[i]}'");
                foreach (var row in selectedRows) {
                    var productName = row["SubstrateName"].ToString() ?? throw new Exception("ProductName is null");
                    if (objCbx != null) {
                        objCbx.Text = $"{productName} - {serviveUseSubstrate[i]}";
                    }
                }

                // 在庫テーブルからデータ取得
                var substrateName = string.Empty;
                cmd.CommandText = $"""
                                SELECT
                                    SubstrateName,SubstrateNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                                FROM
                                    {serviceCategoryName}_Substrate
                                WHERE
                                    StockName = @StockName AND SubstrateModel = @SubstrateModel
                                GROUP BY
                                    SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                                HAVING
                                    Stock > 0
                                ORDER BY
                                    MIN(ID);
                                """;
                cmd.Parameters.Add("@StockName", DbType.String).Value = serviceStockName;
                cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = serviveUseSubstrate[i];
                using var dr = cmd.ExecuteReader();
                while (dr.Read()) {
                    var strSubstrateNumber = $"{dr["SubstrateNumber"]}";
                    var intStock = Convert.ToInt32(dr["Stock"]);
                    if (!string.IsNullOrEmpty($"{dr["substrateName"]}")) { substrateName = $"{dr["substrateName"]}"; }
                    objDgv?.Rows.Add(strSubstrateNumber, intStock);

                }
            }
        }
        private void ServiceReg(SQLiteConnection con) {
            var serviceCategoryName = ServiceInfo.ServiceCategoryName;
            var serviceStockName = ServiceInfo.ServiceStockName;
            var serviveUseSubstrate = ServiceInfo.ServiveUseSubstrate;
            _ = con.CreateCommand();

            if (serviveUseSubstrate == null) { throw new Exception("ArrUseSubstrateがnullです。"); }

            var command = con.CreateCommand();

            // 一時テーブルの内容を削除
            command.CommandText = """DELETE FROM "TempSubstrateReduction";""";
            command.ExecuteNonQuery();

            for (var i = 0; i <= serviveUseSubstrate.Length; i++) {

                var objCbx = Controls[_checkBoxNames[i]] as System.Windows.Forms.CheckBox ?? throw new Exception("objCbxがnullです。");

                if (objCbx.Checked) {
                    var objDgv = Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objCbxがnullです。");
                    var dgvRowCnt = objDgv.Rows.Count;

                    for (var j = 0; j <= dgvRowCnt - 1; j++) {
                        var boolCbx = Convert.ToBoolean(objDgv.Rows[j].Cells[3].Value);
                        if (boolCbx) {
                            var substrateName = string.Empty;
                            var substrateModel = string.Empty;
                            var orderNum = string.Empty;
                            var substrateNum = objDgv.Rows[j].Cells[0].Value.ToString() ?? string.Empty;
                            var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);

                            command = con.CreateCommand();
                            command.CommandText = $"""
                                SELECT
                                    SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                                FROM
                                    {serviceCategoryName}_Substrate
                                WHERE
                                    StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                                GROUP BY
                                    SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                                ORDER BY
                                    MIN(ID);
                                """;
                            command.Parameters.Add("@StockName", DbType.String).Value = serviceStockName;
                            command.Parameters.Add("@SubstrateModel", DbType.String).Value = serviveUseSubstrate[i];
                            command.Parameters.Add("@SubstrateNumber", DbType.String).Value = substrateNum;

                            using var dr = command.ExecuteReader();
                            while (dr.Read()) {
                                substrateName = $"{dr["SubstrateName"]}";
                                substrateModel = $"{dr["SubstrateModel"]}";
                                orderNum = $"{dr["OrderNumber"]}";
                            }

                            // 一時テーブルに登録
                            command = con.CreateCommand();
                            command.CommandText =
                                $"""
                                INSERT INTO "TempSubstrateReduction"
                                    (StockName,SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,Defect,UsedProductType,UsedProductNumber,UsedOrderNumber,Person,RegDate,Comment)
                                VALUES
                                    (@StockName,@SubstrateName,@SubstrateModel,@SubstrateNumber,@OrderNumber,@Defect,@UsedProductType,@UsedProductNumber,@UsedOrderNumber,@Person,@RegDate,@Comment)
                                """;

                            command.Parameters.Add("@StockName", DbType.String).Value = serviceStockName;
                            command.Parameters.Add("@SubstrateName", DbType.String).Value = string.IsNullOrWhiteSpace(substrateName) ? DBNull.Value : substrateName;
                            command.Parameters.Add("@SubstrateModel", DbType.String).Value = string.IsNullOrWhiteSpace(substrateModel) ? DBNull.Value : substrateModel;
                            command.Parameters.Add("@SubstrateNumber", DbType.String).Value = objDgv.Rows[j].Cells[0].Value;
                            command.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(orderNum) ? DBNull.Value : orderNum;
                            command.Parameters.Add("@Defect", DbType.String).Value = 0 - useValue;
                            command.Parameters.Add("@UsedProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                            command.Parameters.Add("@UsedProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                            command.Parameters.Add("@UsedOrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                            command.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                            command.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                            command.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                            command.ExecuteNonQuery();
                        }
                    }
                }
            }

            // 製品テーブルに追加
            command = con.CreateCommand();
            command.CommandText =
                $"""
                INSERT INTO "{ProductInfo.CategoryName}_Product"
                    (ProductName,OrderNumber,ProductNumber,ProductType,ProductModel,Quantity,Person,RegDate,Revision,RevisionGroup,SerialFirst,SerialLast,SerialLastNumber,Comment)
                VALUES
                    (@ProductName,@OrderNumber,@ProductNumber,@ProductType,@ProductModel,@Quantity,@Person,@RegDate,@Revision,@RevisionGroup,@SerialFirst,@SerialLast,@SerialLastNumber,@Comment)
                """;

            command.Parameters.Add("@ProductName", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductName) ? DBNull.Value : ProductInfo.ProductName;
            command.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
            command.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
            command.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
            command.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
            command.Parameters.Add("@Quantity", DbType.String).Value = ProductInfo.Quantity;
            command.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
            command.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
            command.Parameters.Add("@Revision", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Revision) ? DBNull.Value : ProductInfo.Revision;
            command.Parameters.Add("@RevisionGroup", DbType.String).Value = ProductInfo.RevisionGroup;
            command.Parameters.Add("@SerialFirst", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialFirst) ? DBNull.Value : ProductInfo.SerialFirst;
            command.Parameters.Add("@SerialLast", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialLast) ? DBNull.Value : ProductInfo.SerialLast;
            command.Parameters.Add("@SerialLastNumber", DbType.String).Value = ProductInfo.IsSerialGeneration ? _serialLastNumber : DBNull.Value;
            command.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

            command.ExecuteNonQuery();

            // 最終行取得
            command.CommandText = $"""SELECT MAX(ID) FROM "{ProductInfo.CategoryName}_Product";""";
            ProductInfo.ProductID = Convert.ToInt32(command.ExecuteScalar());

            // 一時テーブルから基板テーブルにコピー
            command = con.CreateCommand();
            command.CommandText =
                $"""
                INSERT INTO "{serviceCategoryName}_Substrate"
                    (StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber, Defect, UsedProductType, UsedProductNumber, UsedOrderNumber, Person, RegDate, Comment)
                SELECT
                    tsr.StockName, tsr.SubstrateName, tsr.SubstrateModel, tsr.SubstrateNumber, tsr.OrderNumber, tsr.Defect, tsr.UsedProductType, tsr.UsedProductNumber, tsr.UsedOrderNumber, tsr.Person, tsr.RegDate, tsr.Comment
                FROM
                    TempSubstrateReduction tsr
                """;
            command.ExecuteNonQuery();

            if (ProductInfo.IsSerialGeneration) {
                foreach (var b in _strSerial) {
                    command.CommandText =
                        $"""
                        INSERT INTO "{ProductInfo.CategoryName}_Serial"
                            (Serial,UsedID,ProductName)
                        VALUES
                            (@Serial,@productRowId,@ProductName)
                        """;

                    command.Parameters.Add("@Serial", DbType.String).Value = b;
                    command.Parameters.Add("@productRowId", DbType.String).Value = ProductInfo.ProductID;
                    command.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;

                    command.ExecuteNonQuery();
                }
            }
        }

        // 印刷処理
        private bool StartPrint() {
            // PrintDocumentオブジェクトの作成
            using System.Drawing.Printing.PrintDocument pd = new();
            pd.BeginPrint += (sender, e) => _printAction = e.PrintAction;
            pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

            _labelProNumLabelsToPrint = ProductInfo.Quantity;
            _pageCount = 1;

            ProductRegistration2PrintDialog.Document = pd;

            if (ProductRegistration2PrintDialog.ShowDialog() == DialogResult.OK) {
                // ローディング画面の表示
                using var loadingForm = new LoadingForm();
                // 別スレッドで印刷処理を実行
                Task.Run(() => {
                    try {
                        ProductRegistration2PrintDialog.Document.Print();
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
        }
        private bool PreviewPrint() {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();
                pd.BeginPrint += (sender, e) => _printAction = e.PrintAction;
                pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

                _labelProNumLabelsToPrint = ProductInfo.Quantity;
                _pageCount = 1;

                // 最大で表示
                ProductRegistration2PrintPreviewDialog.Shown += (sender, e) => {
                    if (sender is Form form) {
                        form.WindowState = FormWindowState.Maximized;
                    }
                };
                ProductRegistration2PrintPreviewDialog.PrintPreviewControl.Zoom = 3;
                ProductRegistration2PrintPreviewDialog.Document = pd;
                ProductRegistration2PrintPreviewDialog.ShowDialog();

                return true;
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
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;

                switch (_serialType) {
                    case "Label":
                        if (ProductPrintSettings == null) { throw new Exception("ProductPrintSettingsがnullです。"); }
                        if (ProductPrintSettings.LabelPageSettings != null) {
                            labelCountX = ProductPrintSettings.LabelPageSettings.LabelsPerColumn;
                            labelCountY = ProductPrintSettings.LabelPageSettings.LabelsPerRow;
                            labelWidth = ProductPrintSettings.LabelPageSettings.LabelWidth;
                            labelHeight = ProductPrintSettings.LabelPageSettings.LabelHeight;
                            marginX = ProductPrintSettings.LabelPageSettings.MarginX;
                            marginY = ProductPrintSettings.LabelPageSettings.MarginY;
                            intervalX = ProductPrintSettings.LabelPageSettings.IntervalX;
                            intervalY = ProductPrintSettings.LabelPageSettings.IntervalY;
                            headerPositionX = ProductPrintSettings.LabelPageSettings.HeaderPositionX;
                            headerPositionY = ProductPrintSettings.LabelPageSettings.HeaderPositionY;
                            headerString = ConvertHeaderString(ProductPrintSettings.LabelPageSettings.HeaderTextFormat);
                            headerFont = ProductPrintSettings.LabelPageSettings.HeaderFont;
                        }
                        if (ProductPrintSettings.LabelLayoutSettings != null) {
                            copiesPerLabel = ProductPrintSettings.LabelLayoutSettings.CopiesPerLabel;
                        }
                        startLine = (int)SerialPrintPostionNumericUpDown.Value - 1;
                        break;
                    case "Barcode":
                        if (ProductPrintSettings == null) { throw new Exception("ProductPrintSettingsがnullです。"); }
                        if (ProductPrintSettings.BarcodePageSettings != null) {
                            labelCountX = ProductPrintSettings.BarcodePageSettings.LabelsPerColumn;
                            labelCountY = ProductPrintSettings.BarcodePageSettings.LabelsPerRow;
                            labelWidth = ProductPrintSettings.BarcodePageSettings.LabelWidth;
                            labelHeight = ProductPrintSettings.BarcodePageSettings.LabelHeight;
                            marginX = ProductPrintSettings.BarcodePageSettings.MarginX;
                            marginY = ProductPrintSettings.BarcodePageSettings.MarginY;
                            intervalX = ProductPrintSettings.BarcodePageSettings.IntervalX;
                            intervalY = ProductPrintSettings.BarcodePageSettings.IntervalY;
                            headerPositionX = ProductPrintSettings.BarcodePageSettings.HeaderPositionX;
                            headerPositionY = ProductPrintSettings.BarcodePageSettings.HeaderPositionY;
                            headerString = ConvertHeaderString(ProductPrintSettings.BarcodePageSettings.HeaderTextFormat);
                            headerFont = ProductPrintSettings.BarcodePageSettings.HeaderFont;
                        }
                        if (ProductPrintSettings.BarcodeLayoutSettings != null) {
                            copiesPerLabel = ProductPrintSettings.BarcodeLayoutSettings.CopiesPerLabel;
                        }
                        startLine = (int)BarcodePrintPostionNumericUpDown.Value - 1;
                        break;
                    default:
                        break;
                }

                if (labelCountX == 0 || labelCountY == 0 || copiesPerLabel == 0) { throw new Exception("印刷設定が異常です。"); }

                // ハードマージンをミリメートルに変換
                const double MM_PER_HUNDREDTH_INCH = 0.254;
                marginX -= e.PageSettings.HardMarginX * MM_PER_HUNDREDTH_INCH;
                marginY -= e.PageSettings.HardMarginY * MM_PER_HUNDREDTH_INCH;

                // 最初のページのみオフセットを調整
                var verticalOffset = _pageCount == 1 ? startLine * (intervalY + labelHeight) : 0;
                // ヘッダーの描画
                e.Graphics.DrawString(headerString, headerFont, Brushes.Black, (float)headerPositionX, (float)(verticalOffset + headerPositionY));

                if (_pageCount == 1) {
                    _remainingCount = copiesPerLabel;
                    _labelProNSerial = ProductInfo.SerialFirstNumber;
                }
                if (_pageCount >= 2) { startLine = 0; }

                for (var y = startLine; y < labelCountY; y++) {
                    for (var x = 0; x < labelCountX; x++) {
                        var posX = (float)(marginX + (x * (intervalX + labelWidth)));
                        var posY = (float)(marginY + (y * (intervalY + labelHeight)));

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

                        using var labelImage = MakeLabelImage(generatedCode, (int)e.Graphics.DpiX, 1, fontUnderline);
                        e.Graphics.DrawImage(labelImage, posX, posY, (float)labelWidth, (float)labelHeight);

                        _remainingCount--;
                        if (_remainingCount == 0) {
                            _labelProNSerial++;
                            _labelProNumLabelsToPrint--;
                            //印刷するラベルがなくなった場合の処理
                            if (_labelProNumLabelsToPrint <= 0) {
                                // 最終行の行番号を計算
                                var rowOffset = _serialType switch {
                                    "Label" => 2,
                                    "Barcode" => 1,
                                    _ => 0
                                };
                                // 最終行の行番号を表示
                                e.Graphics.DrawString((y + rowOffset).ToString(), headerFont, Brushes.Black, 0, posY);

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
                "Label" => ProductPrintSettings.LabelLayoutSettings?.TextFormat ?? string.Empty,
                "Barcode" => ProductPrintSettings.BarcodeLayoutSettings?.TextFormat ?? string.Empty,
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
        private Bitmap MakeLabelImage(string text, int resolution, int magnitude, bool fontUnderline) {
            Bitmap labelImage = new(1, 1);
            Graphics g;
            SizeF stringSize;
            double sizeX = 0;
            double sizeY = 0;
            double fontSize;
            float textPosX = 0;
            float textPosY = 0;
            var textFont = SystemFonts.DefaultFont;
            const decimal DisplayResolution = 96.0M;
            const int DisplayMagnitude = 3;
            // プレビューかどうかの判定
            var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;

            switch (_serialType) {
                case "Label":
                    if (ProductPrintSettings == null) {
                        throw new Exception("ProductPrintSettingsがnull");
                    }
                    if (ProductPrintSettings.LabelPageSettings != null) {
                        sizeX = ProductPrintSettings.LabelPageSettings.LabelWidth / 25.4 * resolution * magnitude;
                        sizeY = ProductPrintSettings.LabelPageSettings.LabelHeight / 25.4 * resolution * magnitude;
                    }
                    if (ProductPrintSettings.LabelLayoutSettings != null) {
                        textPosX = (float)(ProductPrintSettings.LabelLayoutSettings.TextPositionX / 25.4 * resolution * magnitude);
                        textPosY = (float)(ProductPrintSettings.LabelLayoutSettings.TextPositionY / 25.4 * resolution * magnitude);
                        fontSize = ProductPrintSettings.LabelLayoutSettings.TextFont.SizeInPoints / 72.0 * resolution * magnitude;
                        var style = fontUnderline ? FontStyle.Underline : FontStyle.Regular;
                        textFont = new Font(ProductPrintSettings.LabelLayoutSettings.TextFont.Name, (float)fontSize, style);
                    }

                    labelImage = new((int)sizeX, (int)sizeY);
                    g = Graphics.FromImage(labelImage);
                    // アンチエイリアス処理を改善
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                    // StringFormat を使用して中心に配置
                    var sf = new StringFormat {
                        Alignment = ProductPrintSettings?.LabelLayoutSettings?.AlignTextXCenter ?? false ? StringAlignment.Center : StringAlignment.Near,
                        LineAlignment = ProductPrintSettings?.LabelLayoutSettings?.AlignTextYCenter ?? false ? StringAlignment.Center : StringAlignment.Near
                    };

                    var x = ProductPrintSettings?.LabelLayoutSettings?.AlignTextXCenter ?? false ? 0 : textPosX;
                    var y = ProductPrintSettings?.LabelLayoutSettings?.AlignTextYCenter ?? false ? 0 : textPosY;

                    // 矩形領域を計算 (文字列を配置する領域)
                    var layoutRect = new RectangleF(x, y, labelImage.Width - x, labelImage.Height - y);
                    g.DrawString(text, textFont, Brushes.Black, layoutRect, sf);

                    // プレビュー時、黒枠を描画
                    if (isPreview) {
                        using var p = new Pen(Color.Black, 3);
                        g.DrawRectangle(p, 0, 0, labelImage.Width - 1, labelImage.Height - 1);
                    }
                    g.Dispose();
                    break;
                case "Barcode":
                    float barcodePosX = 0;
                    float barcodePosY = 0;
                    double barcodeHeight = 1;
                    double barcodeMagnitude = 1;
                    if (ProductPrintSettings == null) {
                        throw new Exception("SettingsBarcodeProがnull");
                    }
                    if (ProductPrintSettings.BarcodePageSettings != null) {
                        sizeX = ProductPrintSettings.BarcodePageSettings.LabelWidth / 25.4 * resolution * magnitude;
                        sizeY = ProductPrintSettings.BarcodePageSettings.LabelHeight / 25.4 * resolution * magnitude;
                    }
                    if (ProductPrintSettings.BarcodeLayoutSettings != null) {
                        textPosX = (float)(ProductPrintSettings.BarcodeLayoutSettings.TextPositionX / 25.4 * resolution * magnitude);
                        textPosY = (float)(ProductPrintSettings.BarcodeLayoutSettings.TextPositionY / 25.4 * resolution * magnitude);
                        fontSize = ProductPrintSettings.BarcodeLayoutSettings.TextFont.SizeInPoints / 72.0 * resolution * magnitude;
                        textFont = new Font(ProductPrintSettings.BarcodeLayoutSettings.TextFont.Name, (float)fontSize, FontStyle.Regular);
                        barcodePosX = (float)(ProductPrintSettings.BarcodeLayoutSettings.BarcodePositionX / 25.4 * resolution * magnitude);
                        barcodePosY = (float)(ProductPrintSettings.BarcodeLayoutSettings.BarcodePositionY / 25.4 * resolution * magnitude);
                        barcodeHeight = ProductPrintSettings.BarcodeLayoutSettings.BarcodeHeight / 25.4 * resolution * magnitude;
                        barcodeMagnitude = ProductPrintSettings.BarcodeLayoutSettings.BarcodeMagnitude;
                    }

                    labelImage = new((int)sizeX, (int)sizeY);
                    g = Graphics.FromImage(labelImage);
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit; // アンチエイリアス処理を改善

                    stringSize = TextRenderer.MeasureText(text, textFont);

                    textPosX = ProductPrintSettings?.BarcodeLayoutSettings?.AlignTextXCenter ?? false
                        ? (float)((labelImage.Width / 2) - (stringSize.Width / 2))
                        : textPosX;

                    g.DrawString(text, textFont, Brushes.Black, textPosX, textPosY);

                    var barWeight = resolution == DisplayResolution ? 1 : (int)(1 * resolution / DisplayResolution / DisplayMagnitude);

                    using (var img = Code128Rendering.MakeBarcodeImage(text, barWeight, true)) {
                        var imageWidth = img.Width * barcodeMagnitude;

                        if (imageWidth > labelImage.Width) { MessageBox.Show($"バーコードの幅がラベル幅を超えています{imageWidth}>{labelImage.Width}"); }

                        barcodePosX = ProductPrintSettings?.BarcodeLayoutSettings?.AlignBarcodeXCenter ?? false
                            ? (float)((labelImage.Width / 2) - (imageWidth / 2))
                            : barcodePosX;

                        g.DrawImage(img, barcodePosX, barcodePosY, (float)imageWidth, (float)barcodeHeight);
                        // プレビュー時、黒枠を描画
                        if (isPreview) {
                            var p = new Pen(Color.Black, 3);
                            g.DrawRectangle(p, 0, 0, labelImage.Width - 1, labelImage.Height - 1);
                        }
                        img.Dispose();
                    }
                    g.Dispose();
                    break;
            }
            return labelImage;
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
                CommonUtils.GenerateReport(ProductInfo);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // リスト印刷
        private void GenerateList() {
            try {
                CommonUtils.GenerateList(ProductInfo);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // チェックシート印刷
        private void GenerateCheckSheet() {
            try {
                CommonUtils.GenerateCheckSheet(ProductInfo);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProductRegistration2Window_Load(object sender, EventArgs e) { LoadEvents(); }
        private void RegisterButton_Click(object sender, EventArgs e) { RegisterCheck(); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
        private void SubstrateCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBox_CheckedChanged(sender, e); }
        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e) { Close(); }
        private void シリアルラベル印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Label";
            PreviewPrint();
        }
        private void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Barcode";
            PreviewPrint();
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
            var entries = new[]
                {
                    ("StrProductName", $"{ProductInfo.ProductName}"),
                    ("StrProductModel", $"{ProductInfo.ProductModel}"),
                    ("StrStockName", $"{ProductInfo.StockName}"),
                    ("StrProductType", $"{ProductInfo.ProductType}"),
                    ("StrOrderNumber", $"{ProductInfo.OrderNumber}"),
                    ("StrProductNumber", $"{ProductInfo.ProductNumber}"),
                    ("StrRevision", $"{ProductInfo.Revision}"),
                    ("StrRegDate", $"{ProductInfo.RegDate}"),
                    ("StrPerson", $"{ProductInfo.Person}"),
                    ("IntQuantity", $"{ProductInfo.Quantity}"),
                    ("IntSerialFirstNumber", $"{ProductInfo.SerialFirstNumber}"),
                    ("IntSerialLastNumber", $"{_serialLastNumber}"),
                    ("StrInitial", $"{ProductInfo.Initial}"),
                    ("IntRegType", $"{ProductInfo.RegType}"),
                    ("IntPrintType", $"{ProductInfo.PrintType}"),
                    ("IntSerialDigit", $"{ProductInfo.SerialDigit}")
                };

            const int ColumnWidth = 15;
            var message = string.Join(Environment.NewLine,
                entries.Select(entry => $"{entry.Item1,-ColumnWidth}[{entry.Item2}]"));

            MessageBox.Show(message, "取得情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void GenerateReportButton_Click(object sender, EventArgs e) { GenerateReport(); }
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) { GenerateList(); }
        private void CheckSheetPrintButton_Click(object sender, EventArgs e) { GenerateCheckSheet(); }
        private void ProductRegistration2PrintPreviewDialog_Load(object sender, EventArgs e) {
            var tool = (ToolStrip)ProductRegistration2PrintPreviewDialog.Controls[1];
            tool.Items[0].Visible = false;
        }
    }
}
