using GenCode128;
using ProductDatabase.Other;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using System.Drawing.Printing;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class ProductRegistration2Window : Form {

        public CSettingsLabelPro SettingsLabelPro { get; set; } = new CSettingsLabelPro();
        public string labelSettingFilePath = string.Empty;

        public CSettingsBarcodePro SettingsBarcodePro { get; set; } = new CSettingsBarcodePro();
        public string barcodeSettingFilePath = string.Empty;

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

        // プロパティ設定
        private bool IsSerialGeneration => IsRegTypeIn(1, 2, 3, 9);
        private bool IsUnderlinePrint => ProductInfo.PrintType == 4 && IsNotRegType9;

        private bool IsLabelPrint => IsPrintTypeIn(1, 3, 4, 5, 6, 7, 9);
        private bool IsBarcodePrint => IsPrintTypeIn(2, 3);
        private bool IsListPrint => IsPrintTypeIn(5, 6) && IsNotRegType9;
        private bool IsCheckSheetPrint => IsPrintTypeIn(6, 7) && IsNotRegType9;
        private bool IsLast4Digits => IsPrintTypeIn(9) && IsNotRegType9;

        // ヘルパーメソッド
        private bool IsRegTypeIn(params int[] values) {
            return values.Contains(ProductInfo.RegType);
        }
        private bool IsPrintTypeIn(params int[] values) {
            return values.Contains(ProductInfo.PrintType);
        }
        private bool IsNotRegType9 => ProductInfo.RegType != 9;

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

                if (IsSerialGeneration) {
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
                                FROM {ProductInfo.CategoryName}_Substrate
                                WHERE StockName = @StockName AND SubstrateModel = @SubstrateModel
                                GROUP BY SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                                HAVING Stock > 0
                                ORDER BY MIN(ID);
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
                                FROM {ProductInfo.CategoryName}_Substrate
                                WHERE StockName = @StockName AND SubstrateModel = @SubstrateModel
                                GROUP BY SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                                HAVING Stock > 0
                                ORDER BY MIN(ID);
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

                LoadSettings(labelSettingFilePath, barcodeSettingFilePath);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }
        private void LoadSettings(string labelSettingFilePath, string barcodeSettingFilePath) {
            try {
                if (labelSettingFilePath != string.Empty) {
                    using StreamReader? srLabel = new(labelSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer serializerLabel = new(typeof(CSettingsLabelPro));
                    if (serializerLabel.Deserialize(srLabel) is CSettingsLabelPro result) { SettingsLabelPro = result; }
                    srLabel?.Close();
                }
                if (barcodeSettingFilePath != string.Empty) {
                    using StreamReader? srBarcode = new(barcodeSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer serializerBarcode = new(typeof(CSettingsBarcodePro));
                    if (serializerBarcode.Deserialize(srBarcode) is CSettingsBarcodePro result) { SettingsBarcodePro = result; }
                    srBarcode?.Close();
                }
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 印刷UI設定
        private void ConfigurePrintSettings() {
            SubstrateListPrintButton.Visible = IsListPrint;
            CheckSheetPrintButton.Visible = IsCheckSheetPrint;

            SerialPrintPostionLabel.Visible = IsLabelPrint;
            SerialPrintPostionNumericUpDown.Visible = IsLabelPrint;

            BarcodePrintPostionLabel.Visible = IsBarcodePrint;
            BarcodePrintPostionNumericUpDown.Visible = IsBarcodePrint;

            ConfigureUI();
        }
        private void ConfigureUI() {
            if (IsLabelPrint) { ConfigureSerialLabelSettings(); }
            if (IsBarcodePrint) { ConfigureBarcodeSettings(); }
            if (IsSerialGeneration && !IsLabelPrint && !IsBarcodePrint) { ConfigureSerialLabelSettings(); }
            SetMenuOptions();
        }
        private void ConfigureSerialLabelSettings() {
            SettingsLabelPro = new CSettingsLabelPro();
            labelSettingFilePath = Path.Combine(Environment.CurrentDirectory, "config", "Product", ProductInfo.CategoryName, ProductInfo.ProductName, $"SerialConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml");
            if (!File.Exists(labelSettingFilePath)) { throw new DirectoryNotFoundException($"ラベル印刷用設定ファイルがありません。"); }
        }
        private void ConfigureBarcodeSettings() {
            SettingsBarcodePro = new CSettingsBarcodePro();
            barcodeSettingFilePath = Path.Combine(Environment.CurrentDirectory, "config", "Product", ProductInfo.CategoryName, ProductInfo.ProductName, $"BarcodeConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml");
            if (!File.Exists(barcodeSettingFilePath)) { throw new DirectoryNotFoundException($"バーコード印刷用設定ファイルがありません。"); }
        }
        private void SetMenuOptions() {
            シリアルラベル印刷プレビューToolStripMenuItem.Enabled = IsLabelPrint;
            シリアルラベル印刷設定ToolStripMenuItem.Enabled = IsLabelPrint;
            バーコード印刷プレビューToolStripMenuItem.Enabled = IsBarcodePrint;
            バーコード印刷設定ToolStripMenuItem.Enabled = IsBarcodePrint;
        }
        // 登録処理
        private void RegisterCheck() {
            try {
                _strSerial.Clear();

                if (!NumberCheck() || !QuantityCheck()) { return; }
                if (IsSerialGeneration) {
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

            } catch (Exception) {
                transaction.Rollback();
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
            command.Parameters.Add("@SerialLastNumber", DbType.String).Value = IsSerialGeneration ? _serialLastNumber : DBNull.Value;
            command.ExecuteNonQuery();

            // 最終行取得
            command.CommandText = $"""SELECT MAX(ID) FROM "{ProductInfo.CategoryName}_Product";""";
            ProductInfo.ProductID = Convert.ToInt32(command.ExecuteScalar());

            if (IsSerialGeneration) {
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
            command.Parameters.Add("@SerialLastNumber", DbType.String).Value = IsSerialGeneration ? _serialLastNumber : DBNull.Value;

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

            if (IsSerialGeneration) {
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
            if (IsLabelPrint) {
                for (var i = 0; i < ProductInfo.Quantity; i++) {
                    _serialType = "Label";
                    _strSerial.Add(GenerateCode(ProductInfo.SerialFirstNumber + i));
                }
            }
            else if (IsBarcodePrint) {
                for (var i = 0; i < ProductInfo.Quantity; i++) {
                    _serialType = "Barcode";
                    _strSerial.Add(GenerateCode(ProductInfo.SerialFirstNumber + i));
                }
            }
            else if (IsSerialGeneration) {
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
            if (IsLabelPrint) {
                MessageBox.Show("シリアルラベルを印刷します。");
                _serialType = "Label";

                if (!PrintBarcode()) {
                    throw new OperationCanceledException("キャンセルしました。");
                }
            }
        }
        private void HandleBarcodePrinting() {
            if (IsBarcodePrint) {
                MessageBox.Show("バーコードラベルを印刷します。");
                _serialType = "Barcode";

                if (!PrintBarcode()) {
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
            _serialType = IsBarcodePrint ? "Barcode" : "Label";
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
            SubstrateListPrintButton.Enabled = IsListPrint;
            CheckSheetPrintButton.Enabled = IsCheckSheetPrint;
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
            command.Parameters.Add("@SerialLastNumber", DbType.String).Value = IsSerialGeneration ? _serialLastNumber : DBNull.Value;
            command.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

            command.ExecuteNonQuery();

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

            if (IsSerialGeneration) {
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
        private bool PrintBarcode() {
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
        private bool PreviewBarcode() {
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
            System.Drawing.Point headerPos = new(0, 0);
            var headerString = string.Empty;
            var headerFooterFont = new Font("ＭＳ Ｐ明朝", 5.25F);
            Point offset;
            var serialCodePrintCopies = 0;

            var maxX = 0;
            var maxY = 0;
            double sizeX = 0;
            double sizeY = 0;
            double offsetX = 0;
            double offsetY = 0;
            double intervalX = 0;
            double intervalY = 0;
            var startLine = 0;
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;

                // プレビューかどうかの判定
                var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;

                switch (_serialType) {
                    case "Label":
                        if (SettingsLabelPro == null) { throw new Exception("SettingsLabelProがnullです。"); }
                        maxX = SettingsLabelPro.LabelProPageSettings.NumLabelsX;
                        maxY = SettingsLabelPro.LabelProPageSettings.NumLabelsY;
                        sizeX = SettingsLabelPro.LabelProPageSettings.SizeX;
                        sizeY = SettingsLabelPro.LabelProPageSettings.SizeY;
                        offsetX = SettingsLabelPro.LabelProPageSettings.OffsetX;
                        offsetY = SettingsLabelPro.LabelProPageSettings.OffsetY;
                        intervalX = SettingsLabelPro.LabelProPageSettings.IntervalX;
                        intervalY = SettingsLabelPro.LabelProPageSettings.IntervalY;
                        headerPos = SettingsLabelPro.LabelProPageSettings.HeaderPos;
                        headerString = ConvertHeaderFooterString(SettingsLabelPro.LabelProPageSettings.HeaderString);
                        headerFooterFont = SettingsLabelPro.LabelProPageSettings.HeaderFooterFont;
                        serialCodePrintCopies = SettingsLabelPro.LabelProLabelSettings.NumLabels;
                        startLine = (int)SerialPrintPostionNumericUpDown.Value - 1;
                        break;
                    case "Barcode":
                        if (SettingsBarcodePro == null) { throw new Exception("SettingsBarcodeProがnullです。"); }
                        maxX = SettingsBarcodePro.BarcodeProPageSettings.NumLabelsX;
                        maxY = SettingsBarcodePro.BarcodeProPageSettings.NumLabelsY;
                        sizeX = SettingsBarcodePro.BarcodeProPageSettings.SizeX;
                        sizeY = SettingsBarcodePro.BarcodeProPageSettings.SizeY;
                        offsetX = SettingsBarcodePro.BarcodeProPageSettings.OffsetX;
                        offsetY = SettingsBarcodePro.BarcodeProPageSettings.OffsetY;
                        intervalX = SettingsBarcodePro.BarcodeProPageSettings.IntervalX;
                        intervalY = SettingsBarcodePro.BarcodeProPageSettings.IntervalY;
                        headerPos = SettingsBarcodePro.BarcodeProPageSettings.HeaderPos;
                        headerString = ConvertHeaderFooterString(SettingsBarcodePro.BarcodeProPageSettings.HeaderString);
                        headerFooterFont = SettingsBarcodePro.BarcodeProPageSettings.HeaderFooterFont;
                        serialCodePrintCopies = SettingsBarcodePro.BarcodeProLabelSettings.NumLabels;
                        startLine = (int)BarcodePrintPostionNumericUpDown.Value;
                        break;
                    default:
                        break;
                }

                if (maxX == 0 || maxY == 0 || serialCodePrintCopies == 0) { throw new Exception("印刷設定が異常です。"); }

                const double MM_PER_HUNDREDTH_INCH = 0.254;

                var pd = (PrintDocument)sender;

                if (!isPreview) {
                    // ハードマージンをミリメートルに変換
                    offsetX -= e.PageSettings.HardMarginX * MM_PER_HUNDREDTH_INCH;
                    offsetY -= e.PageSettings.HardMarginY * MM_PER_HUNDREDTH_INCH;

                    // 最初のページのみオフセットを調整
                    var verticalOffset = _pageCount == 0 ? startLine * (intervalY + sizeY) : 0;

                    // オフセットを計算
                    offset = new System.Drawing.Point(
                        (int)(e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH),
                        (int)((e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + verticalOffset)
                    );
                }
                else {
                    offset = _pageCount == 1
                        ? new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (startLine * (intervalY + sizeY))))
                        : new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (0 * (intervalY + sizeY))));
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;

                headerPos.Offset(offset);
                e.Graphics.DrawString(headerString, headerFooterFont, Brushes.Black, headerPos);

                if (_pageCount == 1) {
                    _remainingCount = serialCodePrintCopies;
                    _labelProNSerial = ProductInfo.SerialFirstNumber;
                }
                if (_pageCount >= 2) { startLine = 0; }

                for (var y = startLine; y < maxY; y++) {
                    for (var x = 0; x < maxX; x++) {
                        var posX = (float)(offsetX + (x * (intervalX + sizeX)));
                        var posY = (float)(offsetY + (y * (intervalY + sizeY)));

                        // タイプ4で残り1の場合、最後のラベルに下線をつける
                        var fontUnderline = IsUnderlinePrint && _remainingCount == 1;

                        // シリアル生成、PrintTypeが9かつ最終行の場合は型式下4桁、それ以外はシリアルを生成
                        string generatedCode;
                        if (!IsLast4Digits || _remainingCount != 1) {
                            generatedCode = GenerateCode(_labelProNSerial); // シリアルコードを生成
                        }
                        else {
                            generatedCode = ProductInfo.ProductModel[^4..]; // 型式の下4桁を使用
                        }

                        using var labelImage = MakeLabelImage(generatedCode, (int)e.Graphics.DpiX, 1, fontUnderline);
                        e.Graphics.DrawImage(labelImage, posX, posY, (float)sizeX, (float)sizeY);

                        _remainingCount--;
                        if (_remainingCount == 0) {
                            _labelProNSerial++;
                            _labelProNumLabelsToPrint--;
                            //印刷するラベルがなくなった場合の処理
                            if (_labelProNumLabelsToPrint <= 0) {
                                e.HasMorePages = false;
                                _pageCount = 1;
                                _labelProNumLabelsToPrint = 0;
                                return;
                            }
                            _remainingCount = serialCodePrintCopies;
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
        private string ConvertHeaderFooterString(string s) {
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
                "Label" => SettingsLabelPro.LabelProLabelSettings.Format,
                "Barcode" => SettingsBarcodePro.BarcodeProLabelSettings.Format,
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
            double sizeX;
            double sizeY;
            double fontSize;
            float stringPosX;
            float stringPosY;
            Font fnt;
            var displayResolution = 96.0M;
            var displayMagnitude = 3;

            // プレビューかどうかの判定
            var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;

            // サイズとフォント情報の設定
            void SetLabelProperties(double labelWidth, double labelHeight, double posY, double fontPointSize, string fontName, bool underlined) {
                sizeX = labelWidth / 25.4 * resolution * magnitude;
                sizeY = labelHeight / 25.4 * resolution * magnitude;
                stringPosY = (float)(posY / 25.4 * resolution * magnitude);
                fontSize = fontPointSize / 72.0 * resolution * magnitude;
                var style = underlined ? FontStyle.Underline : FontStyle.Regular;
                fnt = new Font(fontName, (float)fontSize, style);
            }

            switch (_serialType) {
                case "Label":
                    if (SettingsLabelPro == null) { throw new Exception("SettingsLabelProがnull"); }
                    SetLabelProperties(SettingsLabelPro.LabelProPageSettings.SizeX,
                                        SettingsLabelPro.LabelProPageSettings.SizeY,
                                        SettingsLabelPro.LabelProLabelSettings.StringPosY,
                                        SettingsLabelPro.LabelProLabelSettings.Font.SizeInPoints,
                                        SettingsLabelPro.LabelProLabelSettings.Font.Name,
                                        fontUnderline);

                    labelImage = new((int)sizeX, (int)sizeY);
                    g = Graphics.FromImage(labelImage);
                    // アンチエイリアス処理を改善
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                    // StringFormat を使用して中心に配置
                    var sf = new StringFormat {
                        Alignment = SettingsLabelPro.LabelProLabelSettings.AlignStringXCenter ? StringAlignment.Center : StringAlignment.Near,
                        LineAlignment = SettingsLabelPro.LabelProLabelSettings.AlignStringYCenter ? StringAlignment.Center : StringAlignment.Near
                    };

                    var x = SettingsLabelPro.LabelProLabelSettings.AlignStringXCenter ? 0 : (float)(SettingsLabelPro.LabelProLabelSettings.StringPosX / 25.4 * resolution * magnitude);
                    var y = SettingsLabelPro.LabelProLabelSettings.AlignStringYCenter ? 0 : stringPosY;

                    // 矩形領域を計算 (文字列を配置する領域)
                    var layoutRect = new RectangleF(x, y, labelImage.Width - x, labelImage.Height - y);
                    g.DrawString(text, fnt, Brushes.Black, layoutRect, sf);

                    // プレビュー時、黒枠を描画
                    if (isPreview) {
                        using var p = new Pen(Color.Black, 3);
                        g.DrawRectangle(p, 0, 0, labelImage.Width - 1, labelImage.Height - 1);
                    }

                    g.Dispose();
                    break;
                case "Barcode":
                    if (SettingsBarcodePro == null) { throw new Exception("SettingsBarcodeProがnull"); }
                    SetLabelProperties(SettingsBarcodePro.BarcodeProPageSettings.SizeX,
                                        SettingsBarcodePro.BarcodeProPageSettings.SizeY,
                                        SettingsBarcodePro.BarcodeProLabelSettings.StringPosY,
                                        SettingsBarcodePro.BarcodeProLabelSettings.Font.SizeInPoints,
                                        SettingsBarcodePro.BarcodeProLabelSettings.Font.Name,
                                        false);

                    labelImage = new((int)sizeX, (int)sizeY);
                    g = Graphics.FromImage(labelImage);
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit; // アンチエイリアス処理を改善
                    //stringSize = g.MeasureString(text, fnt);
                    stringSize = TextRenderer.MeasureText(text, fnt);

                    stringPosX = SettingsBarcodePro.BarcodeProLabelSettings.AlignStringCenter
                        ? (float)((labelImage.Width / 2f) - (stringSize.Width / 2f))
                        : (float)(SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4 * resolution * magnitude);

                    g.DrawString(text, fnt, Brushes.Black, stringPosX, stringPosY);

                    var barWeight = resolution == displayResolution ? 1 : (int)(1 * resolution / displayResolution / displayMagnitude);

                    using (var img = Code128Rendering.MakeBarcodeImage(text, barWeight, true)) {
                        var imageWidth = img.Width * SettingsBarcodePro.BarcodeProLabelSettings.BarcodeMagnitude;

                        if (imageWidth > labelImage.Width) { MessageBox.Show($"バーコードの幅がラベル幅を超えています{imageWidth}>{labelImage.Width}"); }

                        var barCodePosX = SettingsBarcodePro.BarcodeProLabelSettings.AlignBarcodeCenter
                            ? (float)((labelImage.Width / 2) - (imageWidth / 2))
                            : (float)(SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4 * resolution * magnitude);
                        var barCodePosY = (float)(SettingsBarcodePro.BarcodeProLabelSettings.BarcodePosY / 25.4 * resolution * magnitude);

                        var barcodeHeight = SettingsBarcodePro.BarcodeProLabelSettings.BarcodeHeight / 25.4 * resolution * magnitude;

                        g.DrawImage(img, barCodePosX, barCodePosY, (float)imageWidth, (float)barcodeHeight);
                        // プレビュー時、黒枠を描画
                        if (isPreview) {
                            var p = new Pen(Color.Black, 3);
                            g.DrawRectangle(p, 0, 0, labelImage.Width - 1, labelImage.Height - 1);
                        }

                        img.Dispose();
                    }
                    g.Dispose();

                    break;
                default:
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

            if (!checkBox.Checked && IsNotRegType9) {
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
            PreviewBarcode();
        }
        private void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Barcode";
            PreviewBarcode();
        }
        private void シリアルラベル印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductLabelSettingsWindow ls = new();
            ls.ShowDialog(this);
            LoadSettings(labelSettingFilePath, barcodeSettingFilePath);
        }
        private void バーコード印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductBarcodeSettingsWindow ls = new();
            ls.ShowDialog(this);
            LoadSettings(labelSettingFilePath, barcodeSettingFilePath);
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
