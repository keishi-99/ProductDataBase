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

        private int _pageCnt = 1;

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
        private readonly List<string> _usedSubstrate = [];
        private readonly List<string> _usedProductNumber = [];
        private readonly List<int> _usedQuantity = [];

        // プロパティ設定
        private bool IsSerialGeneration => ProductInfo.RegType is 1 or 2 or 3;
        private bool IsLabelPrint => ProductInfo.PrintType is 1 or 3 or 4 or 5 or 6 or 7 or 9;
        private bool IsBarcodePrint => ProductInfo.PrintType is 2 or 3;
        private bool RequiresClosing => ProductInfo.PrintType is 0 or 1 or 2 or 3 or 4 or 9;
        private bool IsListPrint => ProductInfo.PrintType is 5 or 6;
        private bool IsCheckSheetPrint => ProductInfo.PrintType is 6 or 7;
        private bool IsUnderlinePrint => ProductInfo.PrintType is 4;
        private bool IsLast4Digits => ProductInfo.PrintType is 9;

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
                                ORDER BY MIN(_rowid_);
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
                                ORDER BY MIN(_rowid_);
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
            labelSettingFilePath = Path.Combine(Environment.CurrentDirectory, "config","Product", ProductInfo.CategoryName, ProductInfo.ProductName, $"SerialConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml");
            if (!File.Exists(labelSettingFilePath)) { throw new DirectoryNotFoundException($"ラベル印刷用設定ファイルがありません。"); }
        }
        private void ConfigureBarcodeSettings() {
            SettingsBarcodePro = new CSettingsBarcodePro();
            barcodeSettingFilePath = Path.Combine(Environment.CurrentDirectory, "config", "Product", ProductInfo.CategoryName, ProductInfo.ProductName, $"BarcodeConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml");
            if (!File.Exists(barcodeSettingFilePath)) { throw new DirectoryNotFoundException($"バーコード印刷用設定ファイルがありません。"); }
        }
        private void SetMenuOptions() {
            シリアルラベル印刷ToolStripMenuItem.Enabled = false;
            シリアルラベル印刷プレビューToolStripMenuItem.Enabled = IsLabelPrint;
            シリアルラベル印刷設定ToolStripMenuItem.Enabled = IsLabelPrint;
            バーコード印刷ToolStripMenuItem.Enabled = false;
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

                if (!Registration()) {
                    throw new Exception("登録失敗しました。");
                }

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
        private bool Registration() {
            string productRowId;
            using SQLiteConnection con = new(GetConnectionRegistration());
            con.Open();
            using var transaction = con.BeginTransaction();

            var cmd = con.CreateCommand();

            switch (ProductInfo.RegType) {
                case 0:
                    cmd.CommandText =
                        $"""
                        INSERT INTO "{ProductInfo.CategoryName}_Product"
                            (
                            ProductName,
                            OrderNumber,
                            ProductNumber,
                            ProductType,
                            ProductModel,
                            Quantity,
                            Person,
                            RegDate,
                            Revision,
                            RevisionGroup,
                            Comment
                            )
                        VALUES
                            (
                            @ProductName,
                            @OrderNumber,
                            @ProductNumber,
                            @ProductType,
                            @ProductModel,
                            @Quantity,
                            @Person,
                            @RegDate,
                            @Revision,
                            @RevisionGroup,
                            @Comment
                            )
                        """;

                    cmd.Parameters.Add("@ProductName", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductName) ? DBNull.Value : ProductInfo.ProductName;
                    cmd.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                    cmd.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                    cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                    cmd.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                    cmd.Parameters.Add("@Quantity", DbType.String).Value = ProductInfo.Quantity;
                    cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                    cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                    cmd.Parameters.Add("@Revision", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Revision) ? DBNull.Value : ProductInfo.Revision;
                    cmd.Parameters.Add("@RevisionGroup", DbType.String).Value = ProductInfo.RevisionGroup;
                    cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                    cmd.ExecuteNonQuery();
                    break;

                case 1:
                    cmd.CommandText =
                        $"""
                        INSERT INTO "{ProductInfo.CategoryName}_Product"
                            (
                            ProductName,
                            OrderNumber,
                            ProductNumber,
                            ProductType,
                            ProductModel,
                            Quantity,
                            Person,
                            RegDate,
                            Revision,
                            RevisionGroup,
                            SerialFirst,
                            SerialLast,
                            SerialLastNumber,
                            Comment
                            )
                        VALUES
                            (
                            @ProductName,
                            @OrderNumber,
                            @ProductNumber,
                            @ProductType,
                            @ProductModel,
                            @Quantity,
                            @Person,
                            @RegDate,
                            @Revision,
                            @RevisionGroup,
                            @SerialFirst,
                            @SerialLast,
                            @SerialLastNumber,
                            @Comment
                            )
                        """;

                    cmd.Parameters.Add("@ProductName", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductName) ? DBNull.Value : ProductInfo.ProductName;
                    cmd.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                    cmd.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                    cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                    cmd.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                    cmd.Parameters.Add("@Quantity", DbType.String).Value = ProductInfo.Quantity;
                    cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                    cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                    cmd.Parameters.Add("@Revision", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Revision) ? DBNull.Value : ProductInfo.Revision;
                    cmd.Parameters.Add("@RevisionGroup", DbType.String).Value = ProductInfo.RevisionGroup;
                    cmd.Parameters.Add("@SerialFirst", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialFirst) ? DBNull.Value : ProductInfo.SerialFirst;
                    cmd.Parameters.Add("@SerialLast", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialLast) ? DBNull.Value : ProductInfo.SerialLast;
                    cmd.Parameters.Add("@SerialLastNumber", DbType.String).Value = IsSerialGeneration ? _serialLastNumber : DBNull.Value;
                    cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "SELECT last_insert_rowid()";
                    productRowId = cmd.ExecuteScalar().ToString() ?? string.Empty;

                    if (IsSerialGeneration) {
                        foreach (var b in _strSerial) {
                            cmd.CommandText =
                                $"""
                        INSERT INTO "{ProductInfo.CategoryName}_Serial"
                            (
                            Serial,
                            UsedID
                            )
                        VALUES
                            (
                            @Serial,
                            @productRowId
                            )
                        """;

                            cmd.Parameters.Add("@Serial", DbType.String).Value = b;
                            cmd.Parameters.Add("@productRowId", DbType.String).Value = int.Parse(productRowId);

                            cmd.ExecuteNonQuery();
                        }
                    }
                    break;

                case 2:
                case 3:
                case 4:
                    if (_useSubstrate == null) { throw new Exception("ArrUseSubstrateがnullです。"); }
                    for (var i = 0; i <= _useSubstrate.Length; i++) {

                        var objCbx = Controls[_checkBoxNames[i]] as System.Windows.Forms.CheckBox ?? throw new Exception("objCbxがnullです。");

                        if (objCbx.Checked) {
                            var objDgv = Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objCbxがnullです。");
                            var dgvRowCnt = objDgv.Rows.Count;
                            var subTotalTemp = string.Empty;

                            for (var j = 0; j <= dgvRowCnt - 1; j++) {
                                var boolCbx = Convert.ToBoolean(objDgv.Rows[j].Cells[3].Value);
                                if (boolCbx) {
                                    var substrateName = string.Empty;
                                    var substrateModel = string.Empty;
                                    var orderNum = string.Empty;
                                    var substrateNum = objDgv.Rows[j].Cells[0].Value.ToString() ?? string.Empty;
                                    var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value);
                                    var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);

                                    cmd = con.CreateCommand();
                                    //cmd.CommandText = $"""SELECT * FROM "{ProductInfo.StockName}_StockView" WHERE SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber""";
                                    cmd.CommandText = $"""
                                        SELECT
                                            SubstrateName,
                                            SubstrateModel,
                                            SubstrateNumber,
                                            OrderNumber,
                                            SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                                        FROM {ProductInfo.CategoryName}_Substrate
                                        WHERE StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                                        GROUP BY SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                                        ORDER BY MIN(_rowid_);
                                        """;
                                    cmd.Parameters.Add("@StockName", DbType.String).Value = ProductInfo.StockName;
                                    cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                                    cmd.Parameters.Add("@SubstrateNumber", DbType.String).Value = substrateNum;

                                    using var dr = cmd.ExecuteReader();
                                    while (dr.Read()) {
                                        substrateName = $"{dr["SubstrateName"]}";
                                        substrateModel = $"{dr["SubstrateModel"]}";
                                        orderNum = $"{dr["OrderNumber"]}";
                                    }

                                    if (useValue != 0) {
                                        subTotalTemp = string.IsNullOrEmpty(subTotalTemp) ? $"{substrateNum}({useValue})" : $"{subTotalTemp},{substrateNum}({useValue})";
                                    }

                                    // 一時テーブルに登録
                                    cmd = con.CreateCommand();
                                    cmd.CommandText =
                                        $"""
                                        INSERT INTO "TempSubstrateReduction"
                                            (
                                            StockName,
                                            SubstrateName,
                                            SubstrateModel,
                                            SubstrateNumber,
                                            OrderNumber,
                                            Decrease,
                                            UsedProductType,
                                            UsedProductNumber,
                                            UsedOrderNumber,
                                            Person,
                                            RegDate,
                                            Comment
                                            )
                                        VALUES
                                            (
                                            @StockName,
                                            @SubstrateName,
                                            @SubstrateModel,
                                            @SubstrateNumber,
                                            @OrderNumber,
                                            @Decrease,
                                            @UsedProductType,
                                            @UsedProductNumber,
                                            @UsedOrderNumber,
                                            @Person,
                                            @RegDate,
                                            @Comment
                                            )
                                        """;

                                    cmd.Parameters.Add("@StockName", DbType.String).Value = ProductInfo.StockName;
                                    cmd.Parameters.Add("@SubstrateName", DbType.String).Value = string.IsNullOrWhiteSpace(substrateName) ? DBNull.Value : substrateName;
                                    cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = string.IsNullOrWhiteSpace(substrateModel) ? DBNull.Value : substrateModel;
                                    cmd.Parameters.Add("@SubstrateNumber", DbType.String).Value = objDgv.Rows[j].Cells[0].Value;
                                    cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(orderNum) ? DBNull.Value : orderNum;
                                    cmd.Parameters.Add("@Decrease", DbType.String).Value = 0 - useValue;
                                    cmd.Parameters.Add("@UsedProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                                    cmd.Parameters.Add("@UsedProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                                    cmd.Parameters.Add("@UsedOrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                                    cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                                    cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                                    cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                                    cmd.ExecuteNonQuery();

                                    if (IsLabelPrint) {
                                        _usedSubstrate.Add(_useSubstrate[i]);
                                        if (substrateNum != null) { _usedProductNumber.Add(substrateNum); }
                                        _usedQuantity.Add(useValue);
                                    }
                                }
                            }
                            ProductInfo.UsedSubstrate = string.IsNullOrEmpty(ProductInfo.UsedSubstrate)
                                ? $"[{_useSubstrate[i]}]{subTotalTemp}"
                                : $"{ProductInfo.UsedSubstrate},{Environment.NewLine}[{_useSubstrate[i]}]{subTotalTemp}";
                            subTotalTemp = string.Empty;
                        }
                    }

                    // 製品テーブルに追加
                    cmd = con.CreateCommand();
                    cmd.CommandText =
                        $"""
                            INSERT INTO "{ProductInfo.CategoryName}_Product"
                            (
                            ProductName,
                            OrderNumber,
                            ProductNumber,
                            ProductType,
                            ProductModel,
                            Quantity,
                            Person,
                            RegDate,
                            Revision,
                            RevisionGroup,
                            SerialFirst,
                            SerialLast,
                            SerialLastNumber,
                            Comment,
                            UsedSubstrate
                            )
                        VALUES
                            (
                            @ProductName,
                            @OrderNumber,
                            @ProductNumber,
                            @ProductType,
                            @ProductModel,
                            @Quantity,
                            @Person,
                            @RegDate,
                            @Revision,
                            @RevisionGroup,
                            @SerialFirst,
                            @SerialLast,
                            @SerialLastNumber,
                            @Comment,
                            @UsedSubstrate
                            )
                        """;

                    cmd.Parameters.Add("@ProductName", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductName) ? DBNull.Value : ProductInfo.ProductName;
                    cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                    cmd.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                    cmd.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                    cmd.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                    cmd.Parameters.Add("@Quantity", DbType.String).Value = ProductInfo.Quantity;
                    cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                    cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                    cmd.Parameters.Add("@Revision", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Revision) ? DBNull.Value : ProductInfo.Revision;
                    cmd.Parameters.Add("@RevisionGroup", DbType.String).Value = ProductInfo.RevisionGroup;
                    cmd.Parameters.Add("@SerialFirst", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialFirst) ? DBNull.Value : ProductInfo.SerialFirst;
                    cmd.Parameters.Add("@SerialLast", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialLast) ? DBNull.Value : ProductInfo.SerialLast;
                    cmd.Parameters.Add("@SerialLastNumber", DbType.String).Value = IsSerialGeneration ? _serialLastNumber : DBNull.Value;
                    cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;
                    cmd.Parameters.Add("@UsedSubstrate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.UsedSubstrate) ? DBNull.Value : ProductInfo.UsedSubstrate;

                    cmd.ExecuteNonQuery();

                    // 製品ROWIDの取得
                    cmd.CommandText = "SELECT last_insert_rowid()";
                    productRowId = cmd.ExecuteScalar().ToString() ?? string.Empty;

                    // 一時テーブルから基板テーブルにコピー
                    cmd = con.CreateCommand();
                    cmd.CommandText =
                        $"""
                        INSERT INTO "{ProductInfo.CategoryName}_Substrate" (
                            StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber, Decrease, UsedProductType, UsedProductNumber, UsedOrderNumber, Person, RegDate, Comment, UseID
                        )
                        SELECT
                            tsr.StockName, tsr.SubstrateName, tsr.SubstrateModel, tsr.SubstrateNumber, tsr.OrderNumber, tsr.Decrease, tsr.UsedProductType, tsr.UsedProductNumber, tsr.UsedOrderNumber, tsr.Person, tsr.RegDate, tsr.Comment, @productRowId
                        FROM TempSubstrateReduction tsr
                        """;
                    cmd.Parameters.Add("@productRowId", DbType.String).Value = int.Parse(productRowId);
                    cmd.ExecuteNonQuery();

                    // 一時テーブルの内容を削除
                    var command = new SQLiteCommand("DELETE FROM TempSubstrateReduction", con, transaction);
                    command.ExecuteNonQuery();

                    if (IsSerialGeneration) {
                        foreach (var b in _strSerial) {
                            cmd.CommandText =
                                $"""
                        INSERT INTO "{ProductInfo.CategoryName}_Serial"
                            (
                            Serial,
                            UsedID
                            )
                        VALUES
                            (
                            @Serial,
                            @productRowId
                            )
                        """;

                            cmd.Parameters.Add("@Serial", DbType.String).Value = b;
                            cmd.Parameters.Add("@productRowId", DbType.String).Value = int.Parse(productRowId);

                            cmd.ExecuteNonQuery();
                        }
                    }
                    break;
            }
            try {
                // ラベル印刷処理
                HandleLabelPrinting();

                // バーコード印刷処理
                HandleBarcodePrinting();

                // 印刷が成功した場合、トランザクションをコミット
                transaction.Commit();
            } catch (Exception) {
                // 印刷に失敗した場合、トランザクションをロールバック
                transaction.Rollback();
                throw; // 例外を再度スローして上位で処理できるようにする
            }
            // バックアップ作成
            CommonUtils.BackupManager.CreateBackup();
            // ログ出力
            CommonUtils.Logger.AppendLog($"[製品登録];注文番号[{ProductInfo.OrderNumber}];製造番号[{ProductInfo.ProductNumber}];製品名[{ProductInfo.ProductName}];タイプ[{ProductInfo.ProductType}];型式[{ProductInfo.ProductModel}];数量[{ProductInfo.Quantity}];シリアル先頭[{ProductInfo.SerialFirst}];シリアル末尾[{ProductInfo.SerialLast}];Revision[{ProductInfo.Revision}];登録日[{ProductInfo.RegDate}];担当者[{ProductInfo.Person}];");
            return true;
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
                                        if (objDgv.Rows[j].Cells[2].Value == null) { throw new Exception("使用数が入力されていません。"); }
                                        var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value.ToString());

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
                    INNER JOIN
                        "{ProductInfo.CategoryName}_Product" AS p
                    ON
                        s.UsedID = p.ID
                    WHERE
                        p.ProductName = @ProductName
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

                if (!PrintBarcode(1)) {
                    throw new OperationCanceledException("キャンセルしました。");
                }
            }
        }
        private void HandleBarcodePrinting() {
            if (IsBarcodePrint) {
                MessageBox.Show("バーコードラベルを印刷します。");
                _serialType = "Barcode";

                if (!PrintBarcode(1)) {
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

        // 印刷処理
        private bool PrintBarcode(int printFlg) {
            // PrintDocumentオブジェクトの作成
            using System.Drawing.Printing.PrintDocument pd = new();

            // PrintPageイベントハンドラの追加
            pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

            _labelProNumLabelsToPrint = ProductInfo.Quantity;
            _pageCnt = 1;

            switch (printFlg) {
                case 1:
                    ProductRegistration2PrintDialog.Document = pd;
                    var r = ProductRegistration2PrintDialog.ShowDialog();

                    if (r == DialogResult.OK) {
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
                    break;
                case 2:
                    // 最大で表示
                    ProductRegistration2PrintPreviewDialog.Shown += (sender, e) => {
                        if (sender is Form form) {
                            form.WindowState = FormWindowState.Maximized;
                        }
                    };
                    ProductRegistration2PrintPreviewDialog.PrintPreviewControl.Zoom = 3;
                    ProductRegistration2PrintPreviewDialog.Document = pd;
                    ProductRegistration2PrintPreviewDialog.ShowDialog();
                    break;
            }
            return true;
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            System.Drawing.Point headerPos = new(0, 0);
            var headerString = string.Empty;
            var headerFooterFont = new Font("ＭＳ Ｐ明朝", 5.25F);
            Point offset;
            var serialCodePrintCopies = 0;

            var maxX = 0;
            var maxY = 0;
            decimal sizeX = 0;
            decimal sizeY = 0;
            decimal offsetX = 0;
            decimal offsetY = 0;
            decimal intervalX = 0;
            decimal intervalY = 0;
            var startLine = 0;
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;

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

                const decimal MM_PER_HUNDREDTH_INCH = 0.254M;

                var pd = (PrintDocument)sender;
                var bPrintMode = pd.PrintController.IsPreview;

                if (!bPrintMode) {
                    offsetX -= (decimal)e.PageSettings.HardMarginX * MM_PER_HUNDREDTH_INCH;
                    offsetY -= (decimal)e.PageSettings.HardMarginY * MM_PER_HUNDREDTH_INCH;
                    offset = _pageCnt == 1
                        ? new System.Drawing.Point((int)((decimal)e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (startLine * (intervalY + sizeY))))
                        : new System.Drawing.Point((int)((decimal)e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (0 * (intervalY + sizeY))));
                }
                else {
                    offset = _pageCnt == 1
                        ? new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (startLine * (intervalY + sizeY))))
                        : new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (0 * (intervalY + sizeY))));
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;

                headerPos.Offset(offset);
                e.Graphics.DrawString(headerString, headerFooterFont, Brushes.Black, headerPos);

                if (_pageCnt == 1) {
                    _remainingCount = serialCodePrintCopies;
                    _labelProNSerial = ProductInfo.SerialFirstNumber;
                }
                if (_pageCnt >= 2) { startLine = 0; }

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
                                _pageCnt = 1;
                                _labelProNumLabelsToPrint = 0;
                                return;
                            }
                            _remainingCount = serialCodePrintCopies;
                        }
                    }
                }

                if (_labelProNumLabelsToPrint > 0) {
                    _pageCnt++;
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
            decimal sizeX;
            decimal sizeY;
            decimal fontSize;
            float stringPosX;
            float stringPosY;
            Font fnt;
            var displayResolution = 96.0M;
            var displayMagnitude = 3;

            // サイズとフォント情報の設定
            void SetLabelProperties(decimal labelWidth, decimal labelHeight, decimal posY, decimal fontPointSize, string fontName, bool underlined) {
                sizeX = labelWidth / 25.4M * resolution * magnitude;
                sizeY = labelHeight / 25.4M * resolution * magnitude;
                stringPosY = (float)(posY / 25.4M * resolution * magnitude);
                fontSize = fontPointSize / 72.0M * resolution * magnitude;
                var style = underlined ? FontStyle.Underline : FontStyle.Regular;
                fnt = new Font(fontName, (float)fontSize, style);
            }

            switch (_serialType) {
                case "Label":
                    if (SettingsLabelPro == null) { throw new Exception("SettingsLabelProがnull"); }
                    SetLabelProperties(SettingsLabelPro.LabelProPageSettings.SizeX,
                                        SettingsLabelPro.LabelProPageSettings.SizeY,
                                        SettingsLabelPro.LabelProLabelSettings.StringPosY,
                                        (decimal)SettingsLabelPro.LabelProLabelSettings.Font.SizeInPoints,
                                        SettingsLabelPro.LabelProLabelSettings.Font.Name,
                                        fontUnderline);

                    labelImage = new((int)sizeX, (int)sizeY);
                    g = Graphics.FromImage(labelImage);
                    //stringSize = g.MeasureString(text, fnt);
                    stringSize = TextRenderer.MeasureText(text, fnt);

                    stringPosX = SettingsLabelPro.LabelProLabelSettings.AlignStringCenter
                        ? (float)((labelImage.Width / 2f) - (stringSize.Width / 2f))
                        : (float)(SettingsLabelPro.LabelProLabelSettings.StringPosX / 25.4M * resolution * magnitude);

                    g.DrawString(text, fnt, Brushes.Black, stringPosX, stringPosY);

                    g.Dispose();
                    break;
                case "Barcode":
                    if (SettingsBarcodePro == null) { throw new Exception("SettingsBarcodeProがnull"); }
                    SetLabelProperties(SettingsBarcodePro.BarcodeProPageSettings.SizeX,
                                        SettingsBarcodePro.BarcodeProPageSettings.SizeY,
                                        SettingsBarcodePro.BarcodeProLabelSettings.StringPosY,
                                        (decimal)SettingsBarcodePro.BarcodeProLabelSettings.Font.SizeInPoints,
                                        SettingsBarcodePro.BarcodeProLabelSettings.Font.Name,
                                        false);

                    labelImage = new((int)sizeX, (int)sizeY);
                    g = Graphics.FromImage(labelImage);
                    //stringSize = g.MeasureString(text, fnt);
                    stringSize = TextRenderer.MeasureText(text, fnt);

                    stringPosX = SettingsBarcodePro.BarcodeProLabelSettings.AlignStringCenter
                        ? (float)((labelImage.Width / 2f) - (stringSize.Width / 2f))
                        : (float)(SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4M * resolution * magnitude);

                    g.DrawString(text, fnt, Brushes.Black, stringPosX, stringPosY);

                    var barWeight = resolution == displayResolution ? 1 : (int)(1 * resolution / displayResolution / displayMagnitude);

                    using (var img = Code128Rendering.MakeBarcodeImage(text, barWeight, true)) {
                        var imageWidth = img.Width * SettingsBarcodePro.BarcodeProLabelSettings.BarcodeMagnitude;

                        if (imageWidth > labelImage.Width) { MessageBox.Show($"バーコードの幅がラベル幅を超えています{imageWidth}>{labelImage.Width}"); }

                        var barCodePosX = SettingsBarcodePro.BarcodeProLabelSettings.AlignBarcodeCenter
                            ? (float)((labelImage.Width / 2) - (imageWidth / 2))
                            : (float)(SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4M * resolution * magnitude);
                        var barCodePosY = (float)(SettingsBarcodePro.BarcodeProLabelSettings.BarcodePosY / 25.4M * resolution * magnitude);

                        var barcodeHeight = SettingsBarcodePro.BarcodeProLabelSettings.BarcodeHeight / 25.4M * resolution * magnitude;

                        g.DrawImage(img, barCodePosX, barCodePosY, (float)imageWidth, (float)barcodeHeight);

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
            checkBox.ForeColor = checkBox.Checked ? Color.Black : Color.Red;

            if (!checkBox.Checked) {
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
        private void シリアルラベル印刷ToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Label";
            PrintBarcode(1);
        }
        private void シリアルラベル印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Label";
            PrintBarcode(2);
        }
        private void バーコード印刷ToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Barcode";
            PrintBarcode(1);
        }
        private void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Barcode";
            PrintBarcode(2);
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
