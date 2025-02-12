using GenCode128;
using OfficeOpenXml;
using ProductDatabase.Other;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Rendering;
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
        private string _serialFirst = string.Empty;
        private string _serialLast = string.Empty;
        private string _totalSubstrate = string.Empty;
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

                GenerationReportButton.Enabled = false;
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
                            cmd.CommandText = $"""SELECT SubstrateNumber, Stock, SubstrateName FROM "{ProductInfo.StockName}_StockView" WHERE Stock > 0 AND SubstrateModel = @SubstrateModel""";
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

                            cmd.CommandText = $"""SELECT * FROM "{ProductInfo.StockName}_StockView" WHERE Stock > 0 AND SubstrateModel = @SubstrateModel""";
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
            labelSettingFilePath = Path.Combine(Environment.CurrentDirectory, "config", ProductInfo.ProductName, $"SerialConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml");
            if (!File.Exists(labelSettingFilePath)) { throw new DirectoryNotFoundException($"ラベル印刷用設定ファイルがありません。"); }
        }
        private void ConfigureBarcodeSettings() {
            SettingsBarcodePro = new CSettingsBarcodePro();
            barcodeSettingFilePath = Path.Combine(Environment.CurrentDirectory, "config", ProductInfo.ProductName, $"BarcodeConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml");
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
                GenerationReportButton.Enabled = true;
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
            using SQLiteConnection con = new(GetConnectionRegistration());
            con.Open();
            using var transaction = con.BeginTransaction();

            var cmd = con.CreateCommand();
            if (IsSerialGeneration) {
                foreach (var b in _strSerial) {
                    cmd.CommandText =
                        $"""
                        INSERT INTO "{ProductInfo.ProductName}_Serial"
                            (
                            Serial,
                            OrderNumber,
                            ProductNumber,
                            ProductType,
                            ProductModel,
                            RegDate
                            )
                        VALUES
                            (
                            @Serial,
                            @OrderNumber,
                            @ProductNumber,
                            @ProductType,
                            @ProductModel,
                            @RegDate
                            )
                        """;

                    cmd.Parameters.Add("@Serial", DbType.String).Value = b;
                    cmd.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                    cmd.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                    cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                    cmd.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                    cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;

                    cmd.ExecuteNonQuery();
                }
            }

            switch (ProductInfo.RegType) {
                case 0:
                    cmd.CommandText =
                        $"""
                        INSERT INTO "{ProductInfo.ProductName}_Product"
                            (
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
                        INSERT INTO "{ProductInfo.ProductName}_Product"
                            (
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

                    cmd.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                    cmd.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                    cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                    cmd.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                    cmd.Parameters.Add("@Quantity", DbType.String).Value = ProductInfo.Quantity;
                    cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                    cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                    cmd.Parameters.Add("@Revision", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Revision) ? DBNull.Value : ProductInfo.Revision;
                    cmd.Parameters.Add("@RevisionGroup", DbType.String).Value = ProductInfo.RevisionGroup;
                    cmd.Parameters.Add("@SerialFirst", DbType.String).Value = string.IsNullOrWhiteSpace(_serialFirst) ? DBNull.Value : _serialFirst;
                    cmd.Parameters.Add("@SerialLast", DbType.String).Value = string.IsNullOrWhiteSpace(_serialLast) ? DBNull.Value : _serialLast;
                    cmd.Parameters.Add("@SerialLastNumber", DbType.String).Value = IsSerialGeneration ? _serialLastNumber : DBNull.Value;
                    cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                    cmd.ExecuteNonQuery();
                    break;

                case 2:
                case 3:
                case 4:
                    string productRowId;
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
                                    cmd.CommandText = $"""SELECT * FROM "{ProductInfo.StockName}_StockView" WHERE SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber""";
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
                            _totalSubstrate = string.IsNullOrEmpty(_totalSubstrate)
                                ? $"[{_useSubstrate[i]}]{subTotalTemp}"
                                : $"{_totalSubstrate},{Environment.NewLine}[{_useSubstrate[i]}]{subTotalTemp}";
                            subTotalTemp = string.Empty;
                        }
                    }

                    // 製品テーブルに追加
                    cmd = con.CreateCommand();
                    cmd.CommandText =
                        $"""
                            INSERT INTO "{ProductInfo.ProductName}_Product"
                            (
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

                    cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                    cmd.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                    cmd.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                    cmd.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                    cmd.Parameters.Add("@Quantity", DbType.String).Value = ProductInfo.Quantity;
                    cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                    cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                    cmd.Parameters.Add("@Revision", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Revision) ? DBNull.Value : ProductInfo.Revision;
                    cmd.Parameters.Add("@RevisionGroup", DbType.String).Value = ProductInfo.RevisionGroup;
                    cmd.Parameters.Add("@SerialFirst", DbType.String).Value = string.IsNullOrWhiteSpace(_serialFirst) ? DBNull.Value : _serialFirst;
                    cmd.Parameters.Add("@SerialLast", DbType.String).Value = string.IsNullOrWhiteSpace(_serialLast) ? DBNull.Value : _serialLast;
                    cmd.Parameters.Add("@SerialLastNumber", DbType.String).Value = IsSerialGeneration ? _serialLastNumber : DBNull.Value;
                    cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;
                    cmd.Parameters.Add("@UsedSubstrate", DbType.String).Value = string.IsNullOrWhiteSpace(_totalSubstrate) ? DBNull.Value : _totalSubstrate;

                    cmd.ExecuteNonQuery();

                    // 製品ROWIDの取得
                    cmd.CommandText = "SELECT last_insert_rowid()";
                    productRowId = cmd.ExecuteScalar().ToString() ?? string.Empty;

                    // 一時テーブルから基板テーブルにコピー
                    cmd = con.CreateCommand();
                    cmd.CommandText =
                        $"""
                        INSERT INTO "{ProductInfo.StockName}_Substrate" (
                            SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber, Decrease, UsedProductType, UsedProductNumber, UsedOrderNumber, Person, RegDate, Comment, UseID
                        )
                        SELECT
                            tsr.SubstrateName, tsr.SubstrateModel, tsr.SubstrateNumber, tsr.OrderNumber, tsr.Decrease, tsr.UsedProductType, tsr.UsedProductNumber, tsr.UsedOrderNumber, tsr.Person, tsr.RegDate, tsr.Comment, @productRowId
                        FROM TempSubstrateReduction tsr
                        """;
                    cmd.Parameters.Add("@productRowId", DbType.String).Value = int.Parse(productRowId);
                    cmd.ExecuteNonQuery();

                    // 一時テーブルの内容を削除
                    var command = new SQLiteCommand("DELETE FROM TempSubstrateReduction", con, transaction);
                    command.ExecuteNonQuery();
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
            BackupManager.CreateBackup();
            // ログ出力
            Logger.AppendLog($";[製品登録];注文番号[{ProductInfo.OrderNumber}];製造番号[{ProductInfo.ProductNumber}];製品名[{ProductInfo.ProductName}];タイプ[{ProductInfo.ProductType}];型式[{ProductInfo.ProductModel}];数量[{ProductInfo.Quantity}];シリアル先頭[{_serialFirst}];シリアル末尾[{_serialLast}];Revision[{ProductInfo.Revision}];登録日[{ProductInfo.RegDate}];担当者[{ProductInfo.Person}];");
            return true;
        }
        private bool NumberCheck() {
            using SQLiteConnection con = new(GetConnectionRegistration());
            con.Open();

            var productModel = string.Empty;

            if (!string.IsNullOrEmpty(ProductInfo.ProductNumber)) {
                // 製番が新規かチェック
                using (var cmd = con.CreateCommand()) {
                    cmd.CommandText = $"""SELECT * FROM "{ProductInfo.ProductName}_Product" WHERE ProductNumber = @ProductNumber ORDER BY "ID" ASC LIMIT 1""";
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
                    cmd.CommandText = $"""SELECT * FROM "{ProductInfo.ProductName}_Product" WHERE OrderNumber = @OrderNumber ORDER BY "ID" ASC LIMIT 1""";
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
                cmd.CommandText = $"""SELECT Serial FROM "{ProductInfo.ProductName}_Serial" WHERE Serial IN ({string.Join(",", _strSerial.Select((_, i) => $"@Serial{i}"))})""";
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
            _serialFirst = GenerateCode(ProductInfo.SerialFirstNumber);
            _serialLast = GenerateCode(_serialLastNumber);
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
        private void GenerationReport() {
            try {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "ConfigReport.xlsx");
                using FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBookConfig = new ExcelPackage(fileStreamConfig);
                //既存ワークシートを取得（workBookはExcelWorkbookクラスオブジェト）
                var sheet = workBookConfig.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Value?.ToString() == ProductInfo.ProductModel) ?? throw new Exception($"Configに品目番号:[{ProductInfo.ProductModel}]が見つかりません。");
                var searchAddressResultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得
                var directoryPath = workSheetMain.Cells[searchAddressResultRow, 3].Value.ToString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(directoryPath)) { throw new Exception("Configのファイルパスが無効です。"); }
                if (!Directory.Exists(directoryPath)) { throw new FileNotFoundException($"指定されたフォルダが存在しません: {directoryPath}"); }

                var searchName = workSheetMain.Cells[searchAddressResultRow, 4].Value.ToString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(searchName)) { throw new Exception("Configのファイル名が無効です。"); }

                var filePaths = Directory.GetFiles(directoryPath, $"*{searchName}*", SearchOption.TopDirectoryOnly);
                var filePath = filePaths[0];
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var fileExtension = Path.GetExtension(filePath).ToLower(); // 開いたファイルの拡張子取得

                var sheetName = !string.IsNullOrEmpty(workSheetMain.Cells[searchAddressResultRow, 5].Value.ToString())
                    ? workSheetMain.Cells[searchAddressResultRow, 5].Value.ToString()
                    : throw new Exception("シート名がありません。");
                var productNumberRange = workSheetMain.Cells[searchAddressResultRow, 6].Value?.ToString() ?? workSheetMain.Cells[2, 6].Value.ToString();
                var orderNumberRange = workSheetMain.Cells[searchAddressResultRow, 7].Value?.ToString() ?? workSheetMain.Cells[2, 7].Value.ToString();
                var quantityRange = workSheetMain.Cells[searchAddressResultRow, 8].Value?.ToString() ?? workSheetMain.Cells[2, 8].Value.ToString();
                var serialFirstRange = workSheetMain.Cells[searchAddressResultRow, 9].Value?.ToString() ?? workSheetMain.Cells[2, 9].Value.ToString();
                var serialLastRange = workSheetMain.Cells[searchAddressResultRow, 10].Value?.ToString() ?? workSheetMain.Cells[2, 10].Value.ToString();
                var productModelRange = workSheetMain.Cells[searchAddressResultRow, 11].Value?.ToString();
                var saveDirectory = workSheetMain.Cells[searchAddressResultRow, 12].Value?.ToString();

                using FileStream fileStreamReport = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBookReport = new ExcelPackage(fileStreamReport);

                // セルに値を挿入
                var workSheetTemp = sheet[sheetName];
                if (!string.IsNullOrEmpty(ProductInfo.ProductNumber)) {
                    var productNumber = ProductInfo.ProductNumber.Split("-");
                    workSheetTemp.Cells[productNumberRange].Value = productNumber[0];
                }
                if (!string.IsNullOrEmpty(orderNumberRange)) { workSheetTemp.Cells[orderNumberRange].Value = ProductInfo.OrderNumber; }
                if (!string.IsNullOrEmpty(quantityRange)) { workSheetTemp.Cells[quantityRange].Value = ProductInfo.Quantity; }
                if (!string.IsNullOrEmpty(serialFirstRange)) { workSheetTemp.Cells[serialFirstRange].Value = _serialFirst; }
                if (!string.IsNullOrEmpty(serialLastRange)) { workSheetTemp.Cells[serialLastRange].Value = _serialLast; }
                if (!string.IsNullOrEmpty(productModelRange)) { workSheetTemp.Cells[productModelRange].Value = ProductInfo.ProductModel; }

                // ダイアログで保存先を選択
                using SaveFileDialog saveFileDialog = new() {
                    Filter = $"Excel Files (*{fileExtension})|*{fileExtension}|All Files (*.*)|*.*",
                    FileName = $"{fileName} のコピー{ProductInfo.ProductNumber}{fileExtension}",
                    Title = "保存先を選択してください",
                    InitialDirectory = saveDirectory
                };
                if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                    var outputPath = saveFileDialog.FileName;
                    var fileInfo = new FileInfo(outputPath);
                    workBookReport.SaveAs(fileInfo);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // リスト印刷
        private void GenerationList() {
            try {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "ConfigList.xlsx");
                using FileStream fileStream = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBook = new ExcelPackage(fileStream);
                //既存ワークシートを取得（workBookはExcelWorkbookクラスオブジェクト）
                var sheet = workBook.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Value?.ToString() == ProductInfo.ProductModel) ?? throw new Exception($"Configに品目番号:[{ProductInfo.ProductModel}]が見つかりません。");
                var resultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得
                var sheetName = workSheetMain.Cells[resultRow, 2].Value?.ToString();
                var productName = workSheetMain.Cells[resultRow, 3].Value?.ToString();
                var productNameRange = workSheetMain.Cells[resultRow, 4].Value?.ToString();
                var productNumberRange = workSheetMain.Cells[resultRow, 5].Value?.ToString();
                var orderNumberRange = workSheetMain.Cells[resultRow, 6].Value?.ToString();
                var regDateRange = workSheetMain.Cells[resultRow, 7].Value?.ToString();
                var productModel = workSheetMain.Cells[resultRow, 8].Value?.ToString();
                var productModelRange = workSheetMain.Cells[resultRow, 9].Value?.ToString();
                var quantityRange = workSheetMain.Cells[resultRow, 10].Value?.ToString();
                var serialFirstRange = workSheetMain.Cells[resultRow, 11].Value?.ToString();
                var serialLastRange = workSheetMain.Cells[resultRow, 12].Value?.ToString();
                var commentRange = workSheetMain.Cells[resultRow, 13].Value?.ToString();
                var qrCodeRange = workSheetMain.Cells[resultRow, 14].Value?.ToString();

                var workSheetTemp = sheet[sheetName];
                if (string.IsNullOrEmpty(productNameRange)) { workSheetTemp.Cells[productNameRange].Value = productName; }
                if (string.IsNullOrEmpty(productNumberRange)) { workSheetTemp.Cells[productNumberRange].Value = ProductInfo.ProductNumber; }
                if (string.IsNullOrEmpty(orderNumberRange)) { workSheetTemp.Cells[orderNumberRange].Value = ProductInfo.OrderNumber; }
                if (string.IsNullOrEmpty(regDateRange)) { workSheetTemp.Cells[regDateRange].Value = ProductInfo.RegDate; }
                if (string.IsNullOrEmpty(productModelRange)) { workSheetTemp.Cells[productModelRange].Value = ProductInfo.ProductModel; }
                if (string.IsNullOrEmpty(quantityRange)) { workSheetTemp.Cells[quantityRange].Value = ProductInfo.Quantity; }
                if (string.IsNullOrEmpty(serialFirstRange)) { workSheetTemp.Cells[serialFirstRange].Value = _serialFirst; }
                if (string.IsNullOrEmpty(serialLastRange)) { workSheetTemp.Cells[serialLastRange].Value = _serialLast; }
                if (string.IsNullOrEmpty(commentRange)) { workSheetTemp.Cells[commentRange].Value = ProductInfo.Comment; }

                for (var i = 0; i <= _usedSubstrate.Count - 1; i++) {
                    var targetRow = resultRow; // 検索対象の行番号
                    var searchValue = $"{_usedSubstrate[i]}";
                    var foundColumn = 0;

                    var searchAddressResult2 = workSheetMain.Cells
                        .Where(x => x.Start.Row == targetRow) // 指定した行のセルのみを対象にする
                        .First(x => x.Value?.ToString() == searchValue);

                    if (searchAddressResult2 != null) {
                        // セルが見つかった場合の処理
                        foundColumn = searchAddressResult2.Start.Column;
                    }
                    if (foundColumn == 0) {
                        throw new Exception($"{_usedSubstrate[i]}が見つかりません。");
                    }

                    var mainCellValue = workSheetMain.Cells[resultRow, foundColumn + 1].Value.ToString();
                    var tempCellValue = workSheetTemp.Cells[mainCellValue].Value?.ToString();

                    if (mainCellValue != string.Empty) {
                        if (tempCellValue == string.Empty) {
                            workSheetTemp.Cells[mainCellValue].Value = $"{_usedProductNumber[i]}({_usedQuantity[i]})";
                        }
                        else {
                            workSheetTemp.Cells[mainCellValue].Value += $"    {_usedProductNumber[i]}({_usedQuantity[i]})";
                        }
                    }
                }

                // QRコード
                if (!string.IsNullOrEmpty(qrCodeRange)) {
                    BarcodeWriter<PixelData> qr = new() {
                        Format = BarcodeFormat.QR_CODE,
                        Options = new QrCodeEncodingOptions {
                            ErrorCorrection = ErrorCorrectionLevel.L,
                            CharacterSet = "Shift_JIS",
                            Width = 100,
                            Height = 100,
                        },
                        Renderer = new PixelDataRenderer {
                            Foreground = new(Color.Gray.ToArgb()),
                            Background = new(Color.White.ToArgb()),
                        },
                    };

                    var pixelData = qr.Write($"{ProductInfo.OrderNumber};{ProductInfo.ProductNumber};{productModel};{ProductInfo.Quantity};{_serialFirst};{_serialLast}");

                    // PixelData を Bitmap に変換
                    using var bitmap = new Bitmap(pixelData.Width, pixelData.Height, PixelFormat.Format32bppArgb);
                    var bmpData = bitmap.LockBits(new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                    System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bmpData.Scan0, pixelData.Pixels.Length);
                    bitmap.UnlockBits(bmpData);
                    using MemoryStream stream = new();
                    bitmap.Save(stream, ImageFormat.Bmp);

                    var image = workSheetTemp.Drawings.AddPicture("QR", bitmap);
                    image.SetPosition(workSheetTemp.Cells[qrCodeRange].Start.Row - 1, 0, workSheetTemp.Cells[qrCodeRange].Start.Column - 1, 0);
                }

                //引数に保存先パスを指定
                var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "temporarilyList.xlsx");
                var fileInfo = new FileInfo(temporarilyPath);
                workBook.SaveAs(fileInfo);

                // 印刷
                Microsoft.Office.Interop.Excel.Application xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                var xlBooks = xlApp.Workbooks;
                var xlBook = xlBooks.Open(temporarilyPath, ReadOnly: true);

                // ワークシート選択
                var xlSheets = xlBook.Sheets;
                Microsoft.Office.Interop.Excel.Worksheet xlSheet = xlSheets[sheetName];

                // ワークシート表示
                xlSheet.Activate();

                // ワークブックを閉じてExcelを終了
                //xlBook.Close(false);
                //xlApp.Quit();

                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // チェックシート印刷
        private void GenerationCheckSheet() {
            try {

                var dialog = new InputDialog1();
                var result = dialog.ShowDialog();

                var temperature = string.Empty;
                var humidity = string.Empty;

                if (result == DialogResult.OK) {
                    temperature = dialog.Temperature;
                    humidity = dialog.Humidity;
                }
                else {
                    return;
                }

                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "ConfigCheckSheet.xlsx");
                using FileStream fileStream = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBook = new ExcelPackage(fileStream);
                //既存ワークシートを取得（workBookはExcelWorkbookクラスオブジェクト）
                var sheet = workBook.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Value?.ToString() == ProductInfo.ProductModel) ?? throw new Exception($"Configに品目番号:[{ProductInfo.ProductModel}]が見つかりません。");
                var resultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得
                var orderNumberRange = workSheetMain.Cells[resultRow, 7].Value?.ToString();
                var quantityRange = workSheetMain.Cells[resultRow, 8].Value?.ToString();
                var serialFirstRange = workSheetMain.Cells[resultRow, 9].Value?.ToString();
                var serialLastRange = workSheetMain.Cells[resultRow, 10].Value?.ToString();
                var regDateYearRange = workSheetMain.Cells[resultRow, 11].Value?.ToString();
                var regDateMonthRange = workSheetMain.Cells[resultRow, 12].Value?.ToString();
                var regDateDayRange = workSheetMain.Cells[resultRow, 13].Value?.ToString();
                var regTemperatureRange = workSheetMain.Cells[resultRow, 14].Value?.ToString();
                var regHumidityRange = workSheetMain.Cells[resultRow, 15].Value?.ToString();

                const int StartColumn = 16;
                var sheetNames = Enumerable.Range(StartColumn, 20) // 無限の範囲
                    .Select(column => workSheetMain.Cells[resultRow, column].Value?.ToString())
                    .TakeWhile(sheetName => !string.IsNullOrWhiteSpace(sheetName)) // 空白でない間
                    .ToList();

                if (sheetNames.Count == 0) { throw new Exception("対象シートがありません。"); }

                var date = DateTime.Parse(ProductInfo.RegDate);

                foreach (var sheetName in sheetNames) {
                    var workSheetTemp = sheet[sheetName] ?? throw new Exception($"シート[{sheetName}]が見つかりません。");
                    if (!string.IsNullOrEmpty(orderNumberRange)) { workSheetTemp.Cells[orderNumberRange].Value = ProductInfo.OrderNumber; }
                    if (!string.IsNullOrEmpty(quantityRange)) { workSheetTemp.Cells[quantityRange].Value = ProductInfo.Quantity; }
                    if (!string.IsNullOrEmpty(serialFirstRange)) { workSheetTemp.Cells[serialFirstRange].Value = _serialFirst; }
                    if (!string.IsNullOrEmpty(serialLastRange)) { workSheetTemp.Cells[serialLastRange].Value = _serialLast; }
                    if (!string.IsNullOrEmpty(regDateYearRange)) { workSheetTemp.Cells[regDateYearRange].Value = date.Year; }
                    if (!string.IsNullOrEmpty(regDateMonthRange)) { workSheetTemp.Cells[regDateMonthRange].Value = date.Month; }
                    if (!string.IsNullOrEmpty(regDateDayRange)) { workSheetTemp.Cells[regDateDayRange].Value = date.Day; }
                    if (!string.IsNullOrEmpty(regTemperatureRange)) { workSheetTemp.Cells[regTemperatureRange].Value = temperature; }
                    if (!string.IsNullOrEmpty(regHumidityRange)) { workSheetTemp.Cells[regHumidityRange].Value = humidity; }
                }

                // 不要なシートを非表示にする
                var allSheetName = sheet
                    .Select(sh => sh.Name.ToString())
                    .ToList();

                var hiddenSheetNames = allSheetName.Except(sheetNames.Where(name => name != null).Cast<string>()).ToList();
                foreach (var sheetName in hiddenSheetNames) {
                    sheet[sheetName].Hidden = eWorkSheetHidden.VeryHidden;
                }
                sheet["Sheet1"].Hidden = eWorkSheetHidden.VeryHidden;

                //引数に保存先パスを指定
                var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "Excel", "temporarilyCheckSheet.xlsx");
                var fileInfo = new FileInfo(temporarilyPath);
                workBook.SaveAs(fileInfo);

                // 印刷
                Microsoft.Office.Interop.Excel.Application xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                var xlBooks = xlApp.Workbooks;
                var xlBook = xlBooks.Open(temporarilyPath, ReadOnly: true);

                //// ワークシート選択
                //var xlSheets = xlBook.Sheets;
                //Microsoft.Office.Interop.Excel.Worksheet xlSheet = xlSheets[0];

                //// ワークシート表示
                //xlSheet.Activate();

                //// ワークブックを閉じてExcelを終了
                //xlBook.Close(false);
                //xlApp.Quit();

                //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                //_ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProductRegistration2Window_Load(object sender, EventArgs e) { LoadEvents(); }
        private void GenerationReportButton_Click(object sender, EventArgs e) { GenerationReport(); }
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
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) { GenerationList(); }
        private void CheckSheetPrintButton_Click(object sender, EventArgs e) { GenerationCheckSheet(); }
        private void ProductRegistration2PrintPreviewDialog_Load(object sender, EventArgs e) {
            var tool = (ToolStrip)ProductRegistration2PrintPreviewDialog.Controls[1];
            tool.Items[0].Visible = false;
        }
    }
}
