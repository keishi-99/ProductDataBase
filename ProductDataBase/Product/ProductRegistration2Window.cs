using ClosedXML.Excel;
using GenCode128;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Rendering;
using static ProductDatabase.MainWindow;
using Color = System.Drawing.Color;
using Control = System.Windows.Forms.Control;
using Excel = Microsoft.Office.Interop.Excel;
using Font = System.Drawing.Font;

namespace ProductDatabase {
    public partial class ProductRegistration2Window : Form {

        public CSettingsLabelPro SettingsLabelPro { get; set; } = new CSettingsLabelPro();
        public string labelSettingFilePath = string.Empty;

        public CSettingsBarcodePro SettingsBarcodePro { get; set; } = new CSettingsBarcodePro();
        public string barcodeSettingFilePath = string.Empty;

        public ProductInfomation ProductInfo { get; set; } = new ProductInfomation();

        private string[] _useSubstrate = [];

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

        public ProductRegistration2Window() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);

                RegisterButton.Enabled = true;
                _useSubstrate = ProductInfo.UseSubstrate.Split(",");

                _labelProNSerial = ProductInfo.SerialFirstNumber;
                _serialLastNumber = ProductInfo.SerialFirstNumber + ProductInfo.Quantity - 1;

                var quantityFlg = false;
                var strQuantity = string.Empty;
                switch (ProductInfo.RegType) {
                    case 2:
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
                                        objDgv.Columns[0].Width = 140;
                                        objDgv.Columns[1].Width = 35;
                                        objDgv.Columns[2].Width = 35;
                                        objDgv.Columns[3].Width = 35;
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
                            cmd.CommandText = $"""SELECT SubstrateNumber, Stock, SubstrateName FROM "{ProductInfo.StockName}_Stock" WHERE Stock > 0 AND SubstrateModel = @SubstrateModel ORDER BY _rowid_ ASC""";
                            cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                            using var dr = cmd.ExecuteReader();
                            while (dr.Read()) {
                                var strSubstrateNumber = $"{dr["SubstrateNumber"]}";
                                var intStock = Convert.ToInt32(dr["Stock"]);
                                objDgv?.Rows.Add(strSubstrateNumber, intStock);

                                if (intQuantity >= intStock) {
                                    intQuantity -= intStock;
                                    if (objDgv != null) {
                                        objDgv.Rows[^1].Cells[2].Value = intStock;
                                        objDgv.Rows[^1].Cells[3].Value = true;
                                    }
                                }
                                else {
                                    if (intQuantity == 0) {
                                        if (objDgv != null) {
                                            objDgv.Rows[^1].Cells[2].Value = 0;
                                        }
                                    }
                                    else {
                                        if (objDgv != null) {
                                            objDgv.Rows[^1].Cells[2].Value = intQuantity;
                                            objDgv.Rows[^1].Cells[3].Value = true;
                                            intQuantity = 0;
                                        }
                                    }
                                }

                                if (intQuantity > 0) {
                                    quantityFlg = false;
                                    var substrateName = $"{dr["SubstrateName"]}";
                                    strQuantity += $"[{substrateName}]{Environment.NewLine}";
                                }

                                if (intQuantity == 0) { quantityFlg = true; }
                            }
                        }

                        if (quantityFlg == false) {
                            Activate();
                            MessageBox.Show($"在庫が足りません。{Environment.NewLine}{strQuantity}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                                        objDgv.Columns[0].Width = 140;
                                        objDgv.Columns[1].Width = 35;
                                        objDgv.Columns[2].Width = 35;
                                        objDgv.Columns[3].Width = 35;
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

                            cmd.CommandText = $"""SELECT * FROM "{ProductInfo.StockName}_Stock" WHERE Stock > 0 AND SubstrateModel = @SubstrateModel ORDER BY _rowid_ ASC""";
                            cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                            using var dr = cmd.ExecuteReader();
                            while (dr.Read()) {
                                var strSubstrateName = string.Empty;
                                strSubstrateName = $"{dr["SubstrateName"]}";

                                var strSubstrateNumber = $"{dr["SubstrateNumber"]}";
                                var intStock = Convert.ToInt32(dr["Stock"]);
                                objDgv?.Rows.Add(strSubstrateNumber, intStock);

                                var j = 0;
                                var strOrderNumber = $"{dr["OrderNumber"]}";
                                if (strOrderNumber == ProductInfo.OrderNumber) {
                                    if (objDgv != null) {
                                        var intQuantity = ProductInfo.Quantity;
                                        var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value);
                                        var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);
                                        objDgv.Rows[j].Cells[2].Value = intQuantity;
                                        objDgv.Rows[j].Cells[3].Value = true;
                                        // 必要数量分割り当てられたかチェック
                                        if (intQuantity > stockValue) {
                                            quantityFlg = false;
                                            strQuantity = $"{strQuantity}[{strSubstrateName}]{Environment.NewLine}";
                                        }
                                        if (stockValue >= useValue) { quantityFlg = true; }
                                    }
                                }
                                j++;
                            }
                        }

                        if (quantityFlg == false) {
                            Activate();
                            MessageBox.Show($"在庫が足りません。{Environment.NewLine}{strQuantity}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        break;
                    default:
                        for (var i = 0; i <= 9; i++) {
                            if (Controls[_checkBoxNames[i]] is CheckBox objCbx) {
                                objCbx.Visible = false;
                            }

                            if (Controls[_dataGridViewNames[i]] is DataGridView objDgv) {
                                objDgv.Visible = false;
                            }
                        }
                        break;
                }

                switch (ProductInfo.PrintType) {
                    case 1:
                    case 2:
                    case 3:
                    case 4:
                    case 8:
                    case 9:
                        SubstrateListPrintButton.Visible = false;
                        CheckSheetPrintButton.Visible = false;
                        break;
                    case 5:
                        CheckSheetPrintButton.Visible = false;
                        break;
                    case 7:
                        SubstrateListPrintButton.Visible = false;
                        break;
                }

                switch (ProductInfo.PrintType) {
                    case 1:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 9:
                        SettingsLabelPro = new CSettingsLabelPro();
                        labelSettingFilePath = $"./config/{ProductInfo.ProductName}/SerialConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml";

                        if (ProductInfo.PrintType == 1) { SubstrateListPrintButton.Visible = false; }
                        SerialPrintPostionLabel.Visible = true;
                        SerialPrintPostionNumericUpDown.Visible = true;
                        BarcodePrintPostionLabel.Visible = false;
                        BarcodePrintPostionNumericUpDown.Visible = false;
                        シリアルラベル印刷ToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷プレビューToolStripMenuItem.Enabled = true;
                        バーコード印刷ToolStripMenuItem.Enabled = false;
                        バーコード印刷プレビューToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷設定ToolStripMenuItem.Enabled = true;
                        バーコード印刷設定ToolStripMenuItem.Enabled = false;
                        break;
                    case 2:
                        SettingsBarcodePro = new CSettingsBarcodePro();
                        barcodeSettingFilePath = $"./config/{ProductInfo.ProductName}/BarcodeConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml";

                        SerialPrintPostionLabel.Visible = false;
                        SerialPrintPostionNumericUpDown.Visible = false;
                        BarcodePrintPostionLabel.Visible = true;
                        BarcodePrintPostionNumericUpDown.Visible = true;
                        シリアルラベル印刷ToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷プレビューToolStripMenuItem.Enabled = false;
                        バーコード印刷ToolStripMenuItem.Enabled = false;
                        バーコード印刷プレビューToolStripMenuItem.Enabled = true;
                        シリアルラベル印刷設定ToolStripMenuItem.Enabled = false;
                        バーコード印刷設定ToolStripMenuItem.Enabled = true;
                        break;
                    case 3:
                        SettingsLabelPro = new CSettingsLabelPro();
                        labelSettingFilePath = $"./config/{ProductInfo.ProductName}/SerialConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml";
                        SettingsBarcodePro = new CSettingsBarcodePro();
                        barcodeSettingFilePath = $"./config/{ProductInfo.ProductName}/BarcodeConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml";

                        SerialPrintPostionLabel.Visible = true;
                        SerialPrintPostionNumericUpDown.Visible = true;
                        BarcodePrintPostionLabel.Visible = true;
                        BarcodePrintPostionNumericUpDown.Visible = true;
                        シリアルラベル印刷ToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷プレビューToolStripMenuItem.Enabled = true;
                        バーコード印刷ToolStripMenuItem.Enabled = false;
                        バーコード印刷プレビューToolStripMenuItem.Enabled = true;
                        シリアルラベル印刷設定ToolStripMenuItem.Enabled = true;
                        バーコード印刷設定ToolStripMenuItem.Enabled = true;
                        break;
                    case 8:
                        SerialPrintPostionLabel.Visible = false;
                        SerialPrintPostionNumericUpDown.Visible = false;
                        BarcodePrintPostionLabel.Visible = false;
                        BarcodePrintPostionNumericUpDown.Visible = false;
                        シリアルラベル印刷ToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷プレビューToolStripMenuItem.Enabled = false;
                        バーコード印刷ToolStripMenuItem.Enabled = false;
                        バーコード印刷プレビューToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷設定ToolStripMenuItem.Enabled = true;
                        バーコード印刷設定ToolStripMenuItem.Enabled = false;
                        break;
                    case 0:
                    default:
                        SerialPrintPostionLabel.Visible = false;
                        SerialPrintPostionNumericUpDown.Visible = false;
                        BarcodePrintPostionLabel.Visible = false;
                        BarcodePrintPostionNumericUpDown.Visible = false;
                        シリアルラベル印刷ToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷プレビューToolStripMenuItem.Enabled = false;
                        バーコード印刷ToolStripMenuItem.Enabled = false;
                        バーコード印刷プレビューToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷設定ToolStripMenuItem.Enabled = false;
                        バーコード印刷設定ToolStripMenuItem.Enabled = false;
                        break;
                }

                LoadSettings(labelSettingFilePath, barcodeSettingFilePath);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                Debug.Print(barcodeSettingFilePath);
                if (barcodeSettingFilePath != string.Empty) {
                    using StreamReader? srBarcode = new(barcodeSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer serializerBarcode = new(typeof(CSettingsBarcodePro));
                    if (serializerBarcode.Deserialize(srBarcode) is CSettingsBarcodePro result) { SettingsBarcodePro = result; }
                    srBarcode?.Close();
                }
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // 登録処理
        private bool Registration() {
            if (ProductInfo.RegType > 0) {
                using SQLiteConnection con = new(GetConnectionRegistration());
                con.Open();

                foreach (var b in _strSerial) {
                    using var cmd = con.CreateCommand();
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
                    using (SQLiteConnection con = new(GetConnectionRegistration())) {
                        con.Open();

                        using var cmd = con.CreateCommand();
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
                    }
                    break;

                case 1:
                    using (SQLiteConnection con = new(GetConnectionRegistration())) {
                        con.Open();

                        using var cmd = con.CreateCommand();
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
                        cmd.Parameters.Add("@SerialLastNumber", DbType.String).Value = _serialLastNumber;
                        cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                        cmd.ExecuteNonQuery();
                    }
                    break;

                case 2:
                case 3:
                    using (SQLiteConnection con = new(GetConnectionRegistration())) {
                        con.Open();
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
                                        var substrateNum = objDgv.Rows[j].Cells[0].Value.ToString() ?? string.Empty;
                                        var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value);
                                        var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);

                                        using (var cmd = con.CreateCommand()) {
                                            cmd.CommandText =
                                                $"""
                                                UPDATE "{ProductInfo.StockName}_Stock" SET
                                                    Stock = @Stock,
                                                    History = CASE
                                                                  WHEN ifnull(History, '') = '' THEN @History
                                                                  ELSE History || ',' || char(10) || @History
                                                              END
                                                WHERE
                                                    SubstrateNumber = @SubstrateNumber
                                                """;

                                            cmd.Parameters.Add("@SubstrateNumber", DbType.String).Value = objDgv.Rows[j].Cells[0].Value;

                                            cmd.Parameters.Add("@Stock", DbType.String).Value = stockValue - useValue;
                                            cmd.Parameters.Add("@History", DbType.String).Value = $"{ProductInfo.ProductNumber}({useValue}){Environment.NewLine}";

                                            cmd.ExecuteNonQuery();

                                            cmd.CommandText = $"SELECT * FROM {ProductInfo.StockName}_Stock WHERE SubstrateModel = @SubstrateModel ORDER BY _rowid_ ASC";
                                            cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];

                                            using var dr = cmd.ExecuteReader();
                                            while (dr.Read()) {
                                                substrateName = $"{dr["SubstrateName"]}";
                                                substrateModel = $"{dr["SubstrateModel"]}";
                                            }
                                        }

                                        if (useValue != 0) {
                                            subTotalTemp = string.IsNullOrEmpty(subTotalTemp) ? $"{substrateNum}({useValue})" : $"{subTotalTemp},{substrateNum}({useValue})";
                                        }

                                        using (var cmd = con.CreateCommand()) {
                                            cmd.CommandText =
                                                $"""
                                                INSERT INTO "{ProductInfo.StockName}_Substrate"
                                                    (
                                                    SubstrateName,
                                                    SubstrateModel,
                                                    SubstrateNumber,
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
                                            cmd.Parameters.Add("@Decrease", DbType.String).Value = 0 - useValue;
                                            cmd.Parameters.Add("@UsedProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                                            cmd.Parameters.Add("@UsedProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                                            cmd.Parameters.Add("@UsedOrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                                            cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                                            cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                                            cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                                            cmd.ExecuteNonQuery();
                                        }

                                        switch (ProductInfo.PrintType) {
                                            case 5:
                                            case 6:
                                                _usedSubstrate.Add(_useSubstrate[i]);
                                                if (substrateNum != null) { _usedProductNumber.Add(substrateNum); }
                                                _usedQuantity.Add(useValue);
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                }
                                _totalSubstrate = string.IsNullOrEmpty(_totalSubstrate)
                                    ? $"[{_useSubstrate[i]}]{subTotalTemp}"
                                    : $"{_totalSubstrate},{Environment.NewLine}[{_useSubstrate[i]}]{subTotalTemp}";
                                subTotalTemp = string.Empty;
                            }
                        }

                        using (var cmd = con.CreateCommand()) {
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
                            cmd.Parameters.Add("@SerialLastNumber", DbType.String).Value = _serialLastNumber;
                            cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;
                            cmd.Parameters.Add("@UsedSubstrate", DbType.String).Value = string.IsNullOrWhiteSpace(_totalSubstrate) ? DBNull.Value : _totalSubstrate;

                            cmd.ExecuteNonQuery();
                        }
                    }
                    break;
            }
            return true;
        }
        private void RegisterCheck() {
            try {
                _strSerial.Clear();

                if (!NumberCheck()) { return; }
                if (!QuantityCheck()) { return; }
                if (!SerialCheck()) { return; }

                // ラベル印刷
                switch (ProductInfo.PrintType) {
                    case 1:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 9:
                        MessageBox.Show("シリアルラベルを印刷します。");
                        _serialType = "Label";
                        if (!PrintBarcode(1)) { throw new Exception("キャンセルしました。"); }
                        break;
                    default:
                        break;
                }
                // バーコード印刷
                switch (ProductInfo.PrintType) {
                    case 2:
                    case 3:
                        MessageBox.Show("バーコードラベルを印刷します。");
                        _serialType = "Barcode";
                        if (!PrintBarcode(1)) { throw new Exception("キャンセルしました。"); }
                        break;
                    default:
                        break;
                }

                RegisterButton.Enabled = false;
                SerialPrintPostionNumericUpDown.Enabled = false;
                BarcodePrintPostionNumericUpDown.Enabled = false;

                // シリアル先頭と末尾を生成
                switch (ProductInfo.PrintType) {
                    case 1:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                    case 9:
                        _serialType = "Label";
                        _serialFirst = GenerateCode(ProductInfo.SerialFirstNumber);
                        _serialLast = GenerateCode(_serialLastNumber);
                        break;
                    case 2:
                        _serialType = "Barcode";
                        _serialFirst = GenerateCode(ProductInfo.SerialFirstNumber);
                        _serialLast = GenerateCode(_serialLastNumber);
                        break;
                    default:
                        break;
                }

                if (!Registration()) { throw new Exception("登録失敗しました。"); }

                MessageBox.Show("登録完了");

                // バックアップ作成
                CreateBackup();

                switch (ProductInfo.PrintType) {
                    case 0:
                    case 1:
                    case 2:
                    case 3:
                    case 4:
                    case 8:
                    case 9:
                        Close();
                        break;
                    case 5:
                    case 6:
                    case 7:
                        foreach (Control control in Controls) {
                            if (control is DataGridView dataGridView) {
                                dataGridView.Enabled = false;
                            }
                            if (control is System.Windows.Forms.CheckBox checkBox) {
                                checkBox.Enabled = false;
                            }
                        }
                        if (ProductInfo.PrintType is 5 or 6) { SubstrateListPrintButton.Enabled = true; }
                        if (ProductInfo.PrintType is 6 or 7) { CheckSheetPrintButton.Enabled = true; ; }
                        break;
                    default:
                        break;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool NumberCheck() {
            using SQLiteConnection con = new(GetConnectionRegistration());
            con.Open();

            var productModel = string.Empty;

            if (!string.IsNullOrEmpty(ProductInfo.ProductNumber)) {
                // 製番が新規かチェック
                using (var cmd = con.CreateCommand()) {
                    cmd.CommandText = $"""SELECT * FROM "{ProductInfo.ProductName}_Product" WHERE ProductNumber = @ProductNumber ORDER BY _rowid_ ASC LIMIT 1""";
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
                    cmd.CommandText = $"""SELECT * FROM "{ProductInfo.ProductName}_Product" WHERE OrderNumber = @OrderNumber ORDER BY _rowid_ ASC LIMIT 1""";
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
                        MessageBox.Show($"[{ProductInfo.OrderNumber}]は[{productModel}]として登録があります。確認してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
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
                        if (_useSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (var i = 0; i <= _useSubstrate.GetUpperBound(0); i++) {

                            var objCbx = Controls[_checkBoxNames[i]] as System.Windows.Forms.CheckBox ?? throw new Exception("objCbxがnullです。");
                            objCbx.Enabled = true;
                            objCbx.Checked = true;

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
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private bool SerialCheck() {
            try {
                switch (ProductInfo.PrintType) {
                    case 1:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                    case 9:
                        for (var i = 0; i < ProductInfo.Quantity; i++) {
                            _serialType = "Label";
                            _strSerial.Add(GenerateCode(ProductInfo.SerialFirstNumber + i));
                        }
                        break;
                    case 2:
                        for (var i = 0; i < ProductInfo.Quantity; i++) {
                            _serialType = "Barcode";
                            _strSerial.Add(GenerateCode(ProductInfo.SerialFirstNumber + i));
                        }
                        break;
                    default:
                        break;
                }

                var strSQLSerial = string.Join("','", _strSerial);

                List<string> strSerialDuplication = [];
                using (SQLiteConnection con = new(GetConnectionRegistration())) {
                    con.Open();

                    using var cmd = con.CreateCommand();
                    cmd.CommandText = $"""SELECT Serial FROM "{ProductInfo.ProductName}_Serial" WHERE Serial IN (@Serial)""";
                    cmd.Parameters.Add("@Serial", DbType.String).Value = strSQLSerial;

                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        strSerialDuplication.Add($"{dr["Serial"]}");
                    }
                }

                if (strSerialDuplication.Count > 0) {
                    var strSQLDuplication = string.Join($"{Environment.NewLine}", strSerialDuplication);
                    throw new Exception($"{strSQLDuplication}{Environment.NewLine}は既に使用されているシリアルです。");
                }

                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // 印刷処理
        private bool PrintBarcode(int printFlg) {
            // PrintDocumentオブジェクトの作成
            using System.Drawing.Printing.PrintDocument pd = new();

            // PrintPageイベントハンドラの追加
            pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

            _labelProNumLabelsToPrint = ProductInfo.Quantity;

            switch (printFlg) {
                case 1:
                    ProductRegistration2PrintDialog.Document = pd;
                    var r = ProductRegistration2PrintDialog.ShowDialog();

                    if (r == DialogResult.OK) {
                        ProductRegistration2PrintDialog.Document.Print();

                        if (_pageCnt >= 2) {
                            MessageBox.Show($"{_pageCnt}枚印刷されます。2枚目以降は1行目から印刷されます。");
                        }
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
                    ProductRegistration2PrintPreviewDialog.Document = pd;
                    ProductRegistration2PrintPreviewDialog.ShowDialog();
                    break;
            }
            return true;
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                System.Drawing.Point headerPos = new(0, 0);
                var headerString = string.Empty;
                Font headerFooterFont = new("ＭＳ Ｐ明朝", 5.25F);
                var serialCodePrintCopies = 0;
                var remainingCount = 0;

                var maxX = 0;
                var maxY = 0;
                decimal sizeX = 0;
                decimal sizeY = 0;
                decimal offsetX = 0;
                decimal offsetY = 0;
                decimal intervalX = 0;
                decimal intervalY = 0;
                var startLine = 0;
                var labelProPageNum = 0;

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
                        startLine = (int)BarcodePrintPostionNumericUpDown.Value - 1;
                        break;
                    default:
                        break;
                }

                var offset = new System.Drawing.Point(0, 0);
                const decimal MM_PER_HUNDREDTH_INCH = 0.254M;

                var pd = (PrintDocument)sender;
                var bPrintMode = pd.PrintController.IsPreview;

                if (!bPrintMode) {
                    offsetX -= (decimal)e.PageSettings.HardMarginX * MM_PER_HUNDREDTH_INCH;
                    offsetY -= (decimal)e.PageSettings.HardMarginY * MM_PER_HUNDREDTH_INCH;
                    offset = labelProPageNum == 0
                        ? new System.Drawing.Point((int)((decimal)e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (startLine * (intervalY + sizeY))))
                        : new System.Drawing.Point((int)((decimal)e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (0 * (intervalY + sizeY))));
                }
                else {
                    offset = new System.Drawing.Point(0, 0);
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;

                headerPos.Offset(offset);
                e.Graphics.DrawString(headerString, headerFooterFont, Brushes.Black, headerPos);

                // タイプ4か9の場合、発行枚数+1
                if (ProductInfo.PrintType is 4 or 9) { serialCodePrintCopies++; }
                if (_pageCnt == 1) {
                    remainingCount = serialCodePrintCopies;
                    _labelProNSerial = ProductInfo.SerialFirstNumber;
                }
                if (_pageCnt >= 2) { startLine = 0; }

                var y = 0;
                for (y = startLine; y < maxY; y++) {
                    var x = 0;
                    for (x = 0; x < maxX; x++) {
                        var posX = (float)(offsetX + (x * (intervalX + sizeX)));
                        var posY = (float)(offsetY + (y * (intervalY + sizeY)));

                        // タイプ4の場合、最後のラベルに下線をつける
                        var fontUnderline = ProductInfo.PrintType == 4 && remainingCount == 1;

                        // タイプ9かつ最終行の場合はシリアルを型式下4桁に
                        var generatedCode = ProductInfo.PrintType != 9 || remainingCount != 1
                            ? GenerateCode(_labelProNSerial)
                            : ProductInfo.ProductModel[^4..];

                        var labelImage = MakeLabelImage(generatedCode, (int)e.Graphics.DpiX, 1, fontUnderline);
                        e.Graphics.DrawImage(labelImage, posX, posY, (float)sizeX, (float)sizeY);

                        _labelProNSerial++;
                        _labelProNumLabelsToPrint--;

                        // 印刷するラベルがなくなった場合の処理
                        if (_labelProNumLabelsToPrint <= 0) {
                            remainingCount--;

                            if (remainingCount <= 0) {
                                e.HasMorePages = false;
                                _pageCnt = 1;
                                _labelProNumLabelsToPrint = 0;
                                return;
                            }
                            else {
                                _labelProNSerial -= x + 1;
                                _labelProNumLabelsToPrint += x + 1;
                                break;
                            }
                        }

                        // 列の終わりの処理
                        if (x >= maxX - 1) {
                            remainingCount--;
                            if (remainingCount <= 0) {
                                remainingCount = serialCodePrintCopies;
                            }
                            else if (remainingCount > 0) {
                                _labelProNSerial -= x + 1;
                                _labelProNumLabelsToPrint += x + 1;
                                break;
                            }
                        }
                    }
                }

                if (_labelProNumLabelsToPrint > 0) {
                    _pageCnt++;
                    e.HasMorePages = true;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                _ => string.Empty
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
                    stringSize = g.MeasureString(text, fnt);

                    stringPosX = (float)((labelImage.Width / 2) - (stringSize.Width / 2));

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
                    stringSize = g.MeasureString(text, fnt);

                    stringPosX = SettingsBarcodePro.BarcodeProLabelSettings.AlignStringCenter
                        ? (float)((labelImage.Width / 2) - (stringSize.Width / 2))
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
                default:
                    break;
            }

            dataGridView.Enabled = checkBox.Checked;
            checkBox.ForeColor = checkBox.Checked ? Color.Black : Color.Red;

            if (!checkBox.Checked) {
                MessageBox.Show("チェックがない場合在庫から引き落とされなくなります。", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        // チェックシート印刷
        private static void CheckSheetPrint() {
            try {
                //List<string> _filePath = new();
                //List<string> _sheetName = new();

                //List<string> _orderNumberRange = new();
                //List<string> _regYearRange = new();
                //List<string> _regMonthRange = new();
                //List<string> _regDayRange = new();
                //List<string> _orderFirstSerialRange = new();
                //List<string> _orderLastSerialRange = new();

                //using FileStream _fileStream = new($"""{Environment.CurrentDirectory}./config/Excel/ConfigCheckSheet.xlsx""", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                //using XLWorkbook _workBook = new(_fileStream);
                //IXLWorksheet _workSheet = _workBook.Worksheet("Sheet1");

                //// セル検索
                //foreach (IXLCell _cell in _workSheet.Search(StrProductModel)) {
                //    int _findRow = _cell.Address.RowNumber;
                //    // ワークシートのセルから値を取得してリストに格納
                //    _filePath.Add(_workSheet.Cell("J" + _findRow).Value.ToString());
                //    _sheetName.Add(_workSheet.Cell("I" + _findRow).Value.ToString());

                //    _orderNumberRange.Add(_workSheet.Cell("C" + _findRow).Value.ToString());
                //    _regYearRange.Add(_workSheet.Cell("D" + _findRow).Value.ToString());
                //    _regMonthRange.Add(_workSheet.Cell("E" + _findRow).Value.ToString());
                //    _regDayRange.Add(_workSheet.Cell("F" + _findRow).Value.ToString());
                //    _orderFirstSerialRange.Add(_workSheet.Cell("G" + _findRow).Value.ToString());
                //    _orderLastSerialRange.Add(_workSheet.Cell("H" + _findRow).Value.ToString());
                //}

                //if (_filePath.Count == 0) {
                //    throw new Exception($"Configに品目番号:{StrProductModel}が見つかりません。");
                //}

                //// 見つかった品目番号の数だけ印刷
                //for (int _i = 0; _i < _filePath.Count; _i++) {
                //    // チェックシート作成
                //    using FileStream _fileStream2 = new($"""{_filePath[_i]}""", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                //    using XLWorkbook _workBook2 = new(_fileStream2);
                //    IXLWorksheet _workSheet2 = _workBook2.Worksheet(_sheetName[_i]);

                //    _workSheet2.Cell(_orderNumberRange[_i]).Value = StrOrderNumber;
                //    _workSheet2.Cell(_regYearRange[_i]).Value = $"{DateTime.Parse(StrRegDate):yy}年";
                //    _workSheet2.Cell(_regMonthRange[_i]).Value = $"{DateTime.Parse(StrRegDate):MM}月";
                //    _workSheet2.Cell(_regDayRange[_i]).Value = $"{DateTime.Parse(StrRegDate):dd}日";
                //    _workSheet2.Cell(_orderFirstSerialRange[_i]).Value = StrSerialFirstNumber;
                //    _workSheet2.Cell(_orderLastSerialRange[_i]).Value = StrSerialLastNumber;

                //    // 印刷
                //    Excel.Application _xlApp = new() {
                //        Visible = true // Excelウィンドウを表示します。
                //    };

                //    // ワークブックを開く。
                //    Excel.Workbooks _xlBooks = _xlApp.Workbooks;
                //    Excel.Workbook _xlBook = _xlBooks.Open($"""{System.Environment.CurrentDirectory}./config/Excel/temporarily_{_i}.xlsx""");

                //    // ワークシートを選択
                //    Excel.Sheets _xlSheets = _xlBook.Sheets;
                //    Excel.Worksheet _xlSheet = _xlSheets[1];

                //    // ワークシートを印刷
                //    _xlSheet.PrintOut(Preview: true);

                //    // ワークブックを閉じてExcelを終了します。
                //    _xlBook.Close(false);
                //    _xlApp.Quit();
                //    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlSheet);
                //    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlSheets);
                //    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlBook);
                //    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlBooks);
                //    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlApp);
                //}

            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // リスト印刷
        private void ListPrint() {
            try {
                var sheetName = string.Empty;
                var productName = string.Empty;
                var productModel = string.Empty;

                var productNameRange = string.Empty;
                var productNumberRange = string.Empty;
                var orderNumberRange = string.Empty;
                var regDateRange = string.Empty;
                var productModelRange = string.Empty;
                var quantityRange = string.Empty;
                var serialFirstRange = string.Empty;
                var serialLastRange = string.Empty;
                var commentRange = string.Empty;
                var qrCodeRange = string.Empty;

                using FileStream fileStream = new($"{Environment.CurrentDirectory}./config/Excel/ConfigList.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using XLWorkbook workBook = new(fileStream);
                var workSheetMain = workBook.Worksheet("Sheet1");

                var findRow = 0;
                // セル検索
                foreach (var cell in workSheetMain.Search(ProductInfo.ProductModel)) {
                    findRow = cell.Address.RowNumber;
                }

                if (findRow == 0) {
                    throw new Exception($"Configに品目番号:{ProductInfo.ProductModel}が見つかりません。");
                }

                // ワークシートのセルから値を取得
                sheetName = ProductInfo.ProductModel;
                productName = workSheetMain.Cell(findRow, 3).Value.ToString();
                productNameRange = workSheetMain.Cell(findRow, 4).Value.ToString();
                productNumberRange = workSheetMain.Cell(findRow, 5).Value.ToString();
                orderNumberRange = workSheetMain.Cell(findRow, 6).Value.ToString();
                regDateRange = workSheetMain.Cell(findRow, 7).Value.ToString();
                productModel = workSheetMain.Cell(findRow, 8).Value.ToString();
                productModelRange = workSheetMain.Cell(findRow, 9).Value.ToString();
                quantityRange = workSheetMain.Cell(findRow, 10).Value.ToString();
                serialFirstRange = workSheetMain.Cell(findRow, 11).Value.ToString();
                serialLastRange = workSheetMain.Cell(findRow, 12).Value.ToString();
                commentRange = workSheetMain.Cell(findRow, 13).Value.ToString();
                qrCodeRange = workSheetMain.Cell(findRow, 14).Value.ToString();

                var workSheetTemp = workBook.Worksheet(sheetName);
                workSheetTemp.Cell(productNameRange).Value = productName;
                workSheetTemp.Cell(productNumberRange).Value = ProductInfo.ProductNumber;
                workSheetTemp.Cell(orderNumberRange).Value = ProductInfo.OrderNumber;
                workSheetTemp.Cell(regDateRange).Value = ProductInfo.RegDate;
                workSheetTemp.Cell(productModelRange).Value = productModel;
                workSheetTemp.Cell(quantityRange).Value = ProductInfo.Quantity;
                workSheetTemp.Cell(serialFirstRange).Value = _serialFirst;
                workSheetTemp.Cell(serialLastRange).Value = _serialLast;
                workSheetTemp.Cell(commentRange).Value = ProductInfo.Comment;

                var i = 0;
                var findColumn = 0;
                for (i = 0; i <= _usedSubstrate.Count - 1; i++) {

                    var searchRange = workSheetMain.Range(findRow, 1, findRow, 28);
                    var searchValue = $"{_usedSubstrate[i]}";
                    var foundCell = searchRange.CellsUsed(c => c.Value.ToString() == searchValue).FirstOrDefault();

                    if (foundCell != null) {
                        // セルが見つかった場合の処理
                        var foundRow = foundCell.Address.RowNumber;
                        var foundColumn = foundCell.Address.ColumnNumber;
                    }

                    foreach (var cell in workSheetMain.Search(_usedSubstrate[i])) {
                        if (cell.Address.RowNumber == findRow) {
                            findColumn = cell.Address.ColumnNumber;
                            break;
                        }
                    }

                    if (findColumn == 0) {
                        throw new Exception($"{_usedSubstrate[i]}が見つかりません。");
                    }

                    var mainCellValue = workSheetMain.Cell(findRow, findColumn + 1).Value.ToString();
                    var tempCellValue = workSheetTemp.Cell(mainCellValue).Value.ToString();

                    if (mainCellValue != string.Empty) {
                        if (tempCellValue == string.Empty) {
                            workSheetTemp.Cell(mainCellValue).Value = $"{_usedProductNumber[i]}({_usedQuantity[i]})";
                        }
                        else {
                            workSheetTemp.Cell(mainCellValue).Value += $"    {_usedProductNumber[i]}({_usedQuantity[i]})";
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
                    var bmpData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                    System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bmpData.Scan0, pixelData.Pixels.Length);
                    bitmap.UnlockBits(bmpData);
                    using MemoryStream stream = new();
                    bitmap.Save(stream, ImageFormat.Bmp);

                    var image = workSheetTemp.AddPicture(stream);
                    image.MoveTo(workSheetTemp.Cell(qrCodeRange));
                }

                //引数に保存先パスを指定
                workBook.SaveAs($"{Environment.CurrentDirectory}./config/Excel/temporarily.xlsx");

                // 印刷
                Excel.Application xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                var xlBooks = xlApp.Workbooks;
                var xlBook = xlBooks.Open($"{Environment.CurrentDirectory}./config/Excel/temporarily.xlsx");

                // ワークシート選択
                var xlSheets = xlBook.Sheets;
                Excel.Worksheet xlSheet = xlSheets[sheetName];

                // ワークシート印刷
                xlSheet.PrintOut(Preview: true);

                // ワークブックを閉じてExcelを終了
                xlBook.Close(false);
                xlApp.Quit();

                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // バックアップ作成
        public static void CreateBackup() {
            const string BackupDirectory = "./db/backups"; // バックアップを保存するディレクトリ
            const string OriginalFilePath = "./db/registration.db"; // 元ファイルパス
            const int MaxBackupFiles = 10; // 最大バックアップファイル数
            // バックアップ用ディレクトリが存在しない場合は作成
            if (!Directory.Exists(BackupDirectory)) {
                Directory.CreateDirectory(BackupDirectory);
            }

            // 日付と時間をファイル名に付加
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var backupFileName = $"registration_{timestamp}.db";
            var backupFilePath = Path.Combine(BackupDirectory, backupFileName);

            // 元ファイルをバックアップにコピー
            File.Copy(OriginalFilePath, backupFilePath, true);

            // バックアップファイルを管理
            ManageBackupFiles(BackupDirectory, MaxBackupFiles);

            static void ManageBackupFiles(string backupDirectory, int maxBackupFiles) {
                // バックアップディレクトリ内のファイルを取得
                var backupFiles = Directory.GetFiles(backupDirectory, "registration_*.db")
                                           .OrderBy(f => File.GetCreationTime(f)) // 作成日時順に並べる
                                           .ToList();

                // バックアップが最大数を超えている場合は古いものを削除
                while (backupFiles.Count > maxBackupFiles) {
                    var oldestFile = backupFiles.First();
                    File.Delete(oldestFile);
                    backupFiles.RemoveAt(0);
                }
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
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) { ListPrint(); }
        private void CheckSheetPrintButton_Click(object sender, EventArgs e) { CheckSheetPrint(); }
        private void ProductRegistration2PrintPreviewDialog_Load(object sender, EventArgs e) {
            var tool = (ToolStrip)ProductRegistration2PrintPreviewDialog.Controls[1];
            tool.Items[0].Visible = false;
        }
    }
}
