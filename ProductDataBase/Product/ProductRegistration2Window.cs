using ClosedXML.Excel;
using GenCode128;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using Color = System.Drawing.Color;
using Control = System.Windows.Forms.Control;
using Excel = Microsoft.Office.Interop.Excel;
using Font = System.Drawing.Font;

namespace ProductDatabase {
    public partial class ProductRegistration2Window : Form {

        public CSettingsLabelPro SettingsLabelPro { get; set; } = new CSettingsLabelPro();
        public string StrLabelSettingFilePath { get; set; } = string.Empty;

        public CSettingsBarcodePro SettingsBarcodePro { get; set; } = new CSettingsBarcodePro();
        public string StrBarcodeSettingFilePath { get; set; } = string.Empty;

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public string StrProductName { get; set; } = string.Empty;
        public string StrStockName { get; set; } = string.Empty;
        public string StrProductType { get; set; } = string.Empty;
        public string StrProductModel { get; set; } = string.Empty;
        public string StrUseSubstrate { get; set; } = string.Empty;
        public string[] ArrUseSubstrate = [];
        public string StrInitial { get; set; } = string.Empty;
        public string StrOrderNumber { get; set; } = string.Empty;
        public string StrProductNumber { get; set; } = string.Empty;
        public string StrRegDate { get; set; } = string.Empty;
        public string StrPerson { get; set; } = string.Empty;
        public string StrRevision { get; set; } = string.Empty;
        public string StrComment { get; set; } = string.Empty;

        public int IntQuantity { get; set; }
        public int IntRegType { get; set; }
        public int IntPrintType { get; set; }
        public int IntCheckBin { get; set; }
        public int IntSerialDigit { get; set; }
        public int IntSerialFirstNumber { get; set; }

        public int LabelProNSerial { get; set; }
        public int LabelProNumLabelsToPrint { get; set; }

        public decimal DisplayResolution { get; } = 96.0m;
        public int DisplayMagnitude { get; } = 3;
        public int IntPageCnt { get; set; } = 1;

        public string StrProness1 { get; } = string.Empty;
        public string StrProness2 { get; } = string.Empty;
        public string StrProness3 { get; } = string.Empty;
        public int StrProness4 { get; }
        public string StrProness5 { get; } = string.Empty;

        private string strSerialType = string.Empty;
        private string strSerialFirstNumber = string.Empty;
        private string strSerialLastNumber = string.Empty;
        private string strTotalSubstrate = string.Empty;
        private int intSerialLastNumber;
        private bool fontUnderbar = false;
        private readonly List<string> strSerial = [];
        private readonly List<string> checkBoxNames = [
                        "Substrate1CheckBox", "Substrate2CheckBox", "Substrate3CheckBox", "Substrate4CheckBox","Substrate5CheckBox",
                        "Substrate6CheckBox", "Substrate7CheckBox", "Substrate8CheckBox", "Substrate9CheckBox","Substrate10CheckBox"
                        ];
        private readonly List<string> dataGridViewNames = [
                        "Substrate1DataGridView", "Substrate2DataGridView", "Substrate3DataGridView", "Substrate4DataGridView","Substrate5DataGridView",
                        "Substrate6DataGridView", "Substrate7DataGridView", "Substrate8DataGridView", "Substrate9DataGridView","Substrate10DataGridView"
                        ];
        private readonly List<string> listUsedSubstrate = [];
        private readonly List<string> listUsedProductNumber = [];
        private readonly List<int> listUsedQuantity = [];

        public ProductRegistration2Window() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(StrFontName, IntFontSize);

                RegisterButton.Enabled = true;
                ArrUseSubstrate = StrUseSubstrate.Split(",");

                intSerialLastNumber = IntSerialFirstNumber + IntQuantity - 1;

                var _quantityFlg = false;
                var _strQuantity = string.Empty;
                switch (IntRegType) {
                    case 2:
                        for (var _i = 0; _i <= ArrUseSubstrate.GetUpperBound(0); _i++) {
                            var _objCbx = Controls[checkBoxNames[_i]] as CheckBox;

                            if (_objCbx != null) {
                                _objCbx.Enabled = true;
                                _objCbx.Checked = true;
                            }

                            var _objDgv = Controls[dataGridViewNames[_i]] as DataGridView;
                            if (_objDgv != null) {
                                _objDgv.RowHeadersWidth = 30;
                                _objDgv.Columns[2].ReadOnly = false;
                                _objDgv.Columns[3].ReadOnly = false;
                            }

                            using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
                            _con.Open();

                            using var _cmd = _con.CreateCommand();
                            // 使用基板表示
                            _cmd.CommandText = $"SELECT col_Substrate_Name FROM Stock_{StrStockName} WHERE col_Substrate_Model = @col_Substrate_Model";
                            _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ArrUseSubstrate[_i];
                            using (var _dr = _cmd.ExecuteReader()) {
                                if (_dr.Read()) {
                                    if (_objCbx != null) {
                                        var _substrateName = $"{_dr["col_Substrate_Name"]}";
                                        _objCbx.Text = $"{_substrateName} - {ArrUseSubstrate[_i]}";
                                    }
                                }
                            }

                            // 在庫テーブルからデータ取得
                            var _intQuantity = IntQuantity;
                            _cmd.CommandText = $"SELECT col_Substrate_num, col_Stock, col_Substrate_Name FROM Stock_{StrStockName} WHERE col_flg = 1 AND col_Substrate_Model = @col_Substrate_Model ORDER BY _rowid_ ASC";
                            _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ArrUseSubstrate[_i];
                            using (var _dr = _cmd.ExecuteReader()) {
                                while (_dr.Read()) {
                                    var _strSubstrateNumber = $"{_dr["col_Substrate_num"]}";
                                    var _intStock = Convert.ToInt32(_dr["col_Stock"]);
                                    _objDgv?.Rows.Add(_strSubstrateNumber, _intStock);

                                    if (_intQuantity >= _intStock) {
                                        _intQuantity -= _intStock;
                                        if (_objDgv != null) {
                                            _objDgv.Rows[^1].Cells[2].Value = _intStock;
                                            _objDgv.Rows[^1].Cells[3].Value = true;
                                        }
                                    }
                                    else {
                                        if (_intQuantity == 0) {
                                            if (_objDgv != null) {
                                                _objDgv.Rows[^1].Cells[2].Value = 0;
                                            }
                                        }
                                        else {
                                            if (_objDgv != null) {
                                                _objDgv.Rows[^1].Cells[2].Value = _intQuantity;
                                                _objDgv.Rows[^1].Cells[3].Value = true;
                                                _intQuantity = 0;
                                            }
                                        }
                                    }

                                    if (_intQuantity > 0) {
                                        _quantityFlg = false;
                                        var _substrateName = $"{_dr["col_Substrate_Name"]}";
                                        _strQuantity += $"[{_substrateName}]{Environment.NewLine}";
                                    }

                                    if (_intQuantity == 0) { _quantityFlg = true; }
                                }
                            }
                        }

                        if (_quantityFlg == false) {
                            Activate();
                            MessageBox.Show($"在庫が足りません。{Environment.NewLine}{_strQuantity}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        break;
                    case 3:
                        if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (var _i = 0; _i <= ArrUseSubstrate.GetUpperBound(0); _i++) {
                            var _objCbx = Controls[checkBoxNames[_i]] as CheckBox;

                            if (_objCbx != null) {
                                _objCbx.Enabled = true;
                                _objCbx.Checked = true;
                            }

                            var _objDgv = Controls[dataGridViewNames[_i]] as DataGridView;
                            if (_objDgv != null) {
                                _objDgv.RowHeadersWidth = 30;
                                _objDgv.Columns[2].ReadOnly = false;
                                _objDgv.Columns[3].ReadOnly = false;
                            }

                            using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
                            _con.Open();
                            using var _cmd = _con.CreateCommand();
                            _cmd.CommandText = $"SELECT * FROM Stock_{StrStockName} WHERE col_Substrate_Model = @col_Substrate_Model";
                            _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ArrUseSubstrate[_i];
                            using (var _dr = _cmd.ExecuteReader()) {
                                while (_dr.Read()) {
                                    if (_objCbx != null) { _objCbx.Text = $"{_dr["col_Substrate_Name"]} - {ArrUseSubstrate[_i]}"; }
                                }
                            }

                            _cmd.CommandText = $"SELECT * FROM Stock_{StrStockName} WHERE col_flg = 1 AND col_Substrate_Model = @col_Substrate_Model ORDER BY _rowid_ ASC";
                            _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ArrUseSubstrate[_i];
                            using (var _dr = _cmd.ExecuteReader()) {
                                while (_dr.Read()) {
                                    var _strSubstrateName = string.Empty;
                                    _strSubstrateName = $"{_dr["col_Substrate_Name"]}";

                                    var _strSubstrateNumber = $"{_dr["col_Substrate_Num"]}";
                                    var _intStock = Convert.ToInt32(_dr["col_Stock"]);
                                    _objDgv?.Rows.Add(_strSubstrateNumber, _intStock);

                                    var _j = 0;
                                    var _strOrderNumber = $"{_dr["col_Order_Num"]}";
                                    if (_strOrderNumber == StrOrderNumber) {
                                        if (_objDgv != null) {
                                            var _intQuantity = IntQuantity;
                                            var _stockValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[1].Value);
                                            var _useValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[2].Value);
                                            _objDgv.Rows[_j].Cells[2].Value = _intQuantity;
                                            _objDgv.Rows[_j].Cells[3].Value = true;
                                            // 必要数量分割り当てられたかチェック
                                            if (_intQuantity > _stockValue) {
                                                _quantityFlg = false;
                                                _strQuantity = $"{_strQuantity}[{_strSubstrateName}]{Environment.NewLine}";
                                            }
                                            if (_stockValue >= _useValue) { _quantityFlg = true; }
                                        }
                                    }
                                    _j++;
                                }
                            }
                        }

                        if (_quantityFlg == false) {
                            Activate();
                            MessageBox.Show($"在庫が足りません。{Environment.NewLine}{_strQuantity}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        break;
                    default:
                        break;
                }

                switch (IntPrintType) {
                    case 0:
                        SerialPrintPostionNumericUpDown.Enabled = false;
                        BarcodePrintPostionNumericUpDown.Enabled = false;
                        シリアルラベル印刷ToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷プレビューToolStripMenuItem.Enabled = false;
                        バーコード印刷ToolStripMenuItem.Enabled = false;
                        バーコード印刷プレビューToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷設定ToolStripMenuItem.Enabled = false;
                        バーコード印刷設定ToolStripMenuItem.Enabled = false;

                        break;
                    case 1:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                        SettingsLabelPro = new CSettingsLabelPro();
                        StrLabelSettingFilePath = $"./config/{StrProductName}/SerialConfig_{StrProductName}_{StrProductModel}.xml";

                        SerialPrintPostionNumericUpDown.Enabled = true;
                        BarcodePrintPostionNumericUpDown.Enabled = false;
                        シリアルラベル印刷ToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷プレビューToolStripMenuItem.Enabled = true;
                        バーコード印刷ToolStripMenuItem.Enabled = false;
                        バーコード印刷プレビューToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷設定ToolStripMenuItem.Enabled = true;
                        バーコード印刷設定ToolStripMenuItem.Enabled = false;
                        break;
                    case 2:
                        SettingsBarcodePro = new CSettingsBarcodePro();
                        StrBarcodeSettingFilePath = $"./config/{StrProductName}/BarcodeConfig_{StrProductName}_{StrProductModel}.xml";

                        SerialPrintPostionNumericUpDown.Enabled = false;
                        BarcodePrintPostionNumericUpDown.Enabled = true;
                        シリアルラベル印刷ToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷プレビューToolStripMenuItem.Enabled = false;
                        バーコード印刷ToolStripMenuItem.Enabled = false;
                        バーコード印刷プレビューToolStripMenuItem.Enabled = true;
                        シリアルラベル印刷設定ToolStripMenuItem.Enabled = false;
                        バーコード印刷設定ToolStripMenuItem.Enabled = true;
                        break;
                    case 3:
                        SettingsLabelPro = new CSettingsLabelPro();
                        StrLabelSettingFilePath = $"./config/{StrProductName}/SerialConfig_{StrProductName}_{StrProductModel}.xml";
                        SettingsBarcodePro = new CSettingsBarcodePro();
                        StrBarcodeSettingFilePath = $"./config/{StrProductName}/BarcodeConfig_{StrProductName}_{StrProductModel}.xml";

                        SerialPrintPostionNumericUpDown.Enabled = true;
                        BarcodePrintPostionNumericUpDown.Enabled = true;
                        シリアルラベル印刷ToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷プレビューToolStripMenuItem.Enabled = true;
                        バーコード印刷ToolStripMenuItem.Enabled = false;
                        バーコード印刷プレビューToolStripMenuItem.Enabled = true;
                        シリアルラベル印刷設定ToolStripMenuItem.Enabled = true;
                        バーコード印刷設定ToolStripMenuItem.Enabled = true;
                        break;
                    case 8:
                        SettingsLabelPro = new CSettingsLabelPro();
                        StrLabelSettingFilePath = $"./config/{StrProductName}/SerialConfig_{StrProductName}_{StrProductModel}.xml";

                        SerialPrintPostionNumericUpDown.Enabled = false;
                        BarcodePrintPostionNumericUpDown.Enabled = false;
                        シリアルラベル印刷ToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷プレビューToolStripMenuItem.Enabled = false;
                        バーコード印刷ToolStripMenuItem.Enabled = false;
                        バーコード印刷プレビューToolStripMenuItem.Enabled = false;
                        シリアルラベル印刷設定ToolStripMenuItem.Enabled = true;
                        バーコード印刷設定ToolStripMenuItem.Enabled = false;
                        break;
                    default:
                        break;
                }

                LoadSettings(StrLabelSettingFilePath, StrBarcodeSettingFilePath);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadSettings(string strLabelSettingFilePath, string strBarcodeSettingFilePath) {
            try {
                if (strLabelSettingFilePath != string.Empty) {
                    using StreamReader? _srLabel = new(strLabelSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer _serializerLabel = new(typeof(CSettingsLabelPro));
                    if (_serializerLabel.Deserialize(_srLabel) is CSettingsLabelPro _result) { SettingsLabelPro = _result; }
                    _srLabel?.Close();
                }
                if (strBarcodeSettingFilePath != string.Empty) {
                    using StreamReader? _srBarcode = new(strBarcodeSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer _serializerBarcode = new(typeof(CSettingsBarcodePro));
                    if (_serializerBarcode.Deserialize(_srBarcode) is CSettingsBarcodePro _result) { SettingsBarcodePro = _result; }
                    _srBarcode?.Close();
                }
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // クロージングイベント
        private void ClosingEvents() {
            try {
                System.Xml.Serialization.XmlSerializer _serializerLabel = new(typeof(CSettingsLabelPro));
                StreamWriter? _swLabel = null;
                System.Xml.Serialization.XmlSerializer _serializerBarcode = new(typeof(CSettingsBarcodePro));
                StreamWriter? _swBarcode = null;

                switch (IntPrintType) {
                    case 1:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                        _swLabel = new StreamWriter(StrLabelSettingFilePath, false, new System.Text.UTF8Encoding(false));
                        _serializerLabel.Serialize(_swLabel, SettingsLabelPro);
                        _swLabel?.Close();
                        break;
                    case 2:
                        _swBarcode = new StreamWriter(StrBarcodeSettingFilePath, false, new System.Text.UTF8Encoding(false));
                        _serializerBarcode.Serialize(_swBarcode, SettingsBarcodePro);
                        _swBarcode?.Close();
                        break;
                    case 3:
                        _swLabel = new StreamWriter(StrLabelSettingFilePath, false, new System.Text.UTF8Encoding(false));
                        _serializerLabel.Serialize(_swLabel, SettingsLabelPro);
                        _swLabel?.Close();
                        _swBarcode = new StreamWriter(StrBarcodeSettingFilePath, false, new System.Text.UTF8Encoding(false));
                        _serializerBarcode.Serialize(_swBarcode, SettingsBarcodePro);
                        _swBarcode?.Close();
                        break;
                    default:
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show($"設定ファイルの保存に失敗しました。{Environment.NewLine}{ex.Message}");
            } finally {
                using (ProductRegistration1Window _productRegistration1Window = new()) {
                    _productRegistration1Window.Close();
                }
                Dispose();
            }
        }
        // 登録処理
        private bool Registration() {
            if (IntRegType > 0) {
                using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
                _con.Open();

                foreach (var _b in strSerial) {
                    using var _cmd = _con.CreateCommand();
                    _cmd.CommandText =
                        $"""
                        INSERT INTO Serial_{StrProductName}
                            (
                            col_Serial,
                            col_Order_Num,
                            col_Product_Num,
                            col_Product_Type,
                            col_Product_Model,
                            col_RegDate
                            )
                        VALUES
                            (
                            @col_Serial,
                            @col_Order_Num,
                            @col_Product_Num,
                            @col_Product_Type,
                            @col_Product_Model,
                            @col_RegDate
                            )
                        """;

                    _cmd.Parameters.Add("@col_Serial", DbType.String).Value = _b;
                    _cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                    _cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                    _cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                    _cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                    _cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;

                    _cmd.ExecuteNonQuery();
                }
            }

            switch (IntRegType) {
                case 0:
                    using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                        _con.Open();

                        using var _cmd = _con.CreateCommand();
                        _cmd.CommandText =
                            $"""
                            INSERT INTO Product_Reg_{StrProductName}
                                (
                                col_Order_Num,
                                col_Product_Num,
                                col_Product_Type,
                                col_Product_Model,
                                col_Quantity,
                                col_Person,
                                col_RegDate,
                                col_Revision,
                                col_Comment
                                )
                            VALUES
                                (
                                @col_Order_Num,
                                @col_Product_Num,
                                @col_Product_Type,
                                @col_Product_Model,
                                @col_Quantity,
                                @col_Person,
                                @col_RegDate,
                                @col_Revision,
                                @col_Comment
                                )
                            """;

                        _cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                        _cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                        _cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                        _cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                        _cmd.Parameters.Add("@col_Quantity", DbType.String).Value = IntQuantity;
                        _cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                        _cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                        _cmd.Parameters.Add("@col_Revision", DbType.String).Value = StrRevision;
                        _cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                        _cmd.ExecuteNonQuery();
                    }
                    break;

                case 1:
                    using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                        _con.Open();

                        using var _cmd = _con.CreateCommand();
                        _cmd.CommandText =
                            $"""
                            INSERT INTO Product_Reg_{StrProductName}
                                (
                                col_Order_Num,
                                col_Product_Num,
                                col_Product_Type,
                                col_Product_Model,
                                col_Quantity,
                                col_Person,
                                col_RegDate,
                                col_Revision,
                                col_Serial_First,
                                col_Serial_Last,
                                col_Serial_LastNum,
                                col_Comment
                                )
                            VALUES
                                (
                                @col_Order_Num,
                                @col_Product_Num,
                                @col_Product_Type,
                                @col_Product_Model,
                                @col_Quantity,
                                @col_Person,
                                @col_RegDate,
                                @col_Revision,
                                @col_Serial_First,
                                @col_Serial_Last,
                                @col_Serial_LastNum,
                                @col_Comment
                                )
                            """;

                        _cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                        _cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                        _cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                        _cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                        _cmd.Parameters.Add("@col_Quantity", DbType.String).Value = IntQuantity;
                        _cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                        _cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                        _cmd.Parameters.Add("@col_Revision", DbType.String).Value = StrRevision;
                        _cmd.Parameters.Add("@col_Serial_First", DbType.String).Value = strSerialFirstNumber;
                        _cmd.Parameters.Add("@col_Serial_Last", DbType.String).Value = strSerialLastNumber;
                        _cmd.Parameters.Add("@col_Serial_LastNum", DbType.String).Value = intSerialLastNumber;
                        _cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                        _cmd.ExecuteNonQuery();
                    }
                    break;

                case 2:
                case 3:
                    using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                        _con.Open();
                        if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateがnullです。"); }
                        for (var _i = 0; _i <= ArrUseSubstrate.Length; _i++) {

                            var _objCbx = Controls[checkBoxNames[_i]] as CheckBox ?? throw new Exception("objCbxがnullです。");

                            if (_objCbx.Checked) {
                                var _objDgv = Controls[dataGridViewNames[_i]] as DataGridView ?? throw new Exception("objCbxがnullです。");
                                var _dgvRowCnt = _objDgv.Rows.Count;
                                var _subTotalTemp = string.Empty;

                                for (var _j = 0; _j <= _dgvRowCnt - 1; _j++) {
                                    var _boolCbx = Convert.ToBoolean(_objDgv.Rows[_j].Cells[3].Value);
                                    if (_boolCbx) {
                                        var _substrateName = string.Empty;
                                        var _substrateModel = string.Empty;
                                        var _substrateNum = _objDgv.Rows[_j].Cells[0].Value.ToString() ?? string.Empty;
                                        var _stockValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[1].Value);
                                        var _useValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[2].Value);

                                        using (var _cmd = _con.CreateCommand()) {
                                            _cmd.CommandText =
                                                $"""
                                                UPDATE Stock_{StrStockName} SET
                                                    col_Flg = @col_Flg,
                                                    col_Stock = @col_Stock,
                                                    col_History = ifnull(col_History,'')|| @col_History
                                                WHERE
                                                    col_Substrate_Num = @col_Substrate_Num
                                                """;
                                            _cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = _objDgv.Rows[_j].Cells[0].Value;

                                            _cmd.Parameters.Add("@col_Flg", DbType.String).Value = _stockValue - _useValue == 0 ? 0 : (object)1;

                                            _cmd.Parameters.Add("@col_Stock", DbType.String).Value = _stockValue - _useValue;
                                            _cmd.Parameters.Add("@col_History", DbType.String).Value = $"{StrProductNumber}({_useValue}),";

                                            _cmd.ExecuteNonQuery();

                                            _cmd.CommandText = $"SELECT * FROM Stock_{StrStockName} WHERE col_Substrate_Model = @col_Substrate_Model ORDER BY _rowid_ ASC";
                                            _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ArrUseSubstrate[_i];

                                            using var _dr = _cmd.ExecuteReader();
                                            while (_dr.Read()) {
                                                _substrateName = $"{_dr["col_Substrate_Name"]}";
                                                _substrateModel = $"{_dr["col_Substrate_Model"]}";
                                            }
                                        }

                                        if (_useValue != 0) {
                                            _subTotalTemp = string.IsNullOrEmpty(_subTotalTemp) ? $"{_substrateNum}({_useValue})" : $"{_subTotalTemp},{_substrateNum}({_useValue})";
                                        }

                                        using (var _cmd = _con.CreateCommand()) {
                                            _cmd.CommandText =
                                                $"""
                                                INSERT INTO Substrate_Reg_{StrStockName}
                                                    (
                                                    col_Substrate_Name,
                                                    col_Substrate_Model,
                                                    col_Substrate_Num,
                                                    col_Decrease,
                                                    col_Use_P_Type,
                                                    col_Use_P_Num,
                                                    col_Use_O_Num,
                                                    col_Person,
                                                    col_RegDate,
                                                    col_Comment
                                                    )
                                                VALUES
                                                    (
                                                    @col_Substrate_Name,
                                                    @col_Substrate_Model,
                                                    @col_Substrate_Num,
                                                    @col_Decrease,
                                                    @col_Use_P_Type,
                                                    @col_Use_P_Num,
                                                    @col_Use_O_Num,
                                                    @col_Person,
                                                    @col_RegDate,
                                                    @col_Comment
                                                    )
                                                """;

                                            _cmd.Parameters.Add("@col_Substrate_Name", DbType.String).Value = _substrateName;
                                            _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = _substrateModel;
                                            _cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = _objDgv.Rows[_j].Cells[0].Value;
                                            _cmd.Parameters.Add("@col_Decrease", DbType.String).Value = 0 - _useValue;
                                            _cmd.Parameters.Add("@col_Use_P_Type", DbType.String).Value = StrProductType;
                                            _cmd.Parameters.Add("@col_Use_P_Num", DbType.String).Value = StrProductNumber;
                                            _cmd.Parameters.Add("@col_Use_O_Num", DbType.String).Value = StrOrderNumber;
                                            _cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                                            _cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                                            _cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                                            _cmd.ExecuteNonQuery();
                                        }

                                        switch (IntPrintType) {
                                            case 5:
                                            case 6:
                                                listUsedSubstrate.Add(ArrUseSubstrate[_i]);
                                                if (_substrateNum != null) { listUsedProductNumber.Add(_substrateNum); }
                                                listUsedQuantity.Add(_useValue);
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                }
                                strTotalSubstrate = string.IsNullOrEmpty(strTotalSubstrate)
                                    ? $"[{ArrUseSubstrate[_i]}]{_subTotalTemp}"
                                    : $"{strTotalSubstrate},[{ArrUseSubstrate[_i]}]{_subTotalTemp}";
                                _subTotalTemp = string.Empty;
                            }
                        }

                        using (var _cmd = _con.CreateCommand()) {
                            _cmd.CommandText =
                                $"""
                                 INSERT INTO Product_Reg_{StrProductName}
                                    (
                                    col_Order_Num,
                                    col_Product_Num,
                                    col_Product_Type,
                                    col_Product_Model,
                                    col_Quantity,
                                    col_Person,
                                    col_RegDate,
                                    col_Revision,
                                    col_Serial_First,
                                    col_Serial_Last,
                                    col_Serial_LastNum,
                                    col_Comment,
                                    col_Use_Substrate
                                    )
                                VALUES
                                    (
                                    @col_Order_Num,
                                    @col_Product_Num,
                                    @col_Product_Type,
                                    @col_Product_Model,
                                    @col_Quantity,
                                    @col_Person,
                                    @col_RegDate,
                                    @col_Revision,
                                    @col_Serial_First,
                                    @col_Serial_Last,
                                    @col_Serial_LastNum,
                                    @col_Comment,
                                    @col_Use_Substrate
                                    )
                                """;

                            _cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                            _cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                            _cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                            _cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                            _cmd.Parameters.Add("@col_Quantity", DbType.String).Value = IntQuantity;
                            _cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                            _cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                            _cmd.Parameters.Add("@col_Revision", DbType.String).Value = StrRevision;
                            _cmd.Parameters.Add("@col_Serial_First", DbType.String).Value = strSerialFirstNumber;
                            _cmd.Parameters.Add("@col_Serial_Last", DbType.String).Value = strSerialLastNumber;
                            _cmd.Parameters.Add("@col_Serial_LastNum", DbType.String).Value = intSerialLastNumber;
                            _cmd.Parameters.Add("@col_Use_Substrate", DbType.String).Value = strTotalSubstrate;
                            _cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                            _cmd.ExecuteNonQuery();
                        }
                    }
                    break;
            }
            return true;
        }
        private void RegisterCheck() {
            try {
                strSerial.Clear();

                if (!NumberCheck()) { return; }
                if (!QuantityCheck()) { return; }
                if (!SerialCheck()) { return; }
                // ラベル印刷
                switch (IntPrintType) {
                    case 1:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                        MessageBox.Show("シリアルラベルを印刷します。");
                        strSerialType = "Label";
                        if (!PrintBarcode(1)) { throw new Exception("キャンセルしました。"); }
                        break;
                    default:
                        break;
                }
                // バーコード印刷
                switch (IntPrintType) {
                    case 2:
                    case 3:
                        MessageBox.Show("バーコードラベルを印刷します。");
                        strSerialType = "Barcode";
                        if (!PrintBarcode(1)) { throw new Exception("キャンセルしました。"); }
                        break;
                    default:
                        break;
                }

                RegisterButton.Enabled = false;
                SerialPrintPostionNumericUpDown.Enabled = false;
                BarcodePrintPostionNumericUpDown.Enabled = false;

                // シリアル先頭と末尾を生成
                switch (IntPrintType) {
                    case 1:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                        strSerialType = "Label";
                        strSerialFirstNumber = GenerateCode(IntSerialFirstNumber);
                        strSerialLastNumber = GenerateCode(intSerialLastNumber);
                        break;
                    case 2:
                        strSerialType = "Barcode";
                        strSerialFirstNumber = GenerateCode(IntSerialFirstNumber);
                        strSerialLastNumber = GenerateCode(intSerialLastNumber);
                        break;
                    default:
                        break;
                }

                if (!Registration()) { throw new Exception("登録失敗しました。"); }

                MessageBox.Show("登録完了");

                switch (IntPrintType) {
                    case 0:
                    case 1:
                    case 2:
                    case 3:
                    case 4:
                    case 8:
                        Close();
                        break;
                    case 5:
                    case 6:
                    case 7:
                        foreach (Control _control in Controls) {
                            if (_control is DataGridView _dataGridView) {
                                _dataGridView.Enabled = false;
                            }
                            if (_control is CheckBox _checkBox) {
                                _checkBox.Enabled = false;
                            }
                        }
                        if (IntPrintType is 5 or 6) { SubstrateListPrintButton.Enabled = true; }
                        if (IntPrintType is 6 or 7) { CheckSheetPrintButton.Enabled = true; ; }
                        break;
                    default:
                        break;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool NumberCheck() {
            using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
            _con.Open();

            var _productModel = string.Empty;

            // 製番が新規かチェック
            using (var _cmd = _con.CreateCommand()) {
                _cmd.CommandText = $"SELECT * FROM Product_Reg_{StrProductName} WHERE col_Product_Num = @col_Product_Num ORDER BY _rowid_ ASC LIMIT 1";
                _cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;

                using var _dr = _cmd.ExecuteReader();
                while (_dr.Read()) {
                    _productModel = $"{_dr["col_Product_Model"]}";
                }
            }

            if (_productModel != string.Empty) {
                if (_productModel == StrProductModel) {
                    Activate();
                    var _result = MessageBox.Show($"製番[{StrProductNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                    if (_result == DialogResult.No) {
                        return false;
                    }
                }
                else {
                    Activate();
                    MessageBox.Show($"[{StrProductNumber}]は[{_productModel}]として登録があります。確認してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            _productModel = string.Empty;

            // 注文番号が新規かチェック
            using (var _cmd = _con.CreateCommand()) {
                _cmd.CommandText = $"SELECT * FROM Product_Reg_{StrProductName} WHERE col_Order_Num = @col_Order_Num ORDER BY _rowid_ ASC LIMIT 1";
                _cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;

                using var _dr = _cmd.ExecuteReader();
                while (_dr.Read()) {
                    _productModel = $"{_dr["col_Product_Model"]}";
                }
            }

            if (_productModel != string.Empty) {
                if (_productModel == StrProductModel) {
                    Activate();
                    var _result = MessageBox.Show($"注文番号[{StrOrderNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                    if (_result == DialogResult.No) {
                        return false;
                    }
                }
                else {
                    Activate();
                    MessageBox.Show($"[{StrOrderNumber}]は[{_productModel}]として登録があります。確認してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            return true;
        }
        private bool QuantityCheck() {
            try {
                switch (IntRegType) {
                    case 0:
                    case 1:
                        break;
                    case 2:
                    case 3:
                        if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (var _i = 0; _i <= ArrUseSubstrate.GetUpperBound(0); _i++) {

                            var _objCbx = Controls[checkBoxNames[_i]] as CheckBox ?? throw new Exception("objCbxがnullです。");
                            _objCbx.Enabled = true;
                            _objCbx.Checked = true;

                            var _objDgv = Controls[dataGridViewNames[_i]] as DataGridView ?? throw new Exception("objDgvがnullです。");
                            _objDgv.Columns[2].ReadOnly = false;
                            _objDgv.Columns[3].ReadOnly = false;

                            if (_objCbx.Checked) {
                                var _intQuantityCheck = IntQuantity;
                                var _dgvRowCnt = _objDgv.Rows.Count;

                                for (var _j = 0; _j < _dgvRowCnt; _j++) {
                                    var _boolCbx = _objDgv.Rows[_j].Cells[3].Value != null && (bool)_objDgv.Rows[_j].Cells[3].Value;
                                    if (_boolCbx) {
                                        var _stockValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[1].Value.ToString());
                                        if (_objDgv.Rows[_j].Cells[2].Value == null) { throw new Exception("使用数が入力されていません。"); }
                                        var _useValue = Convert.ToInt32(_objDgv.Rows[_j].Cells[2].Value.ToString());

                                        if (_stockValue < _useValue) {
                                            throw new Exception("在庫より多い数量が入力されています。");
                                        }
                                        _intQuantityCheck -= _useValue;
                                    }
                                }

                                if (_intQuantityCheck != 0) {
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
                switch (IntPrintType) {
                    case 1:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                        for (var _i = 0; _i < IntQuantity; _i++) {
                            strSerialType = "Label";
                            strSerial.Add(GenerateCode(IntSerialFirstNumber + _i));
                        }
                        break;
                    case 2:
                        for (var _i = 0; _i < IntQuantity; _i++) {
                            strSerialType = "Barcode";
                            strSerial.Add(GenerateCode(IntSerialFirstNumber + _i));
                        }
                        break;
                    default:
                        break;
                }

                var _strSQLSerial = string.Join("','", strSerial);

                List<string> _strSerialDuplication = [];
                using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                    _con.Open();

                    using var _cmd = _con.CreateCommand();
                    _cmd.CommandText = $"SELECT col_Serial FROM Serial_{StrProductName} WHERE col_Serial IN (@col_Serial)";
                    _cmd.Parameters.Add("@col_Serial", DbType.String).Value = _strSQLSerial;

                    using var _dr = _cmd.ExecuteReader();
                    while (_dr.Read()) {
                        _strSerialDuplication.Add($"{_dr["col_Serial"]}");
                    }
                }

                if (_strSerialDuplication.Count > 0) {
                    var _strSQLDuplication = string.Join($"{Environment.NewLine}", _strSerialDuplication);
                    throw new Exception($"{_strSQLDuplication}{Environment.NewLine}は既に使用されているシリアルです。");
                }

                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // 印刷処理
        private bool PrintBarcode(int PrintFlg) {
            // PrintDocumentオブジェクトの作成
            using System.Drawing.Printing.PrintDocument _pd = new();

            // PrintPageイベントハンドラの追加
            _pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

            LabelProNumLabelsToPrint = IntQuantity;
            fontUnderbar = false;

            switch (PrintFlg) {
                case 1:
                    ProductRegistration2PrintDialog.Document = _pd;
                    var _r = ProductRegistration2PrintDialog.ShowDialog();

                    if (_r == DialogResult.OK) {
                        ProductRegistration2PrintDialog.Document.Print();

                        if (IntPageCnt >= 2) {
                            MessageBox.Show($"{IntPageCnt}枚印刷されます。2枚目以降は1行目から印刷されます。");
                        }
                    }
                    else {
                        return false;
                    }
                    break;
                case 2:
                    ProductRegistration2PrintPreviewDialog.Document = _pd;
                    ProductRegistration2PrintPreviewDialog.ShowDialog();
                    break;
            }
            return true;
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                Point _headerPos = new(0, 0);
                var _headerString = string.Empty;
                Font _headerFooterFont = new("Arial", 6);
                var _intNumLabels = 0;
                var _intCountNumLabels = 0;

                var _maxX = 0;
                var _maxY = 0;
                float _sizeX = 0;
                float _sizeY = 0;
                double _offsetX = 0;
                double _offsetY = 0;
                double _intervalX = 0;
                double _intervalY = 0;
                var _startLine = 0;
                var _labelProPageNum = 0;

                switch (strSerialType) {
                    case "Label":
                        if (SettingsLabelPro == null) { throw new Exception("SettingsLabelProがnullです。"); }
                        _maxX = SettingsLabelPro._labelProPageSettings.NumLabelsX;
                        _maxY = SettingsLabelPro._labelProPageSettings.NumLabelsY;
                        _sizeX = (float)SettingsLabelPro._labelProPageSettings.SizeX;
                        _sizeY = (float)SettingsLabelPro._labelProPageSettings.SizeY;
                        _offsetX = SettingsLabelPro._labelProPageSettings.OffsetX;
                        _offsetY = SettingsLabelPro._labelProPageSettings.OffsetY;
                        _intervalX = SettingsLabelPro._labelProPageSettings.IntervalX;
                        _intervalY = SettingsLabelPro._labelProPageSettings.IntervalY;
                        _headerPos = SettingsLabelPro._labelProPageSettings.HeaderPos;
                        _headerString = ConvertHeaderFooterString(SettingsLabelPro._labelProPageSettings.HeaderString);
                        _headerFooterFont = SettingsLabelPro._labelProPageSettings.HeaderFooterFont;
                        _intNumLabels = SettingsLabelPro._labelProLabelSettings.NumLabels;
                        _intCountNumLabels = SettingsLabelPro._labelProLabelSettings.NumLabels;
                        _startLine = (int)SerialPrintPostionNumericUpDown.Value - 1;
                        break;
                    case "Barcode":
                        if (SettingsBarcodePro == null) { throw new Exception("SettingsBarcodeProがnullです。"); }
                        _maxX = SettingsBarcodePro.BarcodeProPageSettings.NumLabelsX;
                        _maxY = SettingsBarcodePro.BarcodeProPageSettings.NumLabelsY;
                        _sizeX = (float)SettingsBarcodePro.BarcodeProPageSettings.SizeX;
                        _sizeY = (float)SettingsBarcodePro.BarcodeProPageSettings.SizeY;
                        _offsetX = SettingsBarcodePro.BarcodeProPageSettings.OffsetX;
                        _offsetY = SettingsBarcodePro.BarcodeProPageSettings.OffsetY;
                        _intervalX = SettingsBarcodePro.BarcodeProPageSettings.IntervalX;
                        _intervalY = SettingsBarcodePro.BarcodeProPageSettings.IntervalY;
                        _headerPos = SettingsBarcodePro.BarcodeProPageSettings.HeaderPos;
                        _headerString = ConvertHeaderFooterString(SettingsBarcodePro.BarcodeProPageSettings.HeaderString);
                        _headerFooterFont = SettingsBarcodePro.BarcodeProPageSettings.HeaderFooterFont;
                        _intNumLabels = SettingsBarcodePro.BarcodeProLabelSettings.NumLabels;
                        _intCountNumLabels = SettingsBarcodePro.BarcodeProLabelSettings.NumLabels;
                        _startLine = (int)BarcodePrintPostionNumericUpDown.Value - 1;
                        break;
                    default:
                        break;
                }

                Point _offset;
                if (!ProductRegistration2PrintDocument.PrintController.IsPreview) {
                    _offsetX -= e.PageSettings.HardMarginX * 0.254;
                    _offsetY -= e.PageSettings.HardMarginY * 0.254;
                    _offset = _labelProPageNum == 0
                        ? new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (_startLine * (_intervalY + _sizeY))))
                        : new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (_intervalY + _sizeY))));
                }
                else {
                    _offset = new Point(0, 0);
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;

                _headerPos.Offset(_offset);
                e.Graphics.DrawString(_headerString, _headerFooterFont, Brushes.Black, _headerPos);

                var _pageNum = 0;
                _offset = _pageNum == 0
                    ? new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (_startLine * (_intervalY + _sizeY))))
                    : new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (_intervalY + _sizeY))));

                if (_labelProPageNum == 0) { LabelProNSerial = IntSerialFirstNumber; }
                if (_labelProPageNum >= 1) { _startLine = 0; }

                var _y = 0;
                for (_y = _startLine; _y < _maxY; _y++) {
                    var _x = 0;
                    for (_x = 0; _x < _maxX; _x++) {
                        var _s = GenerateCode(LabelProNSerial);
                        var _posX = (float)(_offsetX + (_x * (_intervalX + _sizeX)));
                        var _posY = (float)(_offsetY + (_y * (_intervalY + _sizeY)));
                        e.Graphics.DrawImage(MakeLabelImage(_s, (int)e.Graphics.DpiX, 1), _posX, _posY, _sizeX, _sizeY);

                        LabelProNSerial++;
                        LabelProNumLabelsToPrint--;

                        if (LabelProNumLabelsToPrint <= 0) {
                            _intCountNumLabels--;
                            if (_intCountNumLabels <= 0) {
                                e.HasMorePages = false;
                                _labelProPageNum = 0;
                                var _txtNumPublish = 0;
                                LabelProNumLabelsToPrint = _txtNumPublish;
                                return;
                            }
                            else {
                                LabelProNumLabelsToPrint += _x + 1;
                                break;
                            }
                        }

                        if (_x >= _maxX - 1) {
                            _intCountNumLabels--;
                            if (_intCountNumLabels <= 0) {
                                _intCountNumLabels = _intNumLabels;
                            }
                            else if (_intCountNumLabels > 0) {
                                LabelProNumLabelsToPrint += _x + 1;
                                break;
                            }
                        }
                    }
                }

                if (LabelProNumLabelsToPrint > 0) {
                    _labelProPageNum++;
                    IntPageCnt++;
                    e.HasMorePages = true;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        private string ConvertHeaderFooterString(string s) {
            s = s.Replace("%P", StrProductName)
                 .Replace("%T", StrProductModel)
                 .Replace("%D", DateTime.Today.ToShortDateString())
                 .Replace("%M", StrProductNumber)
                 .Replace("%O", StrOrderNumber)
                 .Replace("%N", IntQuantity.ToString())
                 .Replace("%U", "");
            return s;
        }
        private string GenerateCode(int serialCode) {
            var _monthCode = DateTime.Parse(StrRegDate).ToString("MM");

            _monthCode = _monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => _monthCode
            };

            var _outputCode = strSerialType switch {
                "Label" => SettingsLabelPro._labelProLabelSettings.Format,
                "Barcode" => SettingsBarcodePro.BarcodeProLabelSettings.Format,
                _ => string.Empty
            };

            var _serialCode = Convert.ToInt32(serialCode).ToString($"D{IntSerialDigit}");
            _outputCode = _outputCode.Replace("%Y", DateTime.Parse(StrRegDate).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(StrRegDate).ToString("MM"))
                                    .Replace("%T", StrInitial)
                                    .Replace("%R", StrRevision)
                                    .Replace("%M", string.IsNullOrEmpty(_monthCode) ? string.Empty : _monthCode[^1..])
                                    .Replace("%S", _serialCode);
            return _outputCode;
        }
        private Bitmap MakeLabelImage(string text, int resolution, int magnitude) {
            Bitmap _labelImage = new(1, 1);
            Graphics _g;
            SizeF _stringSize;
            decimal _sizeX;
            decimal _sizeY;
            decimal _fontSize;
            float _stringPosX;
            float _stringPosY;
            Font _fnt;
            switch (strSerialType) {
                case "Label":
                    if (SettingsLabelPro == null) { throw new Exception("SettingsLabelProがnull"); }
                    _sizeX = (decimal)SettingsLabelPro._labelProPageSettings.SizeX / 25.4M * resolution * magnitude;
                    _sizeY = (decimal)SettingsLabelPro._labelProPageSettings.SizeY / 25.4M * resolution * magnitude;
                    _fontSize = (decimal)SettingsLabelPro._labelProLabelSettings.Font.SizeInPoints / 72.0M * resolution * magnitude;
                    _stringPosY = (int)((decimal)SettingsLabelPro._labelProLabelSettings.StringPosY / 25.4M * resolution * magnitude);
                    var _style = fontUnderbar ? FontStyle.Underline : FontStyle.Regular;
                    _fnt = new Font(SettingsLabelPro._labelProLabelSettings.Font.Name, (float)_fontSize, _style);

                    _labelImage = new((int)_sizeX, (int)_sizeY);
                    _g = Graphics.FromImage(_labelImage);

                    _stringSize = _g.MeasureString(text, _fnt);

                    _stringPosX = (int)((_labelImage.Width / 2) - (_stringSize.Width / 2));

                    _g.DrawString(text, _fnt, Brushes.Black, _stringPosX, _stringPosY);

                    _g.Dispose();
                    break;
                case "Barcode":
                    if (SettingsBarcodePro == null) { throw new Exception("SettingsBarcodeProがnull"); }

                    _sizeX = (decimal)SettingsBarcodePro.BarcodeProPageSettings.SizeX / 25.4M * resolution * magnitude;
                    _sizeY = (decimal)SettingsBarcodePro.BarcodeProPageSettings.SizeY / 25.4M * resolution * magnitude;
                    _fontSize = (decimal)SettingsBarcodePro.BarcodeProLabelSettings.Font.SizeInPoints / 72.0M * resolution * magnitude;
                    _fnt = new(SettingsBarcodePro.BarcodeProLabelSettings.Font.Name, (float)_fontSize);

                    _labelImage = new((int)_sizeX, (int)_sizeY);
                    _g = Graphics.FromImage(_labelImage);

                    int _barWeight;
                    _barWeight = resolution == DisplayResolution ? 1 : (int)(1 * resolution / DisplayResolution / DisplayMagnitude);

                    using (var _img = Code128Rendering.MakeBarcodeImage(text, _barWeight, true)) {
                        var _imageWidth = (decimal)(_img.Width * SettingsBarcodePro.BarcodeProLabelSettings.BarcodeMagnitude);

                        if (_imageWidth > _labelImage.Width) { MessageBox.Show($"バーコードの幅がラベル幅を超えています{_imageWidth}>{_labelImage.Width}"); }

                        _stringSize = _g.MeasureString(text, _fnt);

                        _stringPosX = SettingsBarcodePro.BarcodeProLabelSettings.AlignStringCenter
                            ? (float)((_labelImage.Width / 2) - (_stringSize.Width / 2))
                            : (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4M * resolution * magnitude);

                        _stringPosY = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosY / 25.4M * resolution * magnitude);

                        var _barCodePosX = SettingsBarcodePro.BarcodeProLabelSettings.AlignBarcodeCenter
                            ? (float)((_labelImage.Width / 2) - (_imageWidth / 2))
                            : (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4M * resolution * magnitude);
                        float _barCodePosY = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.BarcodePosY / 25.4M * resolution * magnitude);

                        _g.DrawString(text, _fnt, Brushes.Black, _stringPosX, _stringPosY);

                        var _barcodeHeight = (int)(SettingsBarcodePro.BarcodeProLabelSettings.BarcodeHeight / 25.4F * resolution * magnitude);
                        _g.DrawImage(_img, _barCodePosX, _barCodePosY, (float)_imageWidth, _barcodeHeight);

                        _g.Dispose();
                        _img.Dispose();
                    }

                    break;
                default:
                    break;
            }
            return _labelImage;
        }
        // チェックボックスイベント
        private void CheckBox_CheckedChanged(object sender, EventArgs e) {
            var _checkBox = (CheckBox)sender;
            DataGridView _dataGridView = new();

            switch (_checkBox.Name) {
                case "Substrate1CheckBox":
                    _dataGridView = Substrate1DataGridView;
                    break;
                case "Substrate2CheckBox":
                    _dataGridView = Substrate2DataGridView;
                    break;
                case "Substrate3CheckBox":
                    _dataGridView = Substrate3DataGridView;
                    break;
                case "Substrate4CheckBox":
                    _dataGridView = Substrate4DataGridView;
                    break;
                case "Substrate5CheckBox":
                    _dataGridView = Substrate5DataGridView;
                    break;
                case "Substrate6CheckBox":
                    _dataGridView = Substrate6DataGridView;
                    break;
                case "Substrate7CheckBox":
                    _dataGridView = Substrate7DataGridView;
                    break;
                case "Substrate8CheckBox":
                    _dataGridView = Substrate8DataGridView;
                    break;
                case "Substrate9CheckBox":
                    _dataGridView = Substrate9DataGridView;
                    break;
                case "Substrate10CheckBox":
                    _dataGridView = Substrate10DataGridView;
                    break;
                default:
                    break;
            }

            _dataGridView.Enabled = _checkBox.Checked;
            _checkBox.ForeColor = _checkBox.Checked ? Color.Black : Color.Red;

            if (!_checkBox.Checked) {
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

                //using FileStream _fileStream = new($@"{Environment.CurrentDirectory}./config/Excel/ConfigCheckSheet.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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
                //    using FileStream _fileStream2 = new($@"{_filePath[_i]}", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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
                //    Excel.Workbook _xlBook = _xlBooks.Open($@"{System.Environment.CurrentDirectory}./config/Excel/temporarily_{_i}.xlsx");

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
                var _sheetName = string.Empty;
                var _productName = string.Empty;
                var _productModel = string.Empty;

                var _productNameRange = string.Empty;
                var _productNumberRange = string.Empty;
                var _orderNumberRange = string.Empty;
                var _regDateRange = string.Empty;
                var _productModelRange = string.Empty;
                var _quantityRange = string.Empty;
                var _serialFirstRange = string.Empty;
                var _serialLastRange = string.Empty;
                var _commentRange = string.Empty;

                using FileStream _fileStream = new($"{Environment.CurrentDirectory}./config/Excel/ConfigList.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using XLWorkbook _workBook = new(_fileStream);
                var _workSheetMain = _workBook.Worksheet("Sheet1");

                var _findRow = 0;
                // セル検索
                foreach (var _cell in _workSheetMain.Search(StrProductModel)) {
                    _findRow = _cell.Address.RowNumber;
                }

                if (_findRow == 0) {
                    throw new Exception($"Configに品目番号:{StrProductModel}が見つかりません。");
                }

                // ワークシートのセルから値を取得
                _sheetName = StrProductModel;
                _productName = _workSheetMain.Cell(_findRow, 2).Value.ToString();
                _productNameRange = _workSheetMain.Cell(_findRow, 3).Value.ToString();
                _productNumberRange = _workSheetMain.Cell(_findRow, 4).Value.ToString();
                _orderNumberRange = _workSheetMain.Cell(_findRow, 5).Value.ToString();
                _regDateRange = _workSheetMain.Cell(_findRow, 6).Value.ToString();
                _productModel = _workSheetMain.Cell(_findRow, 7).Value.ToString();
                _productModelRange = _workSheetMain.Cell(_findRow, 8).Value.ToString();
                _quantityRange = _workSheetMain.Cell(_findRow, 9).Value.ToString();
                _serialFirstRange = _workSheetMain.Cell(_findRow, 10).Value.ToString();
                _serialLastRange = _workSheetMain.Cell(_findRow, 11).Value.ToString();
                _commentRange = _workSheetMain.Cell(_findRow, 12).Value.ToString();

                var _workSheetTemp = _workBook.Worksheet(_sheetName);
                _workSheetTemp.Cell(_productNameRange).Value = _productName;
                _workSheetTemp.Cell(_productNumberRange).Value = StrProductNumber;
                _workSheetTemp.Cell(_orderNumberRange).Value = StrOrderNumber;
                _workSheetTemp.Cell(_regDateRange).Value = StrRegDate;
                _workSheetTemp.Cell(_productModelRange).Value = _productModel;
                _workSheetTemp.Cell(_quantityRange).Value = IntQuantity;
                _workSheetTemp.Cell(_serialFirstRange).Value = strSerialFirstNumber;
                _workSheetTemp.Cell(_serialLastRange).Value = strSerialLastNumber;
                _workSheetTemp.Cell(_commentRange).Value = StrComment;

                var _i = 0;
                var _findColumn = 0;
                for (_i = 0; _i <= listUsedSubstrate.Count - 1; _i++) {

                    var _searchRange = _workSheetMain.Range(_findRow, 1, _findRow, 28);
                    var _searchValue = $"{listUsedSubstrate[_i]}";
                    var _foundCell = _searchRange.CellsUsed(c => c.Value.ToString() == _searchValue).FirstOrDefault();

                    if (_foundCell != null) {
                        // セルが見つかった場合の処理
                        var _foundRow = _foundCell.Address.RowNumber;
                        var _foundColumn = _foundCell.Address.ColumnNumber;
                    }

                    foreach (var _cell in _workSheetMain.Search(listUsedSubstrate[_i])) {
                        if (_cell.Address.RowNumber == _findRow) {
                            _findColumn = _cell.Address.ColumnNumber;
                            break;
                        }
                    }

                    if (_findColumn == 0) {
                        throw new Exception($"{listUsedSubstrate[_i]}が見つかりません。");
                    }

                    var _mainCellValue = _workSheetMain.Cell(_findRow, _findColumn + 1).Value.ToString();
                    var _tempCellValue = _workSheetTemp.Cell(_mainCellValue).Value.ToString();

                    if (_mainCellValue != string.Empty) {
                        if (_tempCellValue == string.Empty) {
                            _workSheetTemp.Cell(_mainCellValue).Value = $"{listUsedProductNumber[_i]}({listUsedQuantity[_i]})";
                        }
                        else {
                            _workSheetTemp.Cell(_mainCellValue).Value += $"    {listUsedProductNumber[_i]}({listUsedQuantity[_i]})";
                        }
                    }
                }

                //引数に保存先パスを指定
                _workBook.SaveAs($"{Environment.CurrentDirectory}./config/Excel/temporarily.xlsx");

                // 印刷
                Excel.Application _xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                var _xlBooks = _xlApp.Workbooks;
                var _xlBook = _xlBooks.Open($"{Environment.CurrentDirectory}./config/Excel/temporarily.xlsx");

                // ワークシート選択
                var _xlSheets = _xlBook.Sheets;
                Excel.Worksheet _xlSheet = _xlSheets[_sheetName];

                // ワークシート印刷
                _xlSheet.PrintOut(Preview: true);

                // ワークブックを閉じてExcelを終了
                _xlBook.Close(false);
                _xlApp.Quit();

                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlSheet);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlSheets);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlBook);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlBooks);
                _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlApp);

            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProductRegistration2Window_Load(object sender, EventArgs e) { LoadEvents(); }
        private void ProductRegistration2Window_FormClosing(object sender, FormClosingEventArgs e) { ClosingEvents(); }
        private void RegisterButton_Click(object sender, EventArgs e) { RegisterCheck(); }
        private void CloseButton_Click(object sender, EventArgs e) { ClosingEvents(); }
        private void SubstrateCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBox_CheckedChanged(sender, e); }
        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e) { ClosingEvents(); }
        private void シリアルラベル印刷ToolStripMenuItem_Click(object sender, EventArgs e) {
            strSerialType = "Label";
            PrintBarcode(1);
        }
        private void シリアルラベル印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            strSerialType = "Label";
            PrintBarcode(2);
        }
        private void バーコード印刷ToolStripMenuItem_Click(object sender, EventArgs e) {
            strSerialType = "Barcode";
            PrintBarcode(1);
        }
        private void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            strSerialType = "Barcode";
            PrintBarcode(2);
        }
        private void シリアルラベル印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductPrintSetting _ls = new();
            _ls.ShowDialog(this);
        }
        private void バーコード印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductBarcodePrintSettingsWindow _ls = new();
            _ls.ShowDialog(this);
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            var message = string.Join(Environment.NewLine,
                $"StrProductName\t\t[{StrProductName}]",
                $"StrProductModel\t\t[{StrProductModel}]",
                $"StrStockName\t\t[{StrStockName}]",
                $"StrProductType\t\t[{StrProductType}]",
                $"StrOrderNumber\t\t[{StrOrderNumber}]",
                $"StrProductNumber\t\t[{StrProductNumber}]",
                $"StrRevision\t\t[{StrRevision}]",
                $"StrRegDate\t\t[{StrRegDate}]",
                $"StrPerson\t\t\t[{StrPerson}]",
                $"IntQuantity\t\t[{IntQuantity}]",
                $"IntSerialFirstNumber\t[{IntSerialFirstNumber}]",
                $"IntSerialLastNumber\t[{intSerialLastNumber}]",
                $"StrInitial\t\t\t[{StrInitial}]",
                $"IntRegType\t\t[{IntRegType}]",
                $"IntPrintType\t\t[{IntPrintType}]",
                $"IntSerialDigit\t\t[{IntSerialDigit}]"
            );
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) { ListPrint(); }
        private void CheckSheetPrintButton_Click(object sender, EventArgs e) { CheckSheetPrint(); }
        private void ProductRegistration2PrintPreviewDialog_Load(object sender, EventArgs e) {
            var _tool = (ToolStrip)ProductRegistration2PrintPreviewDialog.Controls[1];
            _tool.Items[0].Visible = false;
        }
    }
}
