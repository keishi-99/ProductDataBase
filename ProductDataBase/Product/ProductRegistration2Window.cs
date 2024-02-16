using ClosedXML.Excel;
using GenCode128;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using Excel = Microsoft.Office.Interop.Excel;

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
        public string[] ArrUseSubstrate = Array.Empty<string>();
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

        public Bitmap LabelProBmp { get; } = new(1, 1);
        public int LabelProPageNum { get; set; }
        public int LabelProNSerial { get; set; }
        public int LabelProNLabel { get; set; }
        public int LabelProNumLabelsToPrint { get; set; }

        public decimal DisplayResolution { get; } = 96.0m;
        public int DisplayMagnitude { get; } = 3;
        public int IntPageCnt { get; set; } = 1;

        public string StrProness1 { get; } = string.Empty;
        public string StrProness2 { get; } = string.Empty;
        public string StrProness3 { get; } = string.Empty;
        public int StrProness4 { get; }
        public string StrProness5 { get; } = string.Empty;

        private string StrSerialType = string.Empty;
        private string StrSerialFirstNumber = string.Empty;
        private string StrSerialLastNumber = string.Empty;
        private string StrTotalSubstrate = string.Empty;
        private int IntSerialLastNumber;
        private bool FontUnderbar = false;

        readonly List<string> StrSerial = new();
        readonly List<string> CheckBoxNames = new() {
                        "Substrate1CheckBox", "Substrate2CheckBox", "Substrate3CheckBox", "Substrate4CheckBox","Substrate5CheckBox",
                        "Substrate6CheckBox", "Substrate7CheckBox", "Substrate8CheckBox", "Substrate9CheckBox","Substrate10CheckBox"
                        };
        readonly List<string> DataGridViewNames = new() {
                        "DataGridView1", "DataGridView2", "DataGridView3", "DataGridView4","DataGridView5",
                        "DataGridView6", "DataGridView7", "DataGridView8", "DataGridView9","DataGridView10"
                        };
        readonly List<string> ListUsedSubstrate = new();
        readonly List<string> ListUsedProductNumber = new();
        readonly List<int> ListUsedQuantity = new();

        public ProductRegistration2Window() => InitializeComponent();
        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(StrFontName, IntFontSize);

                RegisterButton.Enabled = true;
                ArrUseSubstrate = StrUseSubstrate.Split(",");

                IntSerialLastNumber = IntSerialFirstNumber + IntQuantity - 1;

                bool _quantityFlg = false;
                string _strQuantity = string.Empty;
                switch (IntRegType) {
                    case 2:
                        for (int _i = 0; _i <= ArrUseSubstrate.GetUpperBound(0); _i++) {
                            CheckBox? _objCbx = Controls[CheckBoxNames[_i]] as CheckBox;

                            if (_objCbx != null) {
                                _objCbx.Enabled = true;
                                _objCbx.Checked = true;
                            }

                            DataGridView? _objDgv = Controls[DataGridViewNames[_i]] as DataGridView;
                            if (_objDgv != null) {
                                _objDgv.Columns[2].ReadOnly = false;
                                _objDgv.Columns[3].ReadOnly = false;
                            }

                            using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
                            _con.Open();

                            using SQLiteCommand _cmd = _con.CreateCommand();
                            // 使用基板表示
                            _cmd.CommandText = $"SELECT col_Substrate_Name FROM 'Stock_{StrStockName}' WHERE col_Substrate_Model = '{ArrUseSubstrate[_i]}'";
                            using (SQLiteDataReader _dr = _cmd.ExecuteReader()) {
                                if (_dr.Read()) {
                                    if (_objCbx != null) {
                                        string _substrateName = $"{_dr["col_Substrate_Name"]}";
                                        _objCbx.Text = $"{_substrateName} - {ArrUseSubstrate[_i]}";
                                    }
                                }
                            }

                            // 在庫テーブルからデータ取得
                            _cmd.CommandText = $"SELECT col_Substrate_num, col_Stock, col_Substrate_Name FROM 'Stock_{StrStockName}' WHERE col_flg = 1 And col_Substrate_Model = '{ArrUseSubstrate[_i]}' ORDER BY _rowid_ ASC";
                            using (SQLiteDataReader _dr = _cmd.ExecuteReader()) {
                                while (_dr.Read()) {
                                    string _strSubstrateNumber = $"{_dr["col_Substrate_num"]}";
                                    int _intStock = Convert.ToInt32(_dr["col_Stock"]);
                                    _objDgv?.Rows.Add(_strSubstrateNumber, _intStock);

                                    int _intQuantity = IntQuantity;
                                    if (_intQuantity >= _intStock) {
                                        _intQuantity -= _intStock;
                                        if (_objDgv != null) {
                                            _objDgv.Rows[^1].Cells[2].Value = _intStock;
                                            _objDgv.Rows[^1].Cells[3].Value = true;
                                        }
                                    }
                                    else {
                                        if (_intQuantity == 0)
                                            if (_objDgv != null) {
                                                _objDgv.Rows[^1].Cells[2].Value = null;
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
                                        string _substrateName = $"{_dr["col_Substrate_Name"]}";
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
                        for (int _i = 0; _i <= ArrUseSubstrate.GetUpperBound(0); _i++) {
                            CheckBox? _objCbx = Controls[CheckBoxNames[_i]] as CheckBox;

                            if (_objCbx != null) {
                                _objCbx.Enabled = true;
                                _objCbx.Checked = true;
                            }

                            DataGridView? _objDgv = Controls[DataGridViewNames[_i]] as DataGridView;
                            if (_objDgv != null) {
                                _objDgv.Columns[2].ReadOnly = false;
                                _objDgv.Columns[3].ReadOnly = false;
                            }

                            using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
                            _con.Open();
                            using SQLiteCommand _cmd = _con.CreateCommand();
                            _cmd.CommandText = $"SELECT * FROM 'Stock_{StrStockName}' WHERE col_Substrate_Model = '{ArrUseSubstrate[_i]}'";
                            using (SQLiteDataReader _dr = _cmd.ExecuteReader()) {
                                while (_dr.Read()) {
                                    if (_objCbx != null) { _objCbx.Text = $"{_dr["col_Substrate_Name"]} - {ArrUseSubstrate[_i]}"; }
                                }
                            }

                            _cmd.CommandText = $"SELECT * FROM 'Stock_{StrStockName}' WHERE col_flg = 1 And col_Substrate_Model = '{ArrUseSubstrate[_i]}' ORDER BY _rowid_ ASC";
                            using (SQLiteDataReader _dr = _cmd.ExecuteReader()) {
                                while (_dr.Read()) {
                                    string _strSubstrateName = string.Empty;
                                    _strSubstrateName = $"{_dr["col_Substrate_Name"]}";

                                    string _strSubstrateNumber = $"{_dr["col_Substrate_Num"]}";
                                    int _intStock = Convert.ToInt32(_dr["col_Stock"]);
                                    _objDgv?.Rows.Add(_strSubstrateNumber, _intStock);

                                    int _j = 0;
                                    string _strOrderNumber = $"{_dr["col_Order_Num"]}";
                                    if (_strOrderNumber == StrOrderNumber) {
                                        if (_objDgv != null) {
                                            int _intQuantity = IntQuantity;
                                            _objDgv.Rows[_j].Cells[2].Value = _intQuantity;
                                            _objDgv.Rows[_j].Cells[3].Value = true;
                                            // 必要数量分割り当てられたかチェック
                                            if (_intQuantity > Convert.ToInt32(_objDgv.Rows[_j].Cells[1].Value)) {
                                                _quantityFlg = false;
                                                _strQuantity = $"{_strQuantity}[{_strSubstrateName}]{Environment.NewLine}";
                                            }
                                        }
                                    }
                                    _j++;
                                }
                            }
                        }

                        if (!_quantityFlg) {
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

                switch (IntPrintType) {
                    case 5:
                        SubstrateListPrintButton.Enabled = true;
                        break;
                    case 6:
                        SubstrateListPrintButton.Enabled = true;
                        CheckSheetPrintButton.Enabled = true;
                        break;
                    case 7:
                        CheckSheetPrintButton.Enabled = true;
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

                foreach (var _b in StrSerial) {
                    using SQLiteCommand _cmd = _con.CreateCommand();
                    _cmd.CommandText =
                        $"INSERT INTO 'Serial_{StrProductName}'" +
                        $"(col_Serial," +
                        $"col_Order_Num," +
                        $"col_Product_Num," +
                        $"col_Product_Type," +
                        $"col_Product_Model," +
                        $"col_RegDate)" +
                        $"VALUES " +
                        $"(@col_Serial," +
                        $"@col_Order_Num," +
                        $"@col_Product_Num," +
                        $"@col_Product_Type," +
                        $"@col_Product_Model," +
                        $"@col_RegDate)";

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

                        using SQLiteCommand _cmd = _con.CreateCommand();
                        _cmd.CommandText =
                            $"INSERT INTO 'Product_Reg_{StrProductName}'" +
                            $"(col_Order_Num," +
                            $"col_Product_Num," +
                            $"col_Product_Type," +
                            $"col_Product_Model," +
                            $"col_Quantity," +
                            $"col_Person," +
                            $"col_RegDate," +
                            $"col_Revision," +
                            $"col_Comment)" +
                            $"VALUES " +
                            $"(@col_Order_Num," +
                            $"@col_Product_Num," +
                            $"@col_Product_Type," +
                            $"@col_Product_Model," +
                            $"@col_Quantity," +
                            $"@col_Person," +
                            $"@col_RegDate," +
                            $"@col_Revision," +
                            $"@col_Comment)";

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

                        using SQLiteCommand _cmd = _con.CreateCommand();
                        _cmd.CommandText =
                            $"INSERT INTO 'Product_Reg_{StrProductName}'" +
                            $"(col_Order_Num," +
                            $"col_Product_Num," +
                            $"col_Product_Type," +
                            $"col_Product_Model," +
                            $"col_Quantity," +
                            $"col_Person," +
                            $"col_RegDate," +
                            $"col_Revision," +
                            $"col_Serial_First," +
                            $"col_Serial_Last," +
                            $"col_Serial_LastNum," +
                            $"col_Comment)" +
                            $"VALUES " +
                            $"(@col_Order_Num," +
                            $"@col_Product_Num," +
                            $"@col_Product_Type," +
                            $"@col_Product_Model," +
                            $"@col_Quantity," +
                            $"@col_Person," +
                            $"@col_RegDate," +
                            $"@col_Revision," +
                            $"@col_Serial_First," +
                            $"@col_Serial_Last," +
                            $"@col_Serial_LastNum," +
                            $"@col_Comment)";

                        _cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                        _cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                        _cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                        _cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                        _cmd.Parameters.Add("@col_Quantity", DbType.String).Value = IntQuantity;
                        _cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                        _cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                        _cmd.Parameters.Add("@col_Revision", DbType.String).Value = StrRevision;
                        _cmd.Parameters.Add("@col_Serial_First", DbType.String).Value = StrSerialFirstNumber;
                        _cmd.Parameters.Add("@col_Serial_Last", DbType.String).Value = StrSerialLastNumber;
                        _cmd.Parameters.Add("@col_Serial_LastNum", DbType.String).Value = IntSerialLastNumber;
                        _cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                        _cmd.ExecuteNonQuery();
                    }
                    break;

                case 2:
                case 3:
                    using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                        _con.Open();
                        if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateがnullです。"); }
                        for (int _i = 0; _i <= ArrUseSubstrate.Length; _i++) {

                            CheckBox? _objCbx = Controls[CheckBoxNames[_i]] as CheckBox ?? throw new Exception("objCbxがnullです。");

                            if (_objCbx.Checked) {
                                DataGridView? _objDgv = Controls[DataGridViewNames[_i]] as DataGridView ?? throw new Exception("objCbxがnullです。");
                                int _dgvRowCnt = _objDgv.Rows.Count;

                                for (int _j = 0; _j <= _dgvRowCnt - 1; _j++) {
                                    if (Convert.ToBoolean(_objDgv.Rows[_j].Cells[3].Value)) {

                                        string _substrateName = string.Empty;
                                        string _substrateModel = string.Empty;
                                        using (SQLiteCommand _cmd = _con.CreateCommand()) {
                                            _cmd.CommandText =
                                                $"UPDATE 'Stock_{StrStockName}'SET " +
                                                $"col_Flg = @col_Flg," +
                                                $"col_Stock = @col_Stock," +
                                                $"col_History = ifnull(col_History,'')|| @col_History " +
                                                $"WHERE " +
                                                $"col_Substrate_Num = '{_objDgv.Rows[_j].Cells[0].Value}'";

                                            if (Convert.ToInt32(_objDgv.Rows[_j].Cells[1].Value) - Convert.ToInt32(_objDgv.Rows[_j].Cells[2].Value) == 0) {
                                                _cmd.Parameters.Add("@col_Flg", DbType.String).Value = 0;
                                            }
                                            else {
                                                _cmd.Parameters.Add("@col_Flg", DbType.String).Value = 1;
                                            }

                                            _cmd.Parameters.Add("@col_Stock", DbType.String).Value = Convert.ToInt32(_objDgv.Rows[_j].Cells[1].Value) - Convert.ToInt32(_objDgv.Rows[_j].Cells[2].Value);
                                            _cmd.Parameters.Add("@col_History", DbType.String).Value = $"{StrProductNumber}({_objDgv.Rows[_j].Cells[2].Value}),";

                                            _cmd.ExecuteNonQuery();

                                            _cmd.CommandText = $@"SELECT * FROM ""Stock_{StrStockName}"" WHERE col_Substrate_Model = ""{ArrUseSubstrate[_i]}"" ORDER BY _rowid_ ASC";
                                            using SQLiteDataReader _dr = _cmd.ExecuteReader();
                                            while (_dr.Read()) {
                                                _substrateName = $"{_dr["col_Substrate_Name"]}";
                                                _substrateModel = $"{_dr["col_Substrate_Model"]}";
                                            }
                                        }

                                        if (string.IsNullOrEmpty(StrTotalSubstrate)) {
                                            StrTotalSubstrate = $"[{ArrUseSubstrate[_i]}]{_objDgv.Rows[_j].Cells[0].Value}({_objDgv.Rows[_j].Cells[2].Value})";
                                        }
                                        else {
                                            StrTotalSubstrate = $"{StrTotalSubstrate},[{ArrUseSubstrate[_i]}]{_objDgv.Rows[_j].Cells[0].Value}({_objDgv.Rows[_j].Cells[2].Value})";
                                        }

                                        using (SQLiteCommand _cmd = _con.CreateCommand()) {
                                            _cmd.CommandText =
                                                $"INSERT INTO 'Substrate_Reg_{StrStockName}'" +
                                                $"(col_Substrate_Name," +
                                                $"col_Substrate_Model," +
                                                $"col_Substrate_Num," +
                                                $"col_Decrease," +
                                                $"col_Use_P_Type," +
                                                $"col_Use_P_Num," +
                                                $"col_Use_O_Num," +
                                                $"col_Person," +
                                                $"col_RegDate," +
                                                $"col_Comment)" +
                                                $"VALUES " +
                                                $"(@col_Substrate_Name," +
                                                $"@col_Substrate_Model," +
                                                $"@col_Substrate_Num," +
                                                $"@col_Decrease," +
                                                $"@col_Use_P_Type," +
                                                $"@col_Use_P_Num," +
                                                $"@col_Use_O_Num," +
                                                $"@col_Person," +
                                                $"@col_RegDate," +
                                                $"@col_Comment)";

                                            _cmd.Parameters.Add("@col_Substrate_Name", DbType.String).Value = _substrateName;
                                            _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = _substrateModel;
                                            _cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = _objDgv.Rows[_j].Cells[0].Value;
                                            _cmd.Parameters.Add("@col_Decrease", DbType.String).Value = 0 - Convert.ToInt32(_objDgv.Rows[_j].Cells[2].Value);
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
                                                ListUsedSubstrate.Add(ArrUseSubstrate[_i]);
                                                string? _cellValue = _objDgv.Rows[_j].Cells[0].Value?.ToString();
                                                if (_cellValue != null) { ListUsedProductNumber.Add(_cellValue); }
                                                ListUsedQuantity.Add(Convert.ToInt32(_objDgv.Rows[_j].Cells[2].Value));
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                }
                            }
                        }

                        using (SQLiteCommand _cmd = _con.CreateCommand()) {
                            _cmd.CommandText =
                                $"INSERT INTO 'Product_Reg_{StrProductName}'" +
                                $"(col_Order_Num," +
                                $"col_Product_Num," +
                                $"col_Product_Type," +
                                $"col_Product_Model," +
                                $"col_Quantity," +
                                $"col_Person," +
                                $"col_RegDate," +
                                $"col_Revision," +
                                $"col_Serial_First," +
                                $"col_Serial_Last," +
                                $"col_Serial_LastNum," +
                                $"col_Comment," +
                                $"col_Use_Substrate)" +
                                $"VALUES " +
                                $"(@col_Order_Num," +
                                $"@col_Product_Num," +
                                $"@col_Product_Type," +
                                $"@col_Product_Model," +
                                $"@col_Quantity," +
                                $"@col_Person," +
                                $"@col_RegDate," +
                                $"@col_Revision," +
                                $"@col_Serial_First," +
                                $"@col_Serial_Last," +
                                $"@col_Serial_LastNum," +
                                $"@col_Comment," +
                                $"@col_Use_Substrate)";

                            _cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                            _cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                            _cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                            _cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                            _cmd.Parameters.Add("@col_Quantity", DbType.String).Value = IntQuantity;
                            _cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                            _cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                            _cmd.Parameters.Add("@col_Revision", DbType.String).Value = StrRevision;
                            _cmd.Parameters.Add("@col_Serial_First", DbType.String).Value = StrSerialFirstNumber;
                            _cmd.Parameters.Add("@col_Serial_Last", DbType.String).Value = StrSerialLastNumber;
                            _cmd.Parameters.Add("@col_Serial_LastNum", DbType.String).Value = IntSerialLastNumber;
                            _cmd.Parameters.Add("@col_Use_Substrate", DbType.String).Value = StrTotalSubstrate;
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
                StrSerial.Clear();

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
                        StrSerialType = "Label";
                        PrintBarcode(1);
                        break;
                    default:
                        break;
                }
                // バーコード印刷
                switch (IntPrintType) {
                    case 2:
                    case 3:
                        MessageBox.Show("バーコードラベルを印刷します。");
                        StrSerialType = "Barcode";
                        PrintBarcode(2);
                        break;
                    default:
                        break;
                }

                RegisterButton.Enabled = false;

                // シリアル先頭と末尾を生成
                switch (IntPrintType) {
                    case 1:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                        StrSerialType = "Label";
                        StrSerialFirstNumber = GenerateCode(IntSerialFirstNumber);
                        StrSerialLastNumber = GenerateCode(IntSerialLastNumber);
                        break;
                    case 2:
                        StrSerialType = "Barcode";
                        StrSerialFirstNumber = GenerateCode(IntSerialFirstNumber);
                        StrSerialLastNumber = GenerateCode(IntSerialLastNumber);
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
                        if (IntPrintType == 5 || IntPrintType == 6) { SubstrateListPrintButton.Enabled = true; }
                        if (IntPrintType == 6 || IntPrintType == 7) { CheckSheetPrintButton.Enabled = true; ; }
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

            string _productModel = string.Empty;

            // 製番が新規かチェック
            using (SQLiteCommand _cmd = _con.CreateCommand()) {
                _cmd.CommandText = $"SELECT * FROM 'Product_Reg_{StrProductName}' WHERE col_Product_Num = '{StrProductNumber}' ORDER BY _rowid_ ASC LIMIT 1";

                using SQLiteDataReader _dr = _cmd.ExecuteReader();
                while (_dr.Read()) {
                    _productModel = $"{_dr["col_Product_Model"]}";
                }
            }

            if (_productModel != string.Empty) {
                if (_productModel == StrProductModel) {
                    Activate();
                    DialogResult _result = MessageBox.Show($"製番[{StrProductNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                    if (_result == DialogResult.No) return false;
                }
                else {
                    Activate();
                    MessageBox.Show($"[{StrProductNumber}]は[{_productModel}]として登録があります。確認してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            _productModel = string.Empty;

            // 注文番号が新規かチェック
            using (SQLiteCommand _cmd = _con.CreateCommand()) {
                _cmd.CommandText = $"SELECT * FROM 'Product_Reg_{StrProductName}' WHERE col_Order_Num = '{StrOrderNumber}' ORDER BY _rowid_ ASC LIMIT 1";

                using SQLiteDataReader _dr = _cmd.ExecuteReader();
                while (_dr.Read()) {
                    _productModel = $"{_dr["col_Product_Model"]}";
                }
            }

            if (_productModel != string.Empty) {
                if (_productModel == StrProductModel) {
                    Activate();
                    DialogResult _result = MessageBox.Show($"注文番号[{StrOrderNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                    if (_result == DialogResult.No) return false;
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
                static bool IsNumeric(object value) { return int.TryParse(value.ToString(), out _); }

                switch (IntRegType) {
                    case 0:
                    case 1:
                        break;
                    case 2:
                    case 3:
                        if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (int _i = 0; _i <= ArrUseSubstrate.GetUpperBound(0); _i++) {

                            CheckBox? _objCbx = Controls[CheckBoxNames[_i]] as CheckBox ?? throw new Exception("objCbxがnullです。");
                            _objCbx.Enabled = true;
                            _objCbx.Checked = true;

                            DataGridView? _objDgv = Controls[DataGridViewNames[_i]] as DataGridView ?? throw new Exception("objDgvがnullです。");
                            _objDgv.Columns[2].ReadOnly = false;
                            _objDgv.Columns[3].ReadOnly = false;

                            if (_objCbx.Checked) {
                                int _intQuantityCheck = IntQuantity;
                                int _dgvRowCnt = _objDgv.Rows.Count;

                                for (int _j = 0; _j < _dgvRowCnt; _j++) {
                                    if ((bool)_objDgv.Rows[_j].Cells[3].Value) {
                                        if (!IsNumeric(_objDgv.Rows[_j].Cells[2].Value)) {
                                            throw new Exception("使用する数量に数字以外が含まれています。");
                                        }

                                        if ((int)_objDgv.Rows[_j].Cells[1].Value < (int)_objDgv.Rows[_j].Cells[2].Value) {
                                            throw new Exception("在庫より多い数量が入力されています。");
                                        }

                                        _intQuantityCheck -= (int)_objDgv.Rows[_j].Cells[2].Value;
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
                        for (int _i = 0; _i < IntQuantity; _i++) {
                            StrSerialType = "Label";
                            StrSerial.Add(GenerateCode(IntSerialFirstNumber + _i));
                        }
                        break;
                    case 2:
                        for (int _i = 0; _i < IntQuantity; _i++) {
                            StrSerialType = "Barcode";
                            StrSerial.Add(GenerateCode(IntSerialFirstNumber + _i));
                        }
                        break;
                    default:
                        break;
                }

                string _strSQLSerial = string.Join("','", StrSerial);

                List<string> _strSerialDuplication = new();
                using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                    _con.Open();

                    using SQLiteCommand _cmd = _con.CreateCommand();
                    _cmd.CommandText = $"SELECT 'col_Serial' FROM 'Serial_{StrProductName}' WHERE 'col_Serial' IN ('{_strSQLSerial}')";

                    using SQLiteDataReader _dr = _cmd.ExecuteReader();
                    while (_dr.Read()) {
                        _strSerialDuplication.Add($"{_dr["col_Serial"]}");
                    }
                }

                if (_strSerialDuplication.Count > 0) {
                    string _strSQLDuplication = string.Join($"{Environment.NewLine}", _strSerialDuplication);
                    throw new Exception($"{_strSQLDuplication}{Environment.NewLine}は既に使用されているシリアルです。");
                }

                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // 印刷処理
        private void PrintBarcode(int PrintFlg) {
            // PrintDocumentオブジェクトの作成
            using System.Drawing.Printing.PrintDocument _pd = new();

            // PrintPageイベントハンドラの追加
            _pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

            LabelProNumLabelsToPrint = IntQuantity;
            LabelProPageNum = 0;
            FontUnderbar = false;

            switch (PrintFlg) {
                case 1:
                    ProductRegistration2PrintDialog.Document = _pd;
                    DialogResult _r = ProductRegistration2PrintDialog.ShowDialog();

                    if (_r == DialogResult.OK) {
                        ProductRegistration2PrintDialog.Document.Print();

                        if (IntPageCnt >= 2) {
                            MessageBox.Show($"{IntPageCnt}枚印刷されます。2枚目以降は1行目から印刷されます。");
                        }
                    }
                    else {
                        return;
                    }
                    break;
                case 2:
                    ProductRegistration2PrintPreviewDialog.Document = _pd;
                    ProductRegistration2PrintPreviewDialog.ShowDialog();
                    break;
            }
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                int _startLineLabel = (int)SerialPrintPostionNumericUpDown.Value - 1;
                Point _headerPos = new(0, 0);
                string _headerString = string.Empty;
                Font _headerFooterFont = new("Arial", 6);
                int _intNumLabels = 0;
                int _intCountNumLabels = 0;

                int _maxX = 0;
                int _maxY = 0;
                float _sizeX = 0;
                float _sizeY = 0;
                double _offsetX = 0;
                double _offsetY = 0;
                double _intervalX = 0;
                double _intervalY = 0;

                switch (StrSerialType) {
                    case "Label":
                        if (SettingsLabelPro == null) { throw new Exception("SettingsLabelProがnullです。"); }
                        _maxX = SettingsLabelPro.LabelProPageSettings.NumLabelsX;
                        _maxY = SettingsLabelPro.LabelProPageSettings.NumLabelsY;
                        _sizeX = (float)SettingsLabelPro.LabelProPageSettings.SizeX;
                        _sizeY = (float)SettingsLabelPro.LabelProPageSettings.SizeY;
                        _offsetX = SettingsLabelPro.LabelProPageSettings.OffsetX;
                        _offsetY = SettingsLabelPro.LabelProPageSettings.OffsetY;
                        _intervalX = SettingsLabelPro.LabelProPageSettings.IntervalX;
                        _intervalY = SettingsLabelPro.LabelProPageSettings.IntervalY;
                        _headerPos = SettingsLabelPro.LabelProPageSettings.HeaderPos;
                        _headerString = ConvertHeaderFooterString(SettingsLabelPro.LabelProPageSettings.HeaderString);
                        _headerFooterFont = SettingsLabelPro.LabelProPageSettings.HeaderFooterFont;
                        _intNumLabels = SettingsLabelPro.LabelProLabelSettings.NumLabels;
                        _intCountNumLabels = SettingsLabelPro.LabelProLabelSettings.NumLabels;
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
                        break;
                    default:
                        break;
                }

                int _startLineBarcode = (int)BarcodePrintPostionNumericUpDown.Value - 1;

                Point _offset;
                if (!ProductRegistration2PrintDocument.PrintController.IsPreview) {
                    _offsetX -= e.PageSettings.HardMarginX * 0.254;
                    _offsetY -= e.PageSettings.HardMarginY * 0.254;
                    if (LabelProPageNum == 0) {
                        _offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (_startLineBarcode * (_intervalY + _sizeY))));
                    }
                    else {
                        _offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (_intervalY + _sizeY))));
                    }
                }
                else {
                    _offset = new Point(0, 0);
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;

                _headerPos.Offset(_offset);
                e.Graphics.DrawString(_headerString, _headerFooterFont, Brushes.Black, _headerPos);

                int _barcodePageNum = 0;
                if (_barcodePageNum == 0) {
                    _offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (_startLineBarcode * (_intervalY + _sizeY))));
                }
                else {
                    _offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (_intervalY + _sizeY))));
                }

                if (LabelProPageNum >= 1) { _startLineBarcode = 0; }


                int _y = 0;
                for (_y = _startLineBarcode; _y < _maxY; _y++) {
                    int _x = 0;
                    for (_x = 0; _x < _maxX; _x++) {
                        string _s = GenerateCode(LabelProNSerial);
                        float _posX = (float)(_offsetX + (_x * (_intervalX + _sizeX)));
                        float _posY = (float)(_offsetY + (_y * (_intervalY + _sizeY)));
                        e.Graphics.DrawImage(MakeLabelImage(_s, (int)e.Graphics.DpiX, 1), _posX, _posY, _sizeX, _sizeY);

                        LabelProNLabel = 0;
                        LabelProNumLabelsToPrint--;

                        if (LabelProNumLabelsToPrint <= 0) {
                            _intCountNumLabels--;
                            if (_intCountNumLabels <= 0) {
                                e.HasMorePages = false;
                                LabelProPageNum = 0;
                                int _txtNumPublish = 0;
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
                    LabelProPageNum++;
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
            string _monthCode = DateTime.Parse(StrRegDate).ToString("MM");

            _monthCode = _monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => string.Empty
            };

            string _outputCode = StrSerialType switch {
                "Label" => SettingsLabelPro.LabelProLabelSettings.Format,
                "Barcode" => SettingsBarcodePro.BarcodeProLabelSettings.Format,
                _ => string.Empty
            };

            string _serialCode = Convert.ToInt32(serialCode).ToString($"D{IntSerialDigit}");
            _outputCode = _outputCode.Replace("%Y", DateTime.Parse(StrRegDate).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(StrRegDate).ToString("MM"))
                                    .Replace("%T", StrInitial)
                                    .Replace("%R", StrRevision)
                                    .Replace("%M", _monthCode[^1..])
                                    .Replace("%S", _serialCode);

            return _outputCode;
        }
        private Bitmap MakeLabelImage(string text, int resolution, int magnitude) {
            Bitmap _labelImage = new(0, 0);
            Graphics _g;
            SizeF _stringSize;
            decimal _sizeX;
            decimal _sizeY;
            decimal _fontSize;
            float _stringPosX;
            float _stringPosY;
            Font _fnt;
            switch (StrSerialType) {
                case "Label":
                    if (SettingsLabelPro == null) { throw new Exception("SettingsLabelProがnull"); }
                    _sizeX = (decimal)SettingsLabelPro.LabelProPageSettings.SizeX / 25.4M * resolution * magnitude;
                    _sizeY = (decimal)SettingsLabelPro.LabelProPageSettings.SizeY / 25.4M * resolution * magnitude;
                    _fontSize = (decimal)SettingsLabelPro.LabelProLabelSettings.Font.SizeInPoints / 72.0M * resolution * magnitude;
                    _stringPosY = (int)((decimal)SettingsLabelPro.LabelProLabelSettings.StringPosY / 25.4M * resolution * magnitude);
                    FontStyle _style = FontUnderbar ? FontStyle.Underline : FontStyle.Regular;
                    _fnt = new Font(SettingsLabelPro.LabelProLabelSettings.Font.Name, (float)_fontSize, _style);

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
                    if (resolution == DisplayResolution) { _barWeight = 1; }
                    else { _barWeight = (int)(1 * resolution / DisplayResolution / DisplayMagnitude); }

                    using (Image _img = Code128Rendering.MakeBarcodeImage(text, _barWeight, true)) {
                        decimal _imageWidth = (decimal)(_img.Width * SettingsBarcodePro.BarcodeProLabelSettings.BarcodeMagnitude);

                        if (_imageWidth > _labelImage.Width) { MessageBox.Show($"バーコードの幅がラベル幅を超えています{_imageWidth}>{_labelImage.Width}"); }

                        _stringSize = _g.MeasureString(text, _fnt);

                        if (SettingsBarcodePro.BarcodeProLabelSettings.AlignStringCenter) { _stringPosX = (float)((_labelImage.Width / 2) - (_stringSize.Width / 2)); }
                        else { _stringPosX = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4M * resolution * magnitude); }

                        _stringPosY = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosY / 25.4M * resolution * magnitude);

                        float _barCodePosX;
                        if (SettingsBarcodePro.BarcodeProLabelSettings.AlignBarcodeCenter) { _barCodePosX = (float)((_labelImage.Width / 2) - (_imageWidth / 2)); }
                        else { _barCodePosX = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4M * resolution * magnitude); }

                        float _barCodePosY = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.BarcodePosY / 25.4M * resolution * magnitude);

                        _g.DrawString(text, _fnt, Brushes.Black, _stringPosX, _stringPosY);

                        int _barcodeHeight = (int)(SettingsBarcodePro.BarcodeProLabelSettings.BarcodeHeight / 25.4F * resolution * magnitude);
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
            CheckBox _checkBox = (CheckBox)sender;
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
        private void CheckSheetPrint() {
            try {
                List<string> _filePath = new();
                List<string> _sheetName = new();

                List<string> _orderNumberRange = new();
                List<string> _regYearRange = new();
                List<string> _regMonthRange = new();
                List<string> _regDayRange = new();
                List<string> _orderFirstSerialRange = new();
                List<string> _orderLastSerialRange = new();

                using FileStream _fileStream = new($@"{Environment.CurrentDirectory}./config/Excel/ConfigCheckSheet.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using XLWorkbook _workBook = new(_fileStream);
                IXLWorksheet _workSheet = _workBook.Worksheet("Sheet1");

                // セル検索
                foreach (IXLCell _cell in _workSheet.Search(StrProductModel)) {
                    int _findRow = _cell.Address.RowNumber;
                    // ワークシートのセルから値を取得してリストに格納
                    _filePath.Add(_workSheet.Cell("J" + _findRow).Value.ToString());
                    _sheetName.Add(_workSheet.Cell("I" + _findRow).Value.ToString());

                    _orderNumberRange.Add(_workSheet.Cell("C" + _findRow).Value.ToString());
                    _regYearRange.Add(_workSheet.Cell("D" + _findRow).Value.ToString());
                    _regMonthRange.Add(_workSheet.Cell("E" + _findRow).Value.ToString());
                    _regDayRange.Add(_workSheet.Cell("F" + _findRow).Value.ToString());
                    _orderFirstSerialRange.Add(_workSheet.Cell("G" + _findRow).Value.ToString());
                    _orderLastSerialRange.Add(_workSheet.Cell("H" + _findRow).Value.ToString());
                }

                if (_filePath.Count == 0) {
                    throw new Exception($"Configに品目番号:{StrProductModel}が見つかりません。");
                }

                // 見つかった品目番号の数だけ印刷
                for (int _i = 0; _i < _filePath.Count; _i++) {
                    // チェックシート作成
                    using FileStream _fileStream2 = new($@"{_filePath[_i]}", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using XLWorkbook _workBook2 = new(_fileStream2);
                    IXLWorksheet _workSheet2 = _workBook2.Worksheet(_sheetName[_i]);

                    _workSheet2.Cell(_orderNumberRange[_i]).Value = StrOrderNumber;
                    _workSheet2.Cell(_regYearRange[_i]).Value = $"{DateTime.Parse(StrRegDate):yy}年";
                    _workSheet2.Cell(_regMonthRange[_i]).Value = $"{DateTime.Parse(StrRegDate):MM}月";
                    _workSheet2.Cell(_regDayRange[_i]).Value = $"{DateTime.Parse(StrRegDate):dd}日";
                    _workSheet2.Cell(_orderFirstSerialRange[_i]).Value = StrSerialFirstNumber;
                    _workSheet2.Cell(_orderLastSerialRange[_i]).Value = StrSerialLastNumber;

                    // 印刷
                    Excel.Application _xlApp = new() {
                        Visible = true // Excelウィンドウを表示します。
                    };

                    // ワークブックを開く。
                    Excel.Workbooks _xlBooks = _xlApp.Workbooks;
                    Excel.Workbook _xlBook = _xlBooks.Open($@"{System.Environment.CurrentDirectory}./config/Excel/temporarily_{_i}.xlsx");

                    // ワークシートを選択
                    Excel.Sheets _xlSheets = _xlBook.Sheets;
                    Excel.Worksheet _xlSheet = _xlSheets[1];

                    // ワークシートを印刷
                    _xlSheet.PrintOut(Preview: true);

                    // ワークブックを閉じてExcelを終了します。
                    _xlBook.Close(false);
                    _xlApp.Quit();
                    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlSheet);
                    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlSheets);
                    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlBook);
                    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlBooks);
                    _ = System.Runtime.InteropServices.Marshal.ReleaseComObject(_xlApp);
                }

            } catch (Exception ex) {
                MessageBox.Show($"{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // リスト印刷
        private void ListPrint() {
            try {
                string _cellRange = string.Empty;
                string _filePath = string.Empty;
                string _sheetName = string.Empty;

                string _orderNumberRange = string.Empty;
                string _regYearRange = string.Empty;
                string _regMonthRange = string.Empty;
                string _regDayRange = string.Empty;
                string _orderFirstSerialRange = string.Empty;
                string _orderLastSerialRange = string.Empty;

                using (FileStream _fileStream = new($@"{Environment.CurrentDirectory}./config/Excel/ConfigList.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    using XLWorkbook _workBook = new(_fileStream);
                    IXLWorksheet _workSheet = _workBook.Worksheet(1);

                    int _findRow = 0;
                    // セル検索
                    foreach (IXLCell _cell in _workSheet.Search(StrProductModel)) {
                        _findRow = _cell.Address.RowNumber;
                    }

                    if (_findRow == 0) {
                        throw new Exception($"Configに品目番号:{StrProductModel}が見つかりません。");
                    }

                    // ワークシートのセルから値を取得
                    _filePath = _workSheet.Cell(_findRow, 10).Value.ToString();
                    _sheetName = _workSheet.Cell(_findRow, 9).Value.ToString();

                    _orderNumberRange = _workSheet.Cell(_findRow, 3).Value.ToString();
                    _regYearRange = _workSheet.Cell(_findRow, 4).Value.ToString();
                    _regMonthRange = _workSheet.Cell(_findRow, 5).Value.ToString();
                    _regDayRange = _workSheet.Cell(_findRow, 6).Value.ToString();
                    _orderFirstSerialRange = _workSheet.Cell(_findRow, 7).Value.ToString();
                    _orderLastSerialRange = _workSheet.Cell(_findRow, 8).Value.ToString();
                }

                using (FileStream _fileStream = new($@"{_filePath}", FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    using XLWorkbook _workBook = new(_fileStream);
                    IXLWorksheet _workSheet = _workBook.Worksheet(_sheetName);

                    _workSheet.Cell(_orderNumberRange).Value = StrOrderNumber;
                    _workSheet.Cell(_regYearRange).Value = $"{DateTime.Parse(StrRegDate):yy}年";
                    _workSheet.Cell(_regMonthRange).Value = $"{DateTime.Parse(StrRegDate):MM}月";
                    _workSheet.Cell(_regDayRange).Value = $"{DateTime.Parse(StrRegDate):dd}日";
                    _workSheet.Cell(_orderFirstSerialRange).Value = StrSerialFirstNumber;
                    _workSheet.Cell(_orderLastSerialRange).Value = StrSerialLastNumber;

                    //引数に保存先パスを指定
                    _workBook.SaveAs($@"{Environment.CurrentDirectory}./config/Excel/temporarily.xlsx");
                }
                // 印刷
                Excel.Application _xlApp = new() {
                    Visible = true // Excelウィンドウを表示します。
                };

                // ワークブック開く
                Excel.Workbooks _xlBooks = _xlApp.Workbooks;
                Excel.Workbook _xlBook = _xlBooks.Open($@"{Environment.CurrentDirectory}./config/Excel/temporarily.xlsx");

                // ワークシート選択
                Excel.Sheets _xlSheets = _xlBook.Sheets;
                Excel.Worksheet _xlSheet = _xlSheets[1];

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
            try {
                StrSerialType = "Label";
                PrintBarcode(1);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        private void シリアルラベル印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                StrSerialType = "Label";
                PrintBarcode(2);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        private void バーコード印刷ToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                StrSerialType = "Barcode";
                PrintBarcode(1);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }
        private void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            try {
                StrSerialType = "Barcode";
                PrintBarcode(2);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
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
            MessageBox.Show($"" +
                $"StrProductName\t\t[{StrProductName}]\r\n" +
                $"StrProductModel\t\t[{StrProductModel}]\r\n" +
                $"StrStockName\t\t[{StrStockName}]\r\n" +
                $"StrProductType\t\t[{StrProductType}]\r\n" +
                $"StrOrderNumber\t\t[{StrOrderNumber}]\r\n" +
                $"StrProductNumber\t\t[{StrProductNumber}]\r\n" +
                $"StrRevision\t\t[{StrRevision}]\r\n" +
                $"StrRegDate\t\t[{StrRegDate}]\r\n" +
                $"StrPerson\t\t\t[{StrPerson}]\r\n" +
                $"IntQuantity\t\t[{IntQuantity}]\r\n" +
                $"IntSerialFirstNumber\t[{IntSerialFirstNumber}]\r\n" +
                $"IntSerialLastNumber\t[{IntSerialLastNumber}]\r\n" +
                $"StrInitial\t\t\t[{StrInitial}]\r\n" +
                $"IntRegType\t\t[{IntRegType}]\r\n" +
                $"IntPrintType\t\t[{IntPrintType}]\r\n" +
                $"IntSerialDigit\t\t[{IntSerialDigit}]" +
                $"", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) { ListPrint(); }
        private void CheckSheetPrintButton_Click(object sender, EventArgs e) { CheckSheetPrint(); }
    }
}
