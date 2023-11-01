using GenCode128;
using LabelGenerator;
using System.Data;
using System.Data.SQLite;

namespace ProductDataBase {
    public partial class ProductRegistration2Window : Form {

        public CSettingsLabelPro? SettingsLabelPro { get; set; }
        public string StrLabelSettingFilePath { get; set; } = string.Empty;

        public CSettingsBarcodePro? SettingsBarcodePro { get; set; }
        public string StrBarcodeSettingFilePath { get; set; } = string.Empty;

        public string? StrFontName { get; set; }
        public int IntFontSize { get; set; }

        public string? StrProductName { get; set; }
        public string? StrStockName { get; set; }
        public string? StrProductType { get; set; }
        public string? StrProductModel { get; set; }
        public string? StrUseSubstrate { get; set; }
        public string[]? ArrUseSubstrate;
        public string? StrInitial { get; set; }
        public string? StrOrderNumber { get; set; }
        public string? StrProductNumber { get; set; }
        public string? StrRegDate { get; set; }
        public string? StrPerson { get; set; }
        public string? StrRevision { get; set; }
        public string? StrComment { get; set; }

        public int IntQuantity { get; set; }
        public int IntRegType { get; set; }
        public int IntPrintType { get; set; }
        public int IntCheckBin { get; set; }
        public int IntSerialDigit { get; set; }
        public int IntSerialFirstNumber { get; set; }

        public Bitmap? LabelProBmp { get; }
        public int LabelProPageNum { get; set; }
        public int LabelProNSerial { get; set; }
        public int LabelProNLabel { get; set; }
        public int LabelProNumLabelsToPrint { get; set; }

        public decimal DisplayResolution { get; } = 96.0m;
        public int DisplayMagnitude { get; } = 3;
        public int IntPageCnt { get; set; } = 1;

        public string? StrProness1 { get; }
        public string? StrProness2 { get; }
        public string? StrProness3 { get; }
        public int StrProness4 { get; }
        public string? StrProness5 { get; }

        private string? StrSerialType;
        private string? StrSerialFirstNumber;
        private string? StrSerialLastNumber;
        private string? StrTotalSubstrate;
        private readonly int IntSerialLastNumber;
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

        public ProductRegistration2Window(ProductRegistration2Window productRegistration1Window) {
            StrFontName = productRegistration1Window.StrFontName;
            IntFontSize = productRegistration1Window.IntFontSize;

            StrProductName = productRegistration1Window.StrProductName;
            StrStockName = productRegistration1Window.StrStockName;
            StrProductType = productRegistration1Window.StrProductType;
            StrProductModel = productRegistration1Window.StrProductModel;
            StrUseSubstrate = productRegistration1Window.StrUseSubstrate;
            StrInitial = productRegistration1Window.StrInitial;
            IntQuantity = productRegistration1Window.IntQuantity;
            IntRegType = productRegistration1Window.IntRegType;
            StrPerson = productRegistration1Window.StrPerson;
            IntPrintType = productRegistration1Window.IntPrintType;
            IntCheckBin = productRegistration1Window.IntCheckBin;
            IntSerialDigit = productRegistration1Window.IntSerialDigit;
            IntSerialFirstNumber = productRegistration1Window.IntSerialFirstNumber;
            IntSerialLastNumber = IntSerialFirstNumber + IntQuantity - 1;
            StrOrderNumber = productRegistration1Window.StrOrderNumber;
            StrProductNumber = productRegistration1Window.StrProductNumber;
            StrRegDate = productRegistration1Window.StrRegDate;
            StrRevision = productRegistration1Window.StrRevision;
            StrComment = productRegistration1Window.StrComment;

            StrProness1 = productRegistration1Window.StrProness1;
            StrProness2 = productRegistration1Window.StrProness2;
            StrProness3 = productRegistration1Window.StrProness3;
            StrProness4 = productRegistration1Window.StrProness4;
            StrProness5 = productRegistration1Window.StrProness5;
        }

        public ProductRegistration2Window() {
            InitializeComponent();
        }
        // ロードイベント
        private void LoadEvents() {
            try {
                if (StrFontName != null) { Font = new Font(StrFontName, IntFontSize); }

                RegisterButton.Enabled = true;

                bool _quantityFlg = false;
                string _strQuantity = string.Empty;
                string _strSubstrateName = string.Empty;

                if (StrUseSubstrate != null) { ArrUseSubstrate = StrUseSubstrate.Split(","); }

                switch (IntRegType) {
                    case 2:
                        if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (int i = 0; i <= ArrUseSubstrate.GetUpperBound(0); i++) {
                            int _intQuantity = IntQuantity;

                            CheckBox? objCbx = Controls[CheckBoxNames[i]] as CheckBox;

                            if (objCbx != null) {
                                objCbx.Enabled = true;
                                objCbx.Checked = true;
                            }

                            DataGridView? objDgv = Controls[DataGridViewNames[i]] as DataGridView;
                            if (objDgv != null) {
                                objDgv.Columns[2].ReadOnly = false;
                                objDgv.Columns[3].ReadOnly = false;
                            }

                            using SQLiteConnection con = new(MainWindow.GetConnectionString2());
                            con.Open();

                            using SQLiteCommand cmd = con.CreateCommand();
                            // 使用基板表示
                            cmd.CommandText = $"SELECT col_Substrate_Name FROM 'Stock_{StrStockName}' WHERE col_Substrate_Model = '{ArrUseSubstrate[i]}'";
                            using (SQLiteDataReader dr = cmd.ExecuteReader()) {
                                if (dr.Read()) {
                                    string _substrateName = $"{dr["col_Substrate_Name"]}";
                                    if (objCbx != null) { objCbx.Text = $"{_substrateName} - {ArrUseSubstrate[i]}"; }
                                }
                            }

                            // 在庫テーブルからデータ取得
                            cmd.CommandText = $"SELECT col_Substrate_num, col_Stock, col_Substrate_Name FROM 'Stock_{StrStockName}' WHERE col_flg = 1 And col_Substrate_Model = '{ArrUseSubstrate[i]}' ORDER BY _rowid_ ASC";
                            using (SQLiteDataReader dr = cmd.ExecuteReader()) {
                                while (dr.Read()) {
                                    string _strSubstrateNumber = $"{dr["col_Substrate_num"]}";
                                    int _intStock = Convert.ToInt32(dr["col_Stock"]);
                                    string _substrateName = $"{dr["col_Substrate_Name"]}";

                                    objDgv?.Rows.Add(_strSubstrateNumber, _intStock);

                                    if (_intQuantity >= _intStock) {
                                        _intQuantity -= _intStock;
                                        if (objDgv != null) {
                                            objDgv.Rows[^1].Cells[2].Value = _intStock;
                                            objDgv.Rows[^1].Cells[3].Value = true;
                                        }
                                    }
                                    else {
                                        if (_intQuantity == 0)
                                            if (objDgv != null) {
                                                objDgv.Rows[^1].Cells[2].Value = null;
                                            }
                                            else {
                                                if (objDgv != null) {
                                                    objDgv.Rows[^1].Cells[2].Value = _intQuantity;
                                                    objDgv.Rows[^1].Cells[3].Value = true;
                                                    _intQuantity = 0;
                                                }
                                            }
                                    }

                                    if (_intQuantity > 0) {
                                        _quantityFlg = false;
                                        _strQuantity += $"[{_substrateName}]{Environment.NewLine}";
                                    }

                                    if (_intQuantity == 0) { _quantityFlg = true; }
                                }
                            }
                        }

                        if (_quantityFlg == false) {
                            this.Activate();
                            MessageBox.Show($"在庫が足りません。{Environment.NewLine}{_strQuantity}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        break;
                    case 3:
                        if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (int i = 0; i <= ArrUseSubstrate.GetUpperBound(0); i++) {
                            int _intQuantity = IntQuantity;

                            CheckBox? objCbx = Controls[CheckBoxNames[i]] as CheckBox;

                            if (objCbx != null) {
                                objCbx.Enabled = true;
                                objCbx.Checked = true;
                            }

                            DataGridView? objDgv = Controls[DataGridViewNames[i]] as DataGridView;
                            if (objDgv != null) {
                                objDgv.Columns[2].ReadOnly = false;
                                objDgv.Columns[3].ReadOnly = false;
                            }

                            using SQLiteConnection con = new(MainWindow.GetConnectionString2());
                            con.Open();
                            using SQLiteCommand cmd = con.CreateCommand();
                            cmd.CommandText = $"SELECT * FROM 'Stock_{StrStockName}' WHERE col_Substrate_Model = '{ArrUseSubstrate[i]}'";
                            using (SQLiteDataReader dr = cmd.ExecuteReader()) {
                                while (dr.Read()) {
                                    if (objCbx != null) { objCbx.Text = $"{dr["col_Substrate_Name"]} - {ArrUseSubstrate[i]}"; }
                                }
                            }

                            cmd.CommandText = $"SELECT * FROM 'Stock_{StrStockName}' WHERE col_flg = 1 And col_Substrate_Model = '{ArrUseSubstrate[i]}' ORDER BY _rowid_ ASC";
                            using (SQLiteDataReader dr = cmd.ExecuteReader()) {
                                int j = 0;
                                while (dr.Read()) {
                                    string _StrSubstrateNumber = $"{dr["col_Substrate_Num"]}";
                                    string _strOrderNumber = $"{dr["col_Order_Num"]}";
                                    int _intStock = Convert.ToInt32(dr["col_Stock"]);
                                    _strSubstrateName = $"{dr["col_Substrate_Name"]}";

                                    objDgv?.Rows.Add(_StrSubstrateNumber, _intStock);

                                    if (_strOrderNumber == StrOrderNumber) {
                                        if (objDgv != null) {
                                            objDgv.Rows[j].Cells[2].Value = _intQuantity;
                                            objDgv.Rows[j].Cells[3].Value = true;
                                            // 必要数量分割り当てられたかチェック
                                            if (_intQuantity > Convert.ToInt32(objDgv.Rows[j].Cells[1].Value)) {
                                                _quantityFlg = false;
                                                _strQuantity = $"{_strQuantity}[{_strSubstrateName}]{Environment.NewLine}";
                                            }
                                        }
                                    }
                                    j++;
                                }
                            }
                        }

                        if (!_quantityFlg) {
                            this.Activate();
                            MessageBox.Show($"在庫が足りません。{Environment.NewLine}{_strQuantity}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
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
                }

                LoadSettings(StrLabelSettingFilePath, StrBarcodeSettingFilePath);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadSettings(string strLabelSettingFilePath, string strBarcodeSettingFilePath) {
            System.Xml.Serialization.XmlSerializer serializerLabel = new(typeof(CSettingsLabelPro));
            System.Xml.Serialization.XmlSerializer serializerBarcode = new(typeof(CSettingsBarcodePro));

            try {
                if (strLabelSettingFilePath != string.Empty) {
                    StreamReader? srLabel = new(strLabelSettingFilePath, new System.Text.UTF8Encoding(false));
                    SettingsLabelPro = serializerLabel.Deserialize(srLabel) as CSettingsLabelPro;
                    srLabel?.Close();
                }
                if (strBarcodeSettingFilePath != string.Empty) {
                    StreamReader? srBarcode = new(strBarcodeSettingFilePath, new System.Text.UTF8Encoding(false));
                    SettingsBarcodePro = serializerBarcode.Deserialize(srBarcode) as CSettingsBarcodePro;
                    srBarcode?.Close();
                }
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // クロージングイベント
        private void ClosingEvents() {
            try {
                System.Xml.Serialization.XmlSerializer serializerLabel = new(typeof(CSettingsLabelPro));
                System.IO.StreamWriter? swLabel = null;
                System.Xml.Serialization.XmlSerializer serializerBarcode = new(typeof(CSettingsBarcodePro));
                System.IO.StreamWriter? swBarcode = null;

                switch (IntPrintType) {
                    case 1:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                        swLabel = new StreamWriter(StrLabelSettingFilePath, false, new System.Text.UTF8Encoding(false));
                        serializerLabel.Serialize(swLabel, SettingsLabelPro);
                        swLabel?.Close();
                        break;
                    case 2:
                        swBarcode = new StreamWriter(StrBarcodeSettingFilePath, false, new System.Text.UTF8Encoding(false));
                        serializerBarcode.Serialize(swBarcode, SettingsBarcodePro);
                        swBarcode?.Close();
                        break;
                    case 3:
                        swLabel = new StreamWriter(StrLabelSettingFilePath, false, new System.Text.UTF8Encoding(false));
                        serializerLabel.Serialize(swLabel, SettingsLabelPro);
                        swLabel?.Close();
                        swBarcode = new StreamWriter(StrBarcodeSettingFilePath, false, new System.Text.UTF8Encoding(false));
                        serializerBarcode.Serialize(swBarcode, SettingsBarcodePro);
                        swBarcode?.Close();
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show($"設定ファイルの保存に失敗しました。{Environment.NewLine}{ex.Message}");
            } finally {
                Dispose();
            }
        }
        // 登録処理
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
                }
                // バーコード印刷
                switch (IntPrintType) {
                    case 2:
                    case 3:
                        MessageBox.Show("バーコードラベルを印刷します。");
                        StrSerialType = "Barcode";
                        PrintBarcode(2);
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
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool NumberCheck() {
            using SQLiteConnection con = new(MainWindow.GetConnectionString2());
            con.Open();

            string _productModel = string.Empty;

            // 製番が新規かチェック
            using (SQLiteCommand cmd = con.CreateCommand()) {
                cmd.CommandText = $"SELECT * FROM 'Product_Reg_{StrProductName}' WHERE col_Product_Num = '{StrProductNumber}' ORDER BY _rowid_ ASC LIMIT 1";

                using SQLiteDataReader dr = cmd.ExecuteReader();
                while (dr.Read()) {
                    _productModel = $"{dr["col_Product_Model"]}";
                }
            }

            if (_productModel != string.Empty) {
                if (_productModel == StrProductModel) {
                    DialogResult result;
                    Activate();
                    result = MessageBox.Show($"製番[{StrProductNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                    if (result == DialogResult.No) return false;
                }
                else {
                    Activate();
                    MessageBox.Show($"[{StrProductNumber}]は[{_productModel}]として登録があります。確認してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            _productModel = string.Empty;

            // 注文番号が新規かチェック
            using (SQLiteCommand cmd = con.CreateCommand()) {
                cmd.CommandText = $"SELECT * FROM 'Product_Reg_{StrProductName}' WHERE col_Order_Num = '{StrOrderNumber}' ORDER BY _rowid_ ASC LIMIT 1";

                using SQLiteDataReader dr = cmd.ExecuteReader();
                while (dr.Read()) {
                    _productModel = $"{dr["col_Product_Model"]}";
                }
            }

            if (_productModel != string.Empty) {
                if (_productModel == StrProductModel) {
                    DialogResult result;
                    Activate();
                    result = MessageBox.Show($"注文番号[{StrOrderNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                    if (result == DialogResult.No) return false;
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
                int _dgvRowCnt;
                int _intQuantityCheck;

                switch (IntRegType) {
                    case 0:
                    case 1:
                        break;
                    case 2:
                    case 3:
                        if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (int i = 0; i <= ArrUseSubstrate.GetUpperBound(0); i++) {

                            CheckBox? objCbx = Controls[CheckBoxNames[i]] as CheckBox ?? throw new Exception("objCbxがnullです。");
                            objCbx.Enabled = true;
                            objCbx.Checked = true;

                            DataGridView? objDgv = Controls[DataGridViewNames[i]] as DataGridView ?? throw new Exception("objDgvがnullです。");
                            objDgv.Columns[2].ReadOnly = false;
                            objDgv.Columns[3].ReadOnly = false;

                            if (objCbx.Checked) {
                                _intQuantityCheck = IntQuantity;
                                _dgvRowCnt = objDgv.Rows.Count;

                                for (int j = 0; j < _dgvRowCnt; j++) {
                                    if ((bool)objDgv.Rows[j].Cells[3].Value) {
                                        if (!IsNumeric(objDgv.Rows[j].Cells[2].Value)) {
                                            throw new Exception("使用する数量に数字以外が含まれています。");
                                        }

                                        if ((int)objDgv.Rows[j].Cells[1].Value < (int)objDgv.Rows[j].Cells[2].Value) {
                                            throw new Exception("在庫より多い数量が入力されています。");
                                        }

                                        _intQuantityCheck -= (int)objDgv.Rows[j].Cells[2].Value;
                                    }
                                }

                                if (_intQuantityCheck != 0) {
                                    throw new Exception("入力された数量の合計が必要数と一致しません。");
                                }
                            }
                        }
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
                List<string> _strSerialDuplication = new();
                string _strSQLSerial = string.Empty;
                string _strSQLDuplication = string.Empty;

                switch (IntPrintType) {
                    case 1:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 8:
                        for (int i = 0; i < IntQuantity; i++) {
                            StrSerialType = "Label";
                            StrSerial.Add(GenerateCode(IntSerialFirstNumber + i));
                        }
                        break;
                    case 2:
                        for (int i = 0; i < IntQuantity; i++) {
                            StrSerialType = "Barcode";
                            StrSerial.Add(GenerateCode(IntSerialFirstNumber + i));
                        }
                        break;
                }

                _strSQLSerial = string.Join("','", StrSerial);

                using (SQLiteConnection con = new(MainWindow.GetConnectionString2())) {
                    con.Open();

                    using SQLiteCommand cmd = con.CreateCommand();
                    cmd.CommandText = $"SELECT 'col_Serial' FROM 'Serial_{StrProductName}' WHERE 'col_Serial' IN ('{_strSQLSerial}')";

                    using SQLiteDataReader dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        _strSerialDuplication.Add($"{dr["col_Serial"]}");
                    }
                }

                if (_strSerialDuplication.Count > 0) {
                    _strSQLDuplication = string.Join($"{Environment.NewLine}", _strSerialDuplication);
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
            System.Drawing.Printing.PrintDocument pd = new();

            // PrintPageイベントハンドラの追加
            pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

            LabelProNumLabelsToPrint = IntQuantity;
            LabelProPageNum = 0;
            FontUnderbar = false;

            switch (PrintFlg) {
                case 1:
                    ProductRegistration2PrintDialog.Document = pd;
                    DialogResult r = ProductRegistration2PrintDialog.ShowDialog();

                    if (r == DialogResult.OK) {
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
                    ProductRegistration2PrintPreviewDialog.Document = pd;
                    ProductRegistration2PrintPreviewDialog.ShowDialog();
                    break;
            }
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }
                int _txtNumPublish = 0;
                int _maxX = 0;
                int _maxY = 0;
                float _sizeX = 0;
                float _sizeY = 0;
                double _offsetX = 0;
                double _offsetY = 0;
                double _intervalX = 0;
                double _intervalY = 0;
                float _posX = 0, _posY = 0;
                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                Point _offset;
                int _x = 0, _y = 0;
                string _s = string.Empty;

                int _startLineLabel = (int)SerialPrintPostionNumericUpDown.Value - 1;
                int _startLineBarcode = (int)BarcodePrintPostionNumericUpDown.Value - 1;

                int _barcodePageNum = 0;
                Point _headerPos = new(0, 0);
                string _headerString = string.Empty;
                Font _headerFooterFont = new("Arial", 6);
                int _intNumLabels = 0;
                int _intCountNumLabels = 0;


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
                }

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

                if (_barcodePageNum == 0) {
                    _offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (_startLineBarcode * (_intervalY + _sizeY))));
                }
                else {
                    _offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (_intervalY + _sizeY))));
                }

                if (LabelProPageNum >= 1) { _startLineBarcode = 0; }


                for (_y = _startLineBarcode; _y < _maxY; _y++) {
                    for (_x = 0; _x < _maxX; _x++) {
                        _s = GenerateCode(LabelProNSerial);
                        _posX = (float)(_offsetX + (_x * (_intervalX + _sizeX)));
                        _posY = (float)(_offsetY + (_y * (_intervalY + _sizeY)));
                        e.Graphics.DrawImage(MakeLabelImage(_s, (int)e.Graphics.DpiX, 1), _posX, _posY, _sizeX, _sizeY);

                        LabelProNLabel = 0;
                        LabelProNumLabelsToPrint--;

                        if (LabelProNumLabelsToPrint <= 0) {
                            _intCountNumLabels--;
                            if (_intCountNumLabels <= 0) {
                                e.HasMorePages = false;
                                LabelProPageNum = 0;
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
            string _serialCode = Convert.ToInt32(serialCode).ToString($"D{IntSerialDigit}");
            string _outputCode = string.Empty;
            string _monthCode = StrRegDate != null ? DateTime.Parse(StrRegDate).ToString("MM") : string.Empty;

            _monthCode = _monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => string.Empty
            };

            switch (StrSerialType) {
                case "Label":
                    if (SettingsLabelPro != null) { _outputCode = SettingsLabelPro.LabelProLabelSettings.Format; }
                    break;
                case "Barcode":
                    if (SettingsBarcodePro != null) { _outputCode = SettingsBarcodePro.BarcodeProLabelSettings.Format; }
                    break;
            }

            if (StrRegDate != null) {
                _outputCode = _outputCode.Replace("%Y", DateTime.Parse(StrRegDate).ToString("yy"))
                                            .Replace("%MM", DateTime.Parse(StrRegDate).ToString("MM"));
            }
            _outputCode = _outputCode.Replace("%T", StrInitial)
                                        .Replace("%R", StrRevision)
                                        .Replace("%M", _monthCode[^1..])
                                        .Replace("%S", _serialCode);

            return _outputCode;
        }
        private Bitmap MakeLabelImage(string text, int resolution, int magnitude) {
            decimal _sizeX;
            decimal _sizeY;
            decimal _fontSize;
            float _stringPosX;
            float _stringPosY;
            Font _fnt;
            Bitmap _labelImage = new(0, 0);
            Graphics _g;
            SizeF _stringSize;
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
                    int _barWeight;
                    float _barCodePosX;
                    float _barCodePosY;
                    int _barcodeHeight = (int)(SettingsBarcodePro.BarcodeProLabelSettings.BarcodeHeight / 25.4F * resolution * magnitude);
                    _sizeX = (decimal)SettingsBarcodePro.BarcodeProPageSettings.SizeX / 25.4M * resolution * magnitude;
                    _sizeY = (decimal)SettingsBarcodePro.BarcodeProPageSettings.SizeY / 25.4M * resolution * magnitude;
                    _fontSize = (decimal)SettingsBarcodePro.BarcodeProLabelSettings.Font.SizeInPoints / 72.0M * resolution * magnitude;
                    _fnt = new(SettingsBarcodePro.BarcodeProLabelSettings.Font.Name, (float)_fontSize);

                    _labelImage = new((int)_sizeX, (int)_sizeY);
                    _g = Graphics.FromImage(_labelImage);

                    if (resolution == DisplayResolution) { _barWeight = 1; }
                    else { _barWeight = (int)(1 * resolution / DisplayResolution / DisplayMagnitude); }

                    Image _img = Code128Rendering.MakeBarcodeImage(text, _barWeight, true);
                    decimal _imageWidth = (decimal)(_img.Width * SettingsBarcodePro.BarcodeProLabelSettings.BarcodeMagnitude);

                    if (_imageWidth > _labelImage.Width) { MessageBox.Show($"バーコードの幅がラベル幅を超えています{_imageWidth}>{_labelImage.Width}"); }

                    _stringSize = _g.MeasureString(text, _fnt);

                    if (SettingsBarcodePro.BarcodeProLabelSettings.AlignStringCenter) { _stringPosX = (float)((_labelImage.Width / 2) - (_stringSize.Width / 2)); }
                    else { _stringPosX = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4M * resolution * magnitude); }

                    _stringPosY = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosY / 25.4M * resolution * magnitude);

                    if (SettingsBarcodePro.BarcodeProLabelSettings.AlignBarcodeCenter) { _barCodePosX = (float)((_labelImage.Width / 2) - (_imageWidth / 2)); }
                    else { _barCodePosX = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4M * resolution * magnitude); }

                    _barCodePosY = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.BarcodePosY / 25.4M * resolution * magnitude);

                    _g.DrawString(text, _fnt, Brushes.Black, _stringPosX, _stringPosY);
                    _g.DrawImage(_img, _barCodePosX, _barCodePosY, (float)_imageWidth, _barcodeHeight);

                    _g.Dispose();
                    _img.Dispose();
                    break;
            }
            return _labelImage;
        }

        private bool Registration() {
            int _dgvRowCnt;

            if (IntRegType > 0) {
                using SQLiteConnection con = new(MainWindow.GetConnectionString2());
                con.Open();

                foreach (var b in StrSerial) {
                    using SQLiteCommand cmd = con.CreateCommand();
                    cmd.CommandText =
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

                    cmd.Parameters.Add("@col_Serial", DbType.String).Value = b;
                    cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                    cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                    cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                    cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                    cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;

                    cmd.ExecuteNonQuery();
                }
            }

            switch (IntRegType) {
                case 0:
                    using (SQLiteConnection con = new(MainWindow.GetConnectionString2())) {
                        con.Open();

                        using SQLiteCommand cmd = con.CreateCommand();
                        cmd.CommandText =
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

                        cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                        cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                        cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                        cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                        cmd.Parameters.Add("@col_Quantity", DbType.String).Value = IntQuantity;
                        cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                        cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                        cmd.Parameters.Add("@col_Revision", DbType.String).Value = StrRevision;
                        cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                        cmd.ExecuteNonQuery();
                    }
                    break;

                case 1:
                    using (SQLiteConnection con = new(MainWindow.GetConnectionString2())) {
                        con.Open();

                        using SQLiteCommand cmd = con.CreateCommand();
                        cmd.CommandText =
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

                        cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                        cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                        cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                        cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                        cmd.Parameters.Add("@col_Quantity", DbType.String).Value = IntQuantity;
                        cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                        cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                        cmd.Parameters.Add("@col_Revision", DbType.String).Value = StrRevision;
                        cmd.Parameters.Add("@col_Serial_First", DbType.String).Value = StrSerialFirstNumber;
                        cmd.Parameters.Add("@col_Serial_Last", DbType.String).Value = StrSerialLastNumber;
                        cmd.Parameters.Add("@col_Serial_LastNum", DbType.String).Value = IntSerialLastNumber;
                        cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                        cmd.ExecuteNonQuery();
                    }
                    break;

                case 2:
                case 3:
                    using (SQLiteConnection con = new(MainWindow.GetConnectionString2())) {
                        con.Open();

                        string Substrate_Name = string.Empty;
                        string Substrate_Model = string.Empty;
                        if (ArrUseSubstrate == null) { throw new Exception("ArrUseSubstrateがnullです。"); }
                        for (int i = 0; i <= ArrUseSubstrate.Length; i++) {

                            CheckBox? objCbx = Controls[CheckBoxNames[i]] as CheckBox ?? throw new Exception("objCbxがnullです。");

                            if (objCbx.Checked) {
                                DataGridView? objDgv = Controls[DataGridViewNames[i]] as DataGridView ?? throw new Exception("objCbxがnullです。");
                                _dgvRowCnt = objDgv.Rows.Count;

                                for (int j = 0; j <= _dgvRowCnt - 1; j++) {
                                    if (Convert.ToBoolean(objDgv.Rows[j].Cells[3].Value)) {
                                        using (SQLiteCommand cmd = con.CreateCommand()) {
                                            cmd.CommandText =
                                                $"UPDATE 'Stock_{StrStockName}'SET " +
                                                $"col_Flg = @col_Flg," +
                                                $"col_Stock = @col_Stock," +
                                                $"col_History = ifnull(col_History,'')|| @col_History " +
                                                $"WHERE " +
                                                $"col_Substrate_Num = '{objDgv.Rows[j].Cells[0].Value}'";

                                            if (Convert.ToInt32(objDgv.Rows[j].Cells[1].Value) - Convert.ToInt32(objDgv.Rows[j].Cells[2].Value) == 0) {
                                                cmd.Parameters.Add("@col_Flg", DbType.String).Value = 0;
                                            }
                                            else {
                                                cmd.Parameters.Add("@col_Flg", DbType.String).Value = 1;
                                            }

                                            cmd.Parameters.Add("@col_Stock", DbType.String).Value = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value) - Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);
                                            cmd.Parameters.Add("@col_History", DbType.String).Value = $"{StrProductNumber}({objDgv.Rows[j].Cells[2].Value}),";

                                            cmd.ExecuteNonQuery();

                                            cmd.CommandText = $@"SELECT * FROM ""Stock_{StrStockName}"" WHERE col_Substrate_Model = ""{ArrUseSubstrate[i]}"" ORDER BY _rowid_ ASC";
                                            using SQLiteDataReader dr = cmd.ExecuteReader();
                                            while (dr.Read()) {
                                                Substrate_Name = $"{dr["col_Substrate_Name"]}";
                                                Substrate_Model = $"{dr["col_Substrate_Model"]}";
                                            }
                                        }

                                        if (string.IsNullOrEmpty(StrTotalSubstrate)) {
                                            StrTotalSubstrate = $"[{ArrUseSubstrate[i]}]{objDgv.Rows[j].Cells[0].Value}({objDgv.Rows[j].Cells[2].Value})";
                                        }
                                        else {
                                            StrTotalSubstrate = $"{StrTotalSubstrate},[{ArrUseSubstrate[i]}]{objDgv.Rows[j].Cells[0].Value}({objDgv.Rows[j].Cells[2].Value})";
                                        }

                                        using (SQLiteCommand cmd = con.CreateCommand()) {
                                            cmd.CommandText =
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

                                            cmd.Parameters.Add("@col_Substrate_Name", DbType.String).Value = Substrate_Name;
                                            cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = Substrate_Model;
                                            cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = objDgv.Rows[j].Cells[0].Value;
                                            cmd.Parameters.Add("@col_Decrease", DbType.String).Value = 0 - Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);
                                            cmd.Parameters.Add("@col_Use_P_Type", DbType.String).Value = StrProductType;
                                            cmd.Parameters.Add("@col_Use_P_Num", DbType.String).Value = StrProductNumber;
                                            cmd.Parameters.Add("@col_Use_O_Num", DbType.String).Value = StrOrderNumber;
                                            cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                                            cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                                            cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                                            cmd.ExecuteNonQuery();
                                        }

                                        switch (IntPrintType) {
                                            case 5:
                                            case 6:
                                                ListUsedSubstrate.Add(ArrUseSubstrate[i]);
                                                var _cellValue = objDgv.Rows[j].Cells[0].Value?.ToString();
                                                if (_cellValue != null) { ListUsedProductNumber.Add(_cellValue); }
                                                ListUsedQuantity.Add(Convert.ToInt32(objDgv.Rows[j].Cells[2].Value));
                                                break;
                                        }
                                    }
                                }
                            }
                        }

                        using (SQLiteCommand cmd = con.CreateCommand()) {
                            cmd.CommandText =
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

                            cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                            cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                            cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                            cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                            cmd.Parameters.Add("@col_Quantity", DbType.String).Value = IntQuantity;
                            cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                            cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                            cmd.Parameters.Add("@col_Revision", DbType.String).Value = StrRevision;
                            cmd.Parameters.Add("@col_Serial_First", DbType.String).Value = StrSerialFirstNumber;
                            cmd.Parameters.Add("@col_Serial_Last", DbType.String).Value = StrSerialLastNumber;
                            cmd.Parameters.Add("@col_Serial_LastNum", DbType.String).Value = IntSerialLastNumber;
                            cmd.Parameters.Add("@col_Use_Substrate", DbType.String).Value = StrTotalSubstrate;
                            cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

                            cmd.ExecuteNonQuery();
                        }
                    }
                    break;
            }
            return true;
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e) {
            CheckBox checkBox = (CheckBox)sender;
            DataGridView? dataGridView = null;

            switch (checkBox.Name) {
                case "Substrate1CheckBox": dataGridView = Substrate1DataGridView; break;
                case "Substrate2CheckBox": dataGridView = Substrate2DataGridView; break;
                case "Substrate3CheckBox": dataGridView = Substrate3DataGridView; break;
                case "Substrate4CheckBox": dataGridView = Substrate4DataGridView; break;
                case "Substrate5CheckBox": dataGridView = Substrate5DataGridView; break;
                case "Substrate6CheckBox": dataGridView = Substrate6DataGridView; break;
                case "Substrate7CheckBox": dataGridView = Substrate7DataGridView; break;
                case "Substrate8CheckBox": dataGridView = Substrate8DataGridView; break;
                case "Substrate9CheckBox": dataGridView = Substrate9DataGridView; break;
                case "Substrate10CheckBox": dataGridView = Substrate10DataGridView; break;
            }

            if (dataGridView != null) {
                dataGridView.Enabled = checkBox.Checked;
                checkBox.ForeColor = checkBox.Checked ? Color.Black : Color.Red;

                if (!checkBox.Checked) {
                    MessageBox.Show("チェックがない場合在庫から引き落とされなくなります。", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }

        private void FormClose() {
            ProductRegistration1Window productRegistration1Window = new();
            productRegistration1Window.Close();
            Close();

        }

        private void ProductRegistration2Window_Load(object sender, EventArgs e) { LoadEvents(); }
        private void ProductRegistration2Window_FormClosing(object sender, FormClosingEventArgs e) { ClosingEvents(); }
        private void RegisterButton_Click(object sender, EventArgs e) { RegisterCheck(); }
        private void CloseButton_Click(object sender, EventArgs e) { FormClose(); }
        private void SubstrateCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBox_CheckedChanged(sender, e); }
        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e) { FormClose(); }

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
            ProductPrintSetting ls = new();
            ls.ShowDialog(this);
        }

        private void バーコード印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductBarcodePrintSettingsWindow ls = new();
            ls.ShowDialog(this);
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
                $"StrPerson\t\t[{StrPerson}]\r\n" +
                $"IntQuantity\t\t[{IntQuantity}]\r\n" +
                $"IntSerialFirstNumber\t\t[{IntSerialFirstNumber}]\r\n" +
                $"IntSerialLastNumber\t\t[{IntSerialLastNumber}]\r\n" +
                $"StrInitial\t\t\t[{StrInitial}]\r\n" +
                $"IntRegType\t\t[{IntRegType}]\r\n" +
                $"IntPrintType\t\t[{IntPrintType}]\r\n" +
                $"IntSerialDigit\t\t[{IntSerialDigit}]" +
                $"", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

    }
}
