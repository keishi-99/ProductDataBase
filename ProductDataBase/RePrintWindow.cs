using GenCode128;
using LabelGenerator;
using System.Data;
using System.Data.SQLite;

namespace ProductDatabase {
    public partial class RePrintWindow : Form {

        public CSettingsLabelPro SettingsLabelPro { get; set; } = new CSettingsLabelPro();
        public string StrLabelSettingFilePath { get; set; } = string.Empty;

        public CSettingsBarcodePro SettingsBarcodePro { get; set; } = new CSettingsBarcodePro();
        public string StrBarcodeSettingFilePath { get; set; } = string.Empty;

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public string StrProductName { get; set; } = string.Empty;
        public string StrProductType { get; set; } = string.Empty;
        public string StrProductModel { get; set; } = string.Empty;
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
        public int IntSerialLastNumber { get; set; }

        public Bitmap LabelProBmp { get; } = new(1, 1);
        public int LabelProPageNum { get; set; }
        public int LabelProNSerial { get; set; }
        public int LabelProNLabel { get; set; }
        public int LabelProNumLabelsToPrint { get; set; }

        public decimal DisplayResolution { get; } = 96.0m;
        public int DisplayMagnitude { get; } = 3;
        public int IntPageCnt { get; set; } = 1;

        private string StrSerialType = string.Empty;
        private string StrSerialFirstNumber = string.Empty;
        private string StrSerialLastNumber = string.Empty;
        private bool FontUnderbar = false;

        readonly List<string> checkBoxNames = new() {
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "ExtraCheckBox1",
                    "RevisionCheckBox", "ExtraCheckBox2", "ExtraCheckBox3", "FirstSerialNumberCheckBox", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" };

        public RePrintWindow() => InitializeComponent();

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(StrFontName, IntFontSize);

                ProductNameLabel2.Text = StrProductName;
                SubstrateModelLabel2.Text = $"{StrProductName} - {StrProductModel}";

                FirstSerialNumberTextBox.MaxLength = IntSerialDigit;

                switch (IntPrintType) {
                    case 0:
                        LabelPrintButton.Enabled = false;
                        BarcodePrintButton.Enabled = false;
                        break;
                    case 1:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                        SettingsLabelPro = new CSettingsLabelPro();
                        StrLabelSettingFilePath = $"./config/{StrProductName}/SerialConfig_{StrProductName}_{StrProductModel}.xml";
                        LabelPrintButton.Enabled = true;
                        BarcodePrintButton.Enabled = false;
                        break;
                    case 2:
                        SettingsBarcodePro = new CSettingsBarcodePro();
                        StrBarcodeSettingFilePath = $"./config/{StrProductName}/BarcodeConfig_{StrProductName}_{StrProductModel}.xml";
                        LabelPrintButton.Enabled = false;
                        BarcodePrintButton.Enabled = true;
                        break;
                    case 3:
                        SettingsLabelPro = new CSettingsLabelPro();
                        StrLabelSettingFilePath = $"./config/{StrProductName}/SerialConfig_{StrProductName}_{StrProductModel}.xml";
                        SettingsBarcodePro = new CSettingsBarcodePro();
                        StrBarcodeSettingFilePath = $"./config/{StrProductName}/BarcodeConfig_{StrProductName}_{StrProductModel}.xml";
                        LabelPrintButton.Enabled = true;
                        BarcodePrintButton.Enabled = true;
                        break;
                    case 8:
                        SettingsLabelPro = new CSettingsLabelPro();
                        StrLabelSettingFilePath = $"./config/{StrProductName}/SerialConfig_{StrProductName}_{StrProductModel}.xml";
                        LabelPrintButton.Enabled = false;
                        BarcodePrintButton.Enabled = false;
                        break;
                    default:
                        break;
                }

                // TextBoxへ今日の年月日を入力
                DateTime _dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = _dtNow.ToShortDateString();

                // DB1へ接続し担当者取得
                using (SQLiteConnection _con = new(MainWindow.GetConnectionString1())) {
                    _con.Open();
                    using SQLiteCommand _cmd = _con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    _cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using SQLiteDataReader _dr = _cmd.ExecuteReader();
                    while (_dr.Read()) {
                        PersonComboBox.Items.Add($"{_dr["col_Person_Name"]}");
                    }
                }

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                    _con.Open();
                    using SQLiteCommand _cmd = _con.CreateCommand();
                    // テーブル検索SQL - [Product_Name]_stockテーブルの[col_Substrate_Model]列の[col_Revision]を取得
                    _cmd.CommandText = $"SELECT col_Revision FROM 'Product_Reg_{StrProductName}' ORDER BY _rowid_ DESC";
                    object _result = _cmd.ExecuteScalar();
                    RevisionTextBox.Text = _result?.ToString() ?? "";
                }

                // 変数[check_bin]の値に応じてCheckboxにチェックを入れる
                foreach (string checkBoxName in checkBoxNames) {
                    if (Controls[checkBoxName] is CheckBox checkBox) {
                        checkBox.Checked = (IntCheckBin & 0x1) == 1;
                        IntCheckBin >>= 1;
                    }
                }

                FirstSerialNumberCheckBox.Checked = true;
                LoadSettings(StrLabelSettingFilePath, StrBarcodeSettingFilePath);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadSettings(string strLabelSettingFilePath, string strBarcodeSettingFilePath) {
            try {
                if (strLabelSettingFilePath != string.Empty) {
                    StreamReader? _srLabel = new(strLabelSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer _serializerLabel = new(typeof(CSettingsLabelPro));
                    if (_serializerLabel.Deserialize(_srLabel) is CSettingsLabelPro _result) { SettingsLabelPro = _result; }
                    _srLabel?.Close();
                }
                if (strBarcodeSettingFilePath != string.Empty) {
                    StreamReader? _srBarcode = new(strBarcodeSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer _serializerBarcode = new(typeof(CSettingsBarcodePro));
                    if (_serializerBarcode.Deserialize(_srBarcode) is CSettingsBarcodePro _result) { SettingsBarcodePro = _result; }
                    _srBarcode?.Close();
                }
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // 登録処理
        private void Registeration() {
            try {
                // 入力フォームのチェック
                bool _anyTextBoxEnabled = false;
                bool _allTextBoxesFilled = true;

                foreach (Control control in Controls) {
                    if (control is TextBoxBase textBox && textBox.Enabled) {
                        _anyTextBoxEnabled = true;
                        if (string.IsNullOrWhiteSpace(textBox.Text)) {
                            _allTextBoxesFilled = false;
                            break;
                        }
                    }
                }
                if (!_anyTextBoxEnabled) { throw new Exception("何も入力されていません"); }
                if (!_allTextBoxesFilled) { throw new Exception("空欄があります。"); }

                if (ManufacturingNumberCheckBox.Checked && ManufacturingNumberMaskedTextBox.Text.Length != 15) { throw new Exception("製番を10桁+4桁で入力して下さい。"); }

                if (QuantityCheckBox.Checked && int.Parse(QuantityTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

                DialogResult _result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (_result == DialogResult.Cancel) return;

                int _quantity = Convert.ToInt32(QuantityTextBox.Text);
                if (_quantity == 0) { throw new Exception("1以上入力してください。"); }
                int _firstSerial = Convert.ToInt32(FirstSerialNumberTextBox.Text);
                if (_firstSerial == 0) { throw new Exception("シリアル開始番号を入力してください。"); }

                switch (IntSerialDigit) {
                    case 3:
                        CheckAndAdjustSerial(999, 1);
                        break;
                    case 4:
                        CheckAndAdjustSerial(9999, 1);
                        break;
                    default:
                        break;
                }

                void CheckAndAdjustSerial(int threshold, int resetValue) {
                    if (_quantity + _firstSerial >= threshold) {
                        MessageBox.Show($"シリアルが{threshold}を超えるので{resetValue.ToString().PadLeft(IntSerialDigit, '0')}から開始します。");
                        FirstSerialNumberTextBox.Text = resetValue.ToString();
                    }
                }

                DialogResult result = MessageBox.Show("同一のシリアルラベルが複数存在しないようにして下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) { return; }

                StrOrderNumber = OrderNumberTextBox.Text;
                StrProductNumber = ManufacturingNumberMaskedTextBox.Text;
                IntQuantity = Convert.ToInt32(QuantityTextBox.Text);
                StrPerson = PersonComboBox.Text;
                StrRegDate = RegistrationDateMaskedTextBox.Text;
                StrRevision = RevisionTextBox.Text;
                StrComment = CommentTextBox.Text;

                IntSerialFirstNumber = Convert.ToInt32(FirstSerialNumberTextBox.Text);
                IntSerialLastNumber = IntSerialFirstNumber + IntQuantity - 1;

                StrSerialFirstNumber = GenerateCode(IntSerialFirstNumber);
                StrSerialLastNumber = GenerateCode(IntSerialLastNumber);

                PrintBarcode(1);

                // 基板登録テーブルへ追加
                using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
                _con.Open();
                using SQLiteCommand _cmd = _con.CreateCommand();
                _cmd.CommandText =
                    $"INSERT INTO 'Reprint'" +
                    $"(col_Print_Type, col_Order_Num, col_Product_Num, col_Product_Type, col_Product_Model, col_Quantity, col_Person, col_RegDate, col_Revision, col_Serial_First, col_Serial_Last, col_Comment)" +
                    $"VALUES" +
                    $"(@col_Print_Type, @col_Order_Num, @col_Product_Num, @col_Product_Type, @col_Product_Model, @col_Quantity, @col_Person, @col_RegDate, @col_Revision, @col_Serial_First, @col_Serial_Last, @col_Comment)";

                // チェックボックスにチェックがない場合はNullを
                _cmd.Parameters.Add("@col_Print_Type", DbType.AnsiString).Value = IntPrintType;
                _cmd.Parameters.Add("@col_Order_Num", DbType.AnsiString).Value = StrOrderNumber;
                _cmd.Parameters.Add("@col_Product_Num", DbType.AnsiString).Value = StrProductNumber;
                _cmd.Parameters.Add("@col_Product_Type", DbType.AnsiString).Value = StrProductType;
                _cmd.Parameters.Add("@col_Product_Model", DbType.AnsiString).Value = StrProductModel;
                _cmd.Parameters.Add("@col_Quantity", DbType.AnsiString).Value = IntQuantity;
                _cmd.Parameters.Add("@col_Person", DbType.AnsiString).Value = StrPerson;
                _cmd.Parameters.Add("@col_RegDate", DbType.AnsiString).Value = StrRegDate;
                _cmd.Parameters.Add("@col_Revision", DbType.AnsiString).Value = StrRevision;
                _cmd.Parameters.Add("@col_Serial_First", DbType.AnsiString).Value = StrSerialFirstNumber;
                _cmd.Parameters.Add("@col_Serial_Last", DbType.AnsiString).Value = StrSerialLastNumber;
                _cmd.Parameters.Add("@col_Comment", DbType.AnsiString).Value = StrComment;

                _cmd.ExecuteNonQuery();

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    RePrintPrintDialog.Document = _pd;
                    DialogResult _r = RePrintPrintDialog.ShowDialog();

                    if (_r == DialogResult.OK) {
                        RePrintPrintDialog.Document.Print();

                        if (IntPageCnt >= 2) {
                            MessageBox.Show($"{IntPageCnt}枚印刷されます。2枚目以降は1行目から印刷されます。");
                        }
                    }
                    else {
                        return;
                    }
                    break;
                case 2:
                    RePrintPrintPreviewDialog.Document = _pd;
                    RePrintPrintPreviewDialog.ShowDialog();
                    break;
                default:
                    break;
            }
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                int _startLineLabel = (int)PrintPostionNumericUpDown.Value - 1;
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

                int _startLineBarcode = (int)PrintPostionNumericUpDown.Value - 1;

                Point _offset;
                if (!RePrintPrintDocument.PrintController.IsPreview) {
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
            decimal _sizeX;
            decimal _sizeY;
            decimal _fontSize;
            float _stringPosX;
            float _stringPosY;
            Font _fnt;
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

                    _sizeX = (decimal)SettingsBarcodePro.BarcodeProPageSettings.SizeX / 25.4M * resolution * magnitude;
                    _sizeY = (decimal)SettingsBarcodePro.BarcodeProPageSettings.SizeY / 25.4M * resolution * magnitude;
                    _fontSize = (decimal)SettingsBarcodePro.BarcodeProLabelSettings.Font.SizeInPoints / 72.0M * resolution * magnitude;
                    _fnt = new(SettingsBarcodePro.BarcodeProLabelSettings.Font.Name, (float)_fontSize);

                    _labelImage = new((int)_sizeX, (int)_sizeY);
                    _g = Graphics.FromImage(_labelImage);

                    int _barWeight;
                    if (resolution == DisplayResolution) { _barWeight = 1; }
                    else { _barWeight = (int)(1 * resolution / DisplayResolution / DisplayMagnitude); }

                    Image _img = Code128Rendering.MakeBarcodeImage(text, _barWeight, true);
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
                    break;
                default:
                    break;
            }
            return _labelImage;
        }
        // チェックボックスイベント
        private void CheckBoxChecked(object sender, EventArgs e) {
            using CheckBox _checkBox = (CheckBox)sender;

            switch (_checkBox.Name) {
                case "OrderNumberCheckBox":
                    OrderNumberTextBox.Enabled = _checkBox.Checked;
                    break;
                case "ManufacturingNumberCheckBox":
                    ManufacturingNumberMaskedTextBox.Enabled = _checkBox.Checked;
                    break;
                case "QuantityCheckBox":
                    QuantityTextBox.Enabled = _checkBox.Checked;
                    if (_checkBox.Checked) ExtraCheckBox1.Checked = false;
                    break;
                case "ExtraCheckBox1":
                    ExtraTextBox2.Enabled = _checkBox.Checked;
                    break;
                case "RevisionCheckBox":
                    RevisionTextBox.Enabled = _checkBox.Checked;
                    if (_checkBox.Checked) MessageBox.Show("変更する場合は理由を記載して下さい。");
                    break;
                case "ExtraCheckBox2":
                    ExtraTextBox3.Enabled = _checkBox.Checked;
                    break;
                case "ExtraCheckBox3":
                    ExtraTextBox3.Enabled = _checkBox.Checked;
                    break;
                case "FirstSerialNumberCheckBox":
                    FirstSerialNumberTextBox.Enabled = _checkBox.Checked;
                    break;
                case "RegistrationDateCheckBox":
                    RegistrationDateMaskedTextBox.Enabled = _checkBox.Checked;
                    break;
                case "PersonCheckBox":
                    PersonComboBox.Enabled = _checkBox.Checked;
                    break;
                case "ExtraCheckBox4":
                    ExtraTextBox4.Enabled = _checkBox.Checked;
                    break;
                case "ExtraCheckBox5":
                    ExtraTextBox5.Enabled = _checkBox.Checked;
                    break;
                case "ExtraCheckBox6":
                    ExtraTextBox6.Enabled = _checkBox.Checked;
                    break;
                case "CommentCheckBox":
                    CommentTextBox.Enabled = _checkBox.Checked;
                    CommentComboBox.Enabled = _checkBox.Checked;
                    TemplateButton.Enabled = _checkBox.Checked;
                    break;
            }
        }
        // 入力数値のみ
        private void NumericOnly(object sender, KeyPressEventArgs e) {
            // 0～9と、バックスペース以外の時は、イベントをキャンセルする
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b') {
                e.Handled = true;
            }
        }
        // 日付チェック
        private void RegistrationDateCheck(object sender, TypeValidationEventArgs e) {
            if (!e.IsValidInput) {
                MessageBox.Show("日付が正しくありません。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                RegistrationDateMaskedTextBox.Focus();
            }
        }


        private void RePrintWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void LabelPrintButton_Click(object sender, EventArgs e) {
            StrSerialType = "Label";
            Registeration();
        }
        private void BarcodePrintButton_Click(object sender, EventArgs e) {
            StrSerialType = "Barcode";
            Registeration();
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            MessageBox.Show($"" +
                $"StrProductName\t\t[{StrProductName}]\r\n" +
                $"StrProductModel\t\t[{StrProductModel}]\r\n" +
                $"StrProductType\t\t[{StrProductType}]\r\n" +
                $"StrOrderNumber\t\t[{StrOrderNumber}]\r\n" +
                $"StrProductNumber\t\t[{StrProductNumber}]\r\n" +
                $"StrRevision\t\t[{StrRevision}]\r\n" +
                $"IntRegType\t\t[{IntRegType}]\r\n" +
                $"StrRegDate\t\t[{StrRegDate}]\r\n" +
                $"StrPerson\t\t\t[{StrPerson}]\r\n" +
                $"IntQuantity\t\t[{IntQuantity}]\r\n" +
                $"IntSerialFirstNumber\t[{IntSerialFirstNumber}]\r\n" +
                $"IntSerialLastNumber\t[{IntSerialLastNumber}]\r\n" +
                $"StrInitial\t\t\t[{StrInitial}]\r\n" +
                $"IntPrintType\t\t[{IntPrintType}]\r\n" +
                $"IntSerialDigit\t\t[{IntSerialDigit}]" +
                $"", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
