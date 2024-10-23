using GenCode128;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class RePrintWindow : Form {

        public CSettingsLabelPro SettingsLabelPro { get; set; } = new CSettingsLabelPro();
        private string _labelSettingFilePath = String.Empty;

        public CSettingsBarcodePro SettingsBarcodePro { get; set; } = new CSettingsBarcodePro();
        private string _barcodeSettingFilePath = String.Empty;

        public ProductInfomation ProductInfo { get; set; } = new ProductInfomation();

        private string _orderNumber = String.Empty;
        private string _productNumber = String.Empty;
        private string _regDate = String.Empty;
        private string _person = String.Empty;
        private string _revision = String.Empty;
        private string _comment = String.Empty;

        private int _quantity;
        private int _serialFirstNumber;
        private int _serialLastNumber;

        private int _labelProPageNum;
        private int _labelProNSerial;
        private int _labelProNumLabelsToPrint;

        private readonly decimal _displayResolution = 96.0m;
        private readonly int _displayMagnitude = 3;
        private int _intPageCnt = 1;

        private string _strSerialType = String.Empty;
        private string _strSerialFirstNumber = String.Empty;
        private string _strSerialLastNumber = String.Empty;
        private bool _fontUnderbar = false;
        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "ExtraCheckBox1",
                    "RevisionCheckBox", "ExtraCheckBox2", "ExtraCheckBox3", "FirstSerialNumberCheckBox", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" ];

        public RePrintWindow() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);

                ProductNameLabel2.Text = ProductInfo.ProductName;
                SubstrateModelLabel2.Text = $"{ProductInfo.ProductName} - {ProductInfo.ProductModel}";

                FirstSerialNumberTextBox.MaxLength = ProductInfo.SerialDigit;

                // 変数[check_bin]のビットに応じてCheckboxにチェックを入れる
                for (var i = 0; i < _checkBoxNames.Count; i++) {
                    if (Controls[_checkBoxNames[i]] is CheckBox checkBox) {
                        // i番目のビットが1かどうかをチェック
                        checkBox.Checked = (ProductInfo.CheckBin & (1 << i)) != 0;
                    }
                }

                switch (ProductInfo.PrintType) {
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
                        _labelSettingFilePath = $"./config/{ProductInfo.ProductName}/SerialConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml";
                        LabelPrintButton.Enabled = true;
                        BarcodePrintButton.Enabled = false;
                        break;
                    case 2:
                        SettingsBarcodePro = new CSettingsBarcodePro();
                        _barcodeSettingFilePath = $"./config/{ProductInfo.ProductName}/BarcodeConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml";
                        LabelPrintButton.Enabled = false;
                        BarcodePrintButton.Enabled = true;
                        break;
                    case 3:
                        SettingsLabelPro = new CSettingsLabelPro();
                        _labelSettingFilePath = $"./config/{ProductInfo.ProductName}/SerialConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml";
                        SettingsBarcodePro = new CSettingsBarcodePro();
                        _barcodeSettingFilePath = $"./config/{ProductInfo.ProductName}/BarcodeConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml";
                        LabelPrintButton.Enabled = true;
                        BarcodePrintButton.Enabled = true;
                        break;
                    case 8:
                        SettingsLabelPro = new CSettingsLabelPro();
                        _labelSettingFilePath = $"./config/{ProductInfo.ProductName}/SerialConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml";
                        LabelPrintButton.Enabled = false;
                        BarcodePrintButton.Enabled = false;
                        break;
                    default:
                        break;
                }

                // TextBoxへ今日の年月日を入力
                var dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = dtNow.ToShortDateString();

                // DB1へ接続し担当者取得
                using (SQLiteConnection con = new(GetConnectionString1())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    cmd.CommandText = """SELECT * FROM Person ORDER BY _rowid_ ASC""";
                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        PersonComboBox.Items.Add($"{dr["col_Person_Name"]}");
                    }
                }

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using (SQLiteConnection con = new(GetConnectionString2())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - [Product_Name]_stockテーブルの[col_Substrate_Model]列の[col_Revision]を取得
                    cmd.CommandText = $"""SELECT col_Revision FROM "Product_Reg_{ProductInfo.ProductName}" ORDER BY _rowid_ DESC""";
                    var result = cmd.ExecuteScalar();
                    RevisionTextBox.Text = result?.ToString() ?? "";
                }

                FirstSerialNumberCheckBox.Checked = true;
                LoadSettings(_labelSettingFilePath, _barcodeSettingFilePath);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadSettings(string strLabelSettingFilePath, string strBarcodeSettingFilePath) {
            try {
                if (strLabelSettingFilePath != String.Empty) {
                    StreamReader? srLabel = new(strLabelSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer serializerLabel = new(typeof(CSettingsLabelPro));
                    if (serializerLabel.Deserialize(srLabel) is CSettingsLabelPro result) { SettingsLabelPro = result; }
                    srLabel?.Close();
                }
                if (strBarcodeSettingFilePath != String.Empty) {
                    StreamReader? srBarcode = new(strBarcodeSettingFilePath, new System.Text.UTF8Encoding(false));
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
        private void Registeration() {
            try {
                // 入力フォームのチェック
                var anyTextBoxEnabled = false;
                var allTextBoxesFilled = true;
                DialogResult result;

                foreach (Control control in Controls) {
                    if (control is TextBoxBase textBox && textBox.Enabled) {
                        anyTextBoxEnabled = true;
                        if (String.IsNullOrWhiteSpace(textBox.Text)) {
                            allTextBoxesFilled = false;
                            break;
                        }
                    }
                }
                if (!anyTextBoxEnabled) { throw new Exception("何も入力されていません"); }
                if (!allTextBoxesFilled) { throw new Exception("空欄があります。"); }

                if (ManufacturingNumberCheckBox.Checked && ManufacturingNumberMaskedTextBox.Text.Length != 15) { throw new Exception("製番を10桁+4桁で入力して下さい。"); }

                if (QuantityCheckBox.Checked && Int32.Parse(QuantityTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

                result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) {
                    return;
                }

                var quantity = Convert.ToInt32(QuantityTextBox.Text);
                if (quantity == 0) { throw new Exception("1以上入力してください。"); }
                var firstSerial = Convert.ToInt32(FirstSerialNumberTextBox.Text);
                if (firstSerial == 0) { throw new Exception("シリアル開始番号を入力してください。"); }

                switch (ProductInfo.SerialDigit) {
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
                    if (quantity + firstSerial >= threshold) {
                        MessageBox.Show($"シリアルが{threshold}を超えるので{resetValue.ToString().PadLeft(ProductInfo.SerialDigit, '0')}から開始します。");
                        FirstSerialNumberTextBox.Text = resetValue.ToString();
                    }
                }

                result = MessageBox.Show("同一のシリアルラベルが複数存在しないようにして下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) { return; }

                _orderNumber = OrderNumberTextBox.Text;
                _productNumber = ManufacturingNumberMaskedTextBox.Text;
                _quantity = Convert.ToInt32(QuantityTextBox.Text ?? throw new Exception());
                _person = PersonComboBox.Text;
                _regDate = RegistrationDateMaskedTextBox.Text;
                _revision = RevisionTextBox.Text;
                _comment = CommentTextBox.Text;

                _serialFirstNumber = Convert.ToInt32(FirstSerialNumberTextBox.Text);
                _serialLastNumber = _serialFirstNumber + _quantity - 1;

                _strSerialFirstNumber = GenerateCode(_serialFirstNumber);
                _strSerialLastNumber = GenerateCode(_serialLastNumber);

                if (!PrintBarcode(1)) { throw new Exception("キャンセルしました。"); }

                // 再印刷登録テーブルへ追加
                using SQLiteConnection con = new(GetConnectionString2());
                con.Open();
                using var cmd = con.CreateCommand();
                cmd.CommandText =
                    $"""
                    INSERT INTO Reprint
                        (col_Print_Type, col_Order_Num, col_Product_Num, col_Product_Type, col_Product_Model, col_Quantity, col_Person, col_RegDate, col_Revision, col_Serial_First, col_Serial_Last, col_Comment)
                    VALUES
                        (@col_Print_Type, @col_Order_Num, @col_Product_Num, @col_Product_Type, @col_Product_Model, @col_Quantity, @col_Person, @col_RegDate, @col_Revision, @col_Serial_First, @col_Serial_Last, @col_Comment)
                    """;

                // チェックボックスにチェックがない場合はNullを
                cmd.Parameters.Add("@col_Print_Type", DbType.String).Value = ProductInfo.PrintType;
                cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = _orderNumber;
                cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = _productNumber;
                cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = ProductInfo.ProductType;
                cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = ProductInfo.ProductModel;
                cmd.Parameters.Add("@col_Quantity", DbType.String).Value = _quantity;
                cmd.Parameters.Add("@col_Person", DbType.String).Value = _person;
                cmd.Parameters.Add("@col_RegDate", DbType.String).Value = _regDate;
                cmd.Parameters.Add("@col_Revision", DbType.String).Value = _revision;
                cmd.Parameters.Add("@col_Serial_First", DbType.String).Value = _strSerialFirstNumber;
                cmd.Parameters.Add("@col_Serial_Last", DbType.String).Value = _strSerialLastNumber;
                cmd.Parameters.Add("@col_Comment", DbType.String).Value = _comment;

                cmd.ExecuteNonQuery();

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 印刷処理
        private bool PrintBarcode(int printFlg) {
            // PrintDocumentオブジェクトの作成
            using System.Drawing.Printing.PrintDocument pd = new();

            // PrintPageイベントハンドラの追加
            pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

            _labelProNumLabelsToPrint = _quantity;
            _labelProPageNum = 0;
            _fontUnderbar = false;

            switch (printFlg) {
                case 1:
                    RePrintPrintDialog.Document = pd;
                    var r = RePrintPrintDialog.ShowDialog();

                    if (r == DialogResult.OK) {
                        RePrintPrintDialog.Document.Print();

                        if (_intPageCnt >= 2) {
                            MessageBox.Show($"{_intPageCnt}枚印刷されます。2枚目以降は1行目から印刷されます。");
                        }
                    }
                    else {
                        return false;
                    }
                    break;
                case 2:
                    RePrintPrintPreviewDialog.Document = pd;
                    RePrintPrintPreviewDialog.ShowDialog();
                    break;
                default:
                    break;
            }
            return true;
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                var startLineLabel = (int)PrintPostionNumericUpDown.Value - 1;
                Point headerPos = new(0, 0);
                var headerString = String.Empty;
                Font headerFooterFont = new("Arial", 6);
                var intNumLabels = 0;
                var intCountNumLabels = 0;


                var maxX = 0;
                var maxY = 0;
                float sizeX = 0;
                float sizeY = 0;
                double offsetX = 0;
                double offsetY = 0;
                double intervalX = 0;
                double intervalY = 0;
                switch (_strSerialType) {
                    case "Label":
                        if (SettingsLabelPro == null) { throw new Exception("SettingsLabelProがnullです。"); }
                        maxX = SettingsLabelPro.LabelProPageSettings.NumLabelsX;
                        maxY = SettingsLabelPro.LabelProPageSettings.NumLabelsY;
                        sizeX = (float)SettingsLabelPro.LabelProPageSettings.SizeX;
                        sizeY = (float)SettingsLabelPro.LabelProPageSettings.SizeY;
                        offsetX = SettingsLabelPro.LabelProPageSettings.OffsetX;
                        offsetY = SettingsLabelPro.LabelProPageSettings.OffsetY;
                        intervalX = SettingsLabelPro.LabelProPageSettings.IntervalX;
                        intervalY = SettingsLabelPro.LabelProPageSettings.IntervalY;
                        headerPos = SettingsLabelPro.LabelProPageSettings.HeaderPos;
                        headerString = ConvertHeaderFooterString(SettingsLabelPro.LabelProPageSettings.HeaderString);
                        headerFooterFont = SettingsLabelPro.LabelProPageSettings.HeaderFooterFont;
                        intNumLabels = SettingsLabelPro.LabelProLabelSettings.NumLabels;
                        intCountNumLabels = SettingsLabelPro.LabelProLabelSettings.NumLabels;
                        break;
                    case "Barcode":
                        if (SettingsBarcodePro == null) { throw new Exception("SettingsBarcodeProがnullです。"); }
                        maxX = SettingsBarcodePro.BarcodeProPageSettings.NumLabelsX;
                        maxY = SettingsBarcodePro.BarcodeProPageSettings.NumLabelsY;
                        sizeX = (float)SettingsBarcodePro.BarcodeProPageSettings.SizeX;
                        sizeY = (float)SettingsBarcodePro.BarcodeProPageSettings.SizeY;
                        offsetX = SettingsBarcodePro.BarcodeProPageSettings.OffsetX;
                        offsetY = SettingsBarcodePro.BarcodeProPageSettings.OffsetY;
                        intervalX = SettingsBarcodePro.BarcodeProPageSettings.IntervalX;
                        intervalY = SettingsBarcodePro.BarcodeProPageSettings.IntervalY;
                        headerPos = SettingsBarcodePro.BarcodeProPageSettings.HeaderPos;
                        headerString = ConvertHeaderFooterString(SettingsBarcodePro.BarcodeProPageSettings.HeaderString);
                        headerFooterFont = SettingsBarcodePro.BarcodeProPageSettings.HeaderFooterFont;
                        intNumLabels = SettingsBarcodePro.BarcodeProLabelSettings.NumLabels;
                        intCountNumLabels = SettingsBarcodePro.BarcodeProLabelSettings.NumLabels;
                        break;
                    default:
                        break;
                }

                var startLineBarcode = (int)PrintPostionNumericUpDown.Value - 1;

                Point offset;
                if (!RePrintPrintDocument.PrintController.IsPreview) {
                    offsetX -= e.PageSettings.HardMarginX * 0.254;
                    offsetY -= e.PageSettings.HardMarginY * 0.254;
                    offset = _labelProPageNum == 0
                        ? new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (startLineBarcode * (intervalY + sizeY))))
                        : new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (intervalY + sizeY))));
                }
                else {
                    offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (intervalY + sizeY))));
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;

                headerPos.Offset(offset);
                e.Graphics.DrawString(headerString, headerFooterFont, Brushes.Black, headerPos);

                var barcodePageNum = 0;
                offset = barcodePageNum == 0
                    ? new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (startLineBarcode * (intervalY + sizeY))))
                    : new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (intervalY + sizeY))));

                if (_labelProPageNum == 0) { _labelProNSerial = _serialFirstNumber; }
                if (_labelProPageNum >= 1) { startLineBarcode = 0; }

                var y = 0;
                for (y = startLineBarcode; y < maxY; y++) {
                    var x = 0;
                    for (x = 0; x < maxX; x++) {
                        var s = GenerateCode(_labelProNSerial);
                        var posX = (float)(offsetX + (x * (intervalX + sizeX)));
                        var posY = (float)(offsetY + (y * (intervalY + sizeY)));
                        e.Graphics.DrawImage(MakeLabelImage(s, (int)e.Graphics.DpiX, 1), posX, posY, sizeX, sizeY);

                        _labelProNSerial++;
                        _labelProNumLabelsToPrint--;

                        if (_labelProNumLabelsToPrint <= 0) {
                            intCountNumLabels--;
                            if (intCountNumLabels <= 0) {
                                e.HasMorePages = false;
                                _labelProPageNum = 0;
                                _labelProNumLabelsToPrint = 0;
                                return;
                            }
                            else {
                                _labelProNSerial -= x + 1;
                                _labelProNumLabelsToPrint += x + 1;
                                break;
                            }
                        }

                        if (x >= maxX - 1) {
                            intCountNumLabels--;
                            if (intCountNumLabels <= 0) {
                                intCountNumLabels = intNumLabels;
                            }
                            else if (intCountNumLabels > 0) {
                                _labelProNumLabelsToPrint += x + 1;
                                break;
                            }
                        }
                    }
                }

                if (_labelProNumLabelsToPrint > 0) {
                    _labelProPageNum++;
                    _intPageCnt++;
                    e.HasMorePages = true;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private string ConvertHeaderFooterString(string s) {
            s = s.Replace("%P", ProductInfo.ProductName)
                 .Replace("%T", ProductInfo.ProductModel)
                 .Replace("%D", DateTime.Today.ToShortDateString())
                 .Replace("%M", _productNumber)
                 .Replace("%O", _orderNumber)
                 .Replace("%N", _quantity.ToString())
                 .Replace("%U", "");
            return s;
        }
        private string GenerateCode(int serialCode) {
            var monthCode = DateTime.Parse(_regDate).ToString("MM");

            monthCode = monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => monthCode
            };

            var outputCode = _strSerialType switch {
                "Label" => SettingsLabelPro.LabelProLabelSettings.Format,
                "Barcode" => SettingsBarcodePro.BarcodeProLabelSettings.Format,
                _ => String.Empty
            };

            outputCode = outputCode.Replace("%Y", DateTime.Parse(_regDate).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(_regDate).ToString("MM"))
                                    .Replace("%T", ProductInfo.Initial)
                                    .Replace("%R", _revision)
                                    .Replace("%M", monthCode[^1..])
                                    .Replace("%S", Convert.ToInt32(serialCode).ToString($"D{ProductInfo.SerialDigit}"));

            return outputCode;
        }
        private Bitmap MakeLabelImage(string text, int resolution, int magnitude) {
            Bitmap labelImage = new(1, 1);
            decimal sizeX;
            decimal sizeY;
            decimal fontSize;
            float stringPosX;
            float stringPosY;
            Font fnt;
            Graphics g;
            SizeF stringSize;
            switch (_strSerialType) {
                case "Label":
                    if (SettingsLabelPro == null) { throw new Exception("SettingsLabelProがnull"); }
                    sizeX = (decimal)SettingsLabelPro.LabelProPageSettings.SizeX / 25.4M * resolution * magnitude;
                    sizeY = (decimal)SettingsLabelPro.LabelProPageSettings.SizeY / 25.4M * resolution * magnitude;
                    fontSize = (decimal)SettingsLabelPro.LabelProLabelSettings.Font.SizeInPoints / 72.0M * resolution * magnitude;
                    stringPosY = (int)((decimal)SettingsLabelPro.LabelProLabelSettings.StringPosY / 25.4M * resolution * magnitude);
                    var style = _fontUnderbar ? FontStyle.Underline : FontStyle.Regular;
                    fnt = new Font(SettingsLabelPro.LabelProLabelSettings.Font.Name, (float)fontSize, style);

                    labelImage = new((int)sizeX, (int)sizeY);
                    g = Graphics.FromImage(labelImage);

                    stringSize = g.MeasureString(text, fnt);

                    stringPosX = (int)((labelImage.Width / 2) - (stringSize.Width / 2));

                    g.DrawString(text, fnt, Brushes.Black, stringPosX, stringPosY);

                    g.Dispose();
                    break;
                case "Barcode":
                    if (SettingsBarcodePro == null) { throw new Exception("SettingsBarcodeProがnull"); }

                    sizeX = (decimal)SettingsBarcodePro.BarcodeProPageSettings.SizeX / 25.4M * resolution * magnitude;
                    sizeY = (decimal)SettingsBarcodePro.BarcodeProPageSettings.SizeY / 25.4M * resolution * magnitude;
                    fontSize = (decimal)SettingsBarcodePro.BarcodeProLabelSettings.Font.SizeInPoints / 72.0M * resolution * magnitude;
                    fnt = new(SettingsBarcodePro.BarcodeProLabelSettings.Font.Name, (float)fontSize);

                    labelImage = new((int)sizeX, (int)sizeY);
                    g = Graphics.FromImage(labelImage);

                    int barWeight;
                    barWeight = resolution == _displayResolution ? 1 : (int)(1 * resolution / _displayResolution / _displayMagnitude);

                    var img = Code128Rendering.MakeBarcodeImage(text, barWeight, true);
                    var imageWidth = (decimal)(img.Width * SettingsBarcodePro.BarcodeProLabelSettings.BarcodeMagnitude);

                    if (imageWidth > labelImage.Width) { MessageBox.Show($"バーコードの幅がラベル幅を超えています{imageWidth}>{labelImage.Width}"); }

                    stringSize = g.MeasureString(text, fnt);

                    stringPosX = SettingsBarcodePro.BarcodeProLabelSettings.AlignStringCenter
                        ? (float)((labelImage.Width / 2) - (stringSize.Width / 2))
                        : (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4M * resolution * magnitude);

                    stringPosY = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosY / 25.4M * resolution * magnitude);


                    float barCodePosX;
                    barCodePosX = SettingsBarcodePro.BarcodeProLabelSettings.AlignBarcodeCenter
                        ? (float)((labelImage.Width / 2) - (imageWidth / 2))
                        : (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.StringPosX / 25.4M * resolution * magnitude);

                    float barCodePosY = (int)((decimal)SettingsBarcodePro.BarcodeProLabelSettings.BarcodePosY / 25.4M * resolution * magnitude);

                    g.DrawString(text, fnt, Brushes.Black, stringPosX, stringPosY);

                    var barcodeHeight = (int)(SettingsBarcodePro.BarcodeProLabelSettings.BarcodeHeight / 25.4F * resolution * magnitude);
                    g.DrawImage(img, barCodePosX, barCodePosY, (float)imageWidth, barcodeHeight);

                    g.Dispose();
                    img.Dispose();
                    break;
                default:
                    break;
            }
            return labelImage;
        }
        // チェックボックスイベント
        private void CheckBoxChecked(object sender, EventArgs e) {
            var checkBox = (CheckBox)sender;

            switch (checkBox.Name) {
                case "OrderNumberCheckBox":
                    OrderNumberTextBox.Enabled = checkBox.Checked;
                    break;
                case "ManufacturingNumberCheckBox":
                    ManufacturingNumberMaskedTextBox.Enabled = checkBox.Checked;
                    break;
                case "QuantityCheckBox":
                    QuantityTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) {
                        ExtraCheckBox1.Checked = false;
                    }

                    break;
                case "ExtraCheckBox1":
                    ExtraTextBox2.Enabled = checkBox.Checked;
                    break;
                case "RevisionCheckBox":
                    RevisionTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) {
                        MessageBox.Show("変更する場合は理由を記載して下さい。");
                    }

                    break;
                case "ExtraCheckBox2":
                    ExtraTextBox3.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox3":
                    ExtraTextBox3.Enabled = checkBox.Checked;
                    break;
                case "FirstSerialNumberCheckBox":
                    FirstSerialNumberTextBox.Enabled = checkBox.Checked;
                    break;
                case "RegistrationDateCheckBox":
                    RegistrationDateMaskedTextBox.Enabled = checkBox.Checked;
                    break;
                case "PersonCheckBox":
                    PersonComboBox.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox4":
                    ExtraTextBox4.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox5":
                    ExtraTextBox5.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox6":
                    ExtraTextBox6.Enabled = checkBox.Checked;
                    break;
                case "CommentCheckBox":
                    CommentTextBox.Enabled = checkBox.Checked;
                    CommentComboBox.Enabled = checkBox.Checked;
                    TemplateButton.Enabled = checkBox.Checked;
                    break;
            }
        }
        // 入力数値のみ
        private void NumericOnly(object sender, KeyPressEventArgs e) {
            // 0～9と、バックスペース以外の時は、イベントをキャンセルする
            if (e.KeyChar is (< '0' or > '9') and not '\b') {
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
            _strSerialType = "Label";
            Registeration();
        }
        private void BarcodePrintButton_Click(object sender, EventArgs e) {
            _strSerialType = "Barcode";
            Registeration();
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            var message = String.Join(Environment.NewLine,
                $"StrProductName\t\t[{ProductInfo.ProductName}]",
                $"StrProductModel\t\t[{ProductInfo.ProductModel}]",
                $"StrProductType\t\t[{ProductInfo.ProductType}]",
                $"StrOrderNumber\t\t[{_orderNumber}]",
                $"StrProductNumber\t\t[{_productNumber}]",
                $"StrRevision\t\t[{_revision}]",
                $"IntRegType\t\t[{ProductInfo.RegType}]",
                $"StrRegDate\t\t[{_regDate}]",
                $"StrPerson\t\t\t[{_person}]",
                $"IntQuantity\t\t[{_quantity}]",
                $"IntSerialFirstNumber\t[{_serialFirstNumber}]",
                $"IntSerialLastNumber\t[{_serialLastNumber}]",
                $"StrInitial\t\t\t[{ProductInfo.Initial}]",
                $"IntPrintType\t\t[{ProductInfo.PrintType}]",
                $"IntSerialDigit\t\t[{ProductInfo.SerialDigit}]"
            );
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
