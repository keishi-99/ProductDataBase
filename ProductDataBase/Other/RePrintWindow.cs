using GenCode128;
using ProductDatabase.Product;
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

        public int LabelProPageNum { get; set; }
        public int LabelProNSerial { get; set; }
        public int LabelProNumLabelsToPrint { get; set; }

        public decimal DisplayResolution { get; } = 96.0m;
        public int DisplayMagnitude { get; } = 3;
        public int IntPageCnt { get; set; } = 1;

        private string _strSerialType = string.Empty;
        private string _strSerialFirstNumber = string.Empty;
        private string _strSerialLastNumber = string.Empty;
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
                var dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = dtNow.ToShortDateString();

                // DB1へ接続し担当者取得
                using (SQLiteConnection con = new(MainWindow.GetConnectionString1())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        PersonComboBox.Items.Add($"{dr["col_Person_Name"]}");
                    }
                }

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using (SQLiteConnection con = new(MainWindow.GetConnectionString2())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - [Product_Name]_stockテーブルの[col_Substrate_Model]列の[col_Revision]を取得
                    cmd.CommandText = $"SELECT col_Revision FROM Product_Reg_{StrProductName} ORDER BY _rowid_ DESC";
                    var result = cmd.ExecuteScalar();
                    RevisionTextBox.Text = result?.ToString() ?? "";
                }

                // 変数[check_bin]の値に応じてCheckboxにチェックを入れる
                foreach (var checkBoxName in _checkBoxNames) {
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
                    StreamReader? srLabel = new(strLabelSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer serializerLabel = new(typeof(CSettingsLabelPro));
                    if (serializerLabel.Deserialize(srLabel) is CSettingsLabelPro result) { SettingsLabelPro = result; }
                    srLabel?.Close();
                }
                if (strBarcodeSettingFilePath != string.Empty) {
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
                        if (string.IsNullOrWhiteSpace(textBox.Text)) {
                            allTextBoxesFilled = false;
                            break;
                        }
                    }
                }
                if (!anyTextBoxEnabled) { throw new Exception("何も入力されていません"); }
                if (!allTextBoxesFilled) { throw new Exception("空欄があります。"); }

                if (ManufacturingNumberCheckBox.Checked && ManufacturingNumberMaskedTextBox.Text.Length != 15) { throw new Exception("製番を10桁+4桁で入力して下さい。"); }

                if (QuantityCheckBox.Checked && int.Parse(QuantityTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

                result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) {
                    return;
                }

                var quantity = Convert.ToInt32(QuantityTextBox.Text);
                if (quantity == 0) { throw new Exception("1以上入力してください。"); }
                var firstSerial = Convert.ToInt32(FirstSerialNumberTextBox.Text);
                if (firstSerial == 0) { throw new Exception("シリアル開始番号を入力してください。"); }

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
                    if (quantity + firstSerial >= threshold) {
                        MessageBox.Show($"シリアルが{threshold}を超えるので{resetValue.ToString().PadLeft(IntSerialDigit, '0')}から開始します。");
                        FirstSerialNumberTextBox.Text = resetValue.ToString();
                    }
                }

                result = MessageBox.Show("同一のシリアルラベルが複数存在しないようにして下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) { return; }

                StrOrderNumber = OrderNumberTextBox.Text;
                StrProductNumber = ManufacturingNumberMaskedTextBox.Text;
                IntQuantity = Convert.ToInt32(QuantityTextBox.Text ?? throw new Exception());
                StrPerson = PersonComboBox.Text;
                StrRegDate = RegistrationDateMaskedTextBox.Text;
                StrRevision = RevisionTextBox.Text;
                StrComment = CommentTextBox.Text;

                IntSerialFirstNumber = Convert.ToInt32(FirstSerialNumberTextBox.Text);
                IntSerialLastNumber = IntSerialFirstNumber + IntQuantity - 1;

                _strSerialFirstNumber = GenerateCode(IntSerialFirstNumber);
                _strSerialLastNumber = GenerateCode(IntSerialLastNumber);

                if (!PrintBarcode(1)) { throw new Exception("キャンセルしました。"); }

                // 再印刷登録テーブルへ追加
                using SQLiteConnection con = new(MainWindow.GetConnectionString2());
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
                cmd.Parameters.Add("@col_Print_Type", DbType.String).Value = IntPrintType;
                cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = StrOrderNumber;
                cmd.Parameters.Add("@col_Product_Num", DbType.String).Value = StrProductNumber;
                cmd.Parameters.Add("@col_Product_Type", DbType.String).Value = StrProductType;
                cmd.Parameters.Add("@col_Product_Model", DbType.String).Value = StrProductModel;
                cmd.Parameters.Add("@col_Quantity", DbType.String).Value = IntQuantity;
                cmd.Parameters.Add("@col_Person", DbType.String).Value = StrPerson;
                cmd.Parameters.Add("@col_RegDate", DbType.String).Value = StrRegDate;
                cmd.Parameters.Add("@col_Revision", DbType.String).Value = StrRevision;
                cmd.Parameters.Add("@col_Serial_First", DbType.String).Value = _strSerialFirstNumber;
                cmd.Parameters.Add("@col_Serial_Last", DbType.String).Value = _strSerialLastNumber;
                cmd.Parameters.Add("@col_Comment", DbType.String).Value = StrComment;

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

            LabelProNumLabelsToPrint = IntQuantity;
            LabelProPageNum = 0;
            _fontUnderbar = false;

            switch (printFlg) {
                case 1:
                    RePrintPrintDialog.Document = pd;
                    var r = RePrintPrintDialog.ShowDialog();

                    if (r == DialogResult.OK) {
                        RePrintPrintDialog.Document.Print();

                        if (IntPageCnt >= 2) {
                            MessageBox.Show($"{IntPageCnt}枚印刷されます。2枚目以降は1行目から印刷されます。");
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
                var headerString = string.Empty;
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
                        maxX = SettingsLabelPro._labelProPageSettings.NumLabelsX;
                        maxY = SettingsLabelPro._labelProPageSettings.NumLabelsY;
                        sizeX = (float)SettingsLabelPro._labelProPageSettings.SizeX;
                        sizeY = (float)SettingsLabelPro._labelProPageSettings.SizeY;
                        offsetX = SettingsLabelPro._labelProPageSettings.OffsetX;
                        offsetY = SettingsLabelPro._labelProPageSettings.OffsetY;
                        intervalX = SettingsLabelPro._labelProPageSettings.IntervalX;
                        intervalY = SettingsLabelPro._labelProPageSettings.IntervalY;
                        headerPos = SettingsLabelPro._labelProPageSettings.HeaderPos;
                        headerString = ConvertHeaderFooterString(SettingsLabelPro._labelProPageSettings.HeaderString);
                        headerFooterFont = SettingsLabelPro._labelProPageSettings.HeaderFooterFont;
                        intNumLabels = SettingsLabelPro._labelProLabelSettings.NumLabels;
                        intCountNumLabels = SettingsLabelPro._labelProLabelSettings.NumLabels;
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
                    offset = LabelProPageNum == 0
                        ? new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (startLineBarcode * (intervalY + sizeY))))
                        : new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (intervalY + sizeY))));
                }
                else {
                    offset = new Point(0, 0);
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;

                headerPos.Offset(offset);
                e.Graphics.DrawString(headerString, headerFooterFont, Brushes.Black, headerPos);

                var barcodePageNum = 0;
                offset = barcodePageNum == 0
                    ? new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (startLineBarcode * (intervalY + sizeY))))
                    : new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (intervalY + sizeY))));

                if (LabelProPageNum == 0) { LabelProNSerial = IntSerialFirstNumber; }
                if (LabelProPageNum >= 1) { startLineBarcode = 0; }

                var y = 0;
                for (y = startLineBarcode; y < maxY; y++) {
                    var x = 0;
                    for (x = 0; x < maxX; x++) {
                        var s = GenerateCode(LabelProNSerial);
                        var posX = (float)(offsetX + (x * (intervalX + sizeX)));
                        var posY = (float)(offsetY + (y * (intervalY + sizeY)));
                        e.Graphics.DrawImage(MakeLabelImage(s, (int)e.Graphics.DpiX, 1), posX, posY, sizeX, sizeY);

                        LabelProNSerial++;
                        LabelProNumLabelsToPrint--;

                        if (LabelProNumLabelsToPrint <= 0) {
                            intCountNumLabels--;
                            if (intCountNumLabels <= 0) {
                                e.HasMorePages = false;
                                LabelProPageNum = 0;
                                var txtNumPublish = 0;
                                LabelProNumLabelsToPrint = txtNumPublish;
                                return;
                            }
                            else {
                                LabelProNumLabelsToPrint += x + 1;
                                break;
                            }
                        }

                        if (x >= maxX - 1) {
                            intCountNumLabels--;
                            if (intCountNumLabels <= 0) {
                                intCountNumLabels = intNumLabels;
                            }
                            else if (intCountNumLabels > 0) {
                                LabelProNumLabelsToPrint += x + 1;
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
            var monthCode = DateTime.Parse(StrRegDate).ToString("MM");

            monthCode = monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => monthCode
            };

            var outputCode = _strSerialType switch {
                "Label" => SettingsLabelPro._labelProLabelSettings.Format,
                "Barcode" => SettingsBarcodePro.BarcodeProLabelSettings.Format,
                _ => string.Empty
            };

            outputCode = outputCode.Replace("%Y", DateTime.Parse(StrRegDate).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(StrRegDate).ToString("MM"))
                                    .Replace("%T", StrInitial)
                                    .Replace("%R", StrRevision)
                                    .Replace("%M", monthCode[^1..])
                                    .Replace("%S", Convert.ToInt32(serialCode).ToString($"D{IntSerialDigit}"));

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
                    sizeX = (decimal)SettingsLabelPro._labelProPageSettings.SizeX / 25.4M * resolution * magnitude;
                    sizeY = (decimal)SettingsLabelPro._labelProPageSettings.SizeY / 25.4M * resolution * magnitude;
                    fontSize = (decimal)SettingsLabelPro._labelProLabelSettings.Font.SizeInPoints / 72.0M * resolution * magnitude;
                    stringPosY = (int)((decimal)SettingsLabelPro._labelProLabelSettings.StringPosY / 25.4M * resolution * magnitude);
                    var style = _fontUnderbar ? FontStyle.Underline : FontStyle.Regular;
                    fnt = new Font(SettingsLabelPro._labelProLabelSettings.Font.Name, (float)fontSize, style);

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
                    barWeight = resolution == DisplayResolution ? 1 : (int)(1 * resolution / DisplayResolution / DisplayMagnitude);

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
            var message = string.Join(Environment.NewLine,
                $"StrProductName\t\t[{StrProductName}]",
                $"StrProductModel\t\t[{StrProductModel}]",
                $"StrProductType\t\t[{StrProductType}]",
                $"StrOrderNumber\t\t[{StrOrderNumber}]",
                $"StrProductNumber\t\t[{StrProductNumber}]",
                $"StrRevision\t\t[{StrRevision}]",
                $"IntRegType\t\t[{IntRegType}]",
                $"StrRegDate\t\t[{StrRegDate}]",
                $"StrPerson\t\t\t[{StrPerson}]",
                $"IntQuantity\t\t[{IntQuantity}]",
                $"IntSerialFirstNumber\t[{IntSerialFirstNumber}]",
                $"IntSerialLastNumber\t[{IntSerialLastNumber}]",
                $"StrInitial\t\t\t[{StrInitial}]",
                $"IntPrintType\t\t[{IntPrintType}]",
                $"IntSerialDigit\t\t[{IntSerialDigit}]"
            );
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
