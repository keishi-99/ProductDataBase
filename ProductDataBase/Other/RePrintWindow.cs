using GenCode128;
using ProductDatabase.Other;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class RePrintWindow : Form {

        public CSettingsLabelPro SettingsLabelPro { get; set; } = new CSettingsLabelPro();
        public string labelSettingFilePath = string.Empty;

        public CSettingsBarcodePro SettingsBarcodePro { get; set; } = new CSettingsBarcodePro();
        public string barcodeSettingFilePath = string.Empty;

        public ProductInfomation ProductInfo { get; }

        private string _orderNumber = string.Empty;
        private string _productNumber = string.Empty;
        private string _regDate = string.Empty;
        private string _person = string.Empty;
        private string _revision = string.Empty;
        private string _comment = string.Empty;

        private int _quantity;
        private int _serialFirstNumber;
        private int _serialLastNumber;

        private int _labelProPageNum = 1;
        private int _labelProNSerial;
        private int _labelProNumLabelsToPrint;
        private int _remainingCount;

        private readonly decimal _displayResolution = 96.0m;
        private readonly int _displayMagnitude = 3;
        private int _intPageCnt = 1;

        private string _serialType = string.Empty;
        private string _strSerialFirstNumber = string.Empty;
        private string _strSerialLastNumber = string.Empty;
        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "ExtraCheckBox1",
                    "RevisionCheckBox", "ExtraCheckBox2", "ExtraCheckBox3", "FirstSerialNumberCheckBox", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" ];

        // プロパティ設定
        private bool IsLabelPrint => ProductInfo.PrintType is 1 or 3 or 4 or 5 or 6 or 7 or 9;
        private bool IsBarcodePrint => ProductInfo.PrintType is 2 or 3;
        private bool IsUnderlinePrint => ProductInfo.PrintType is 4;

        public RePrintWindow(ProductInfomation productInfo) {
            InitializeComponent();
            ProductInfo = productInfo;
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

                // 印刷UI設定
                ConfigurePrintSettings();

                // TextBoxへ今日の年月日を入力
                var dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = dtNow.ToShortDateString();

                // DB1へ接続し担当者取得
                using (SQLiteConnection con = new(GetConnectionInfomation())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    cmd.CommandText = """SELECT * FROM Person ORDER BY _rowid_ ASC""";
                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        PersonComboBox.Items.Add($"{dr["PersonName"]}");
                    }
                }

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using (SQLiteConnection con = new(GetConnectionRegistration())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - [ProductName]テーブルの[SubstrateModel]列の[Revision]を取得
                    cmd.CommandText = $"""SELECT Revision FROM "{ProductInfo.ProductName}_Product" ORDER BY "ID" DESC""";
                    var result = cmd.ExecuteScalar();
                    RevisionTextBox.Text = result?.ToString() ?? "";
                }

                FirstSerialNumberCheckBox.Checked = true;
                LoadSettings(labelSettingFilePath, barcodeSettingFilePath);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadSettings(string labelSettingFilePath, string barcodeSettingFilePath) {
            try {
                if (labelSettingFilePath != string.Empty) {
                    StreamReader? srLabel = new(labelSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer serializerLabel = new(typeof(CSettingsLabelPro));
                    if (serializerLabel.Deserialize(srLabel) is CSettingsLabelPro result) { SettingsLabelPro = result; }
                    srLabel?.Close();
                }
                if (barcodeSettingFilePath != string.Empty) {
                    StreamReader? srBarcode = new(barcodeSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer serializerBarcode = new(typeof(CSettingsBarcodePro));
                    if (serializerBarcode.Deserialize(srBarcode) is CSettingsBarcodePro result) { SettingsBarcodePro = result; }
                    srBarcode?.Close();
                }
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // 印刷UI設定
        private void ConfigurePrintSettings() {
            ConfigureUI();
        }
        private void ConfigureUI() {
            if (IsLabelPrint) { ConfigureSerialLabelSettings(); }
            if (IsBarcodePrint) { ConfigureBarcodeSettings(); }
            SetMenuOptions();
        }
        private void ConfigureSerialLabelSettings() {
            SettingsLabelPro = new CSettingsLabelPro();
            labelSettingFilePath = Path.Combine(Environment.CurrentDirectory, "config", ProductInfo.ProductName, $"SerialConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml");
        }
        private void ConfigureBarcodeSettings() {
            SettingsBarcodePro = new CSettingsBarcodePro();
            barcodeSettingFilePath = Path.Combine(Environment.CurrentDirectory, "config", ProductInfo.ProductName, $"BarcodeConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml");
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
                FormCheck();
                DataCheck();

                DialogResult result;
                result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) { return; }

                result = MessageBox.Show("同一のシリアルラベルが複数存在しないようにして下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) { return; }

                ProductInfo.Person = PersonComboBox.Text;
                if (!PrintBarcode(1)) { throw new Exception("キャンセルしました。"); }

                Registeration();

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Registeration() {
            try {
                // 再印刷登録テーブルへ追加
                using SQLiteConnection con = new(GetConnectionRegistration());
                con.Open();
                using var cmd = con.CreateCommand();
                cmd.CommandText =
                    $"""
                    INSERT INTO Reprint
                        (PrintType, OrderNumber, ProductNumber, ProductType, ProductModel, Quantity, Person, RegDate, Revision, SerialFirst, SerialLast, Comment)
                    VALUES
                        (@PrintType, @OrderNumber, @ProductNumber, @ProductType, @ProductModel, @Quantity, @Person, @RegDate, @Revision, @SerialFirst, @SerialLast, @Comment)
                    """;

                // チェックボックスにチェックがない場合はNullを
                cmd.Parameters.Add("@PrintType", DbType.String).Value = string.IsNullOrWhiteSpace(_serialType) ? DBNull.Value : _serialType;
                cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(_orderNumber) ? DBNull.Value : _orderNumber;
                cmd.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(_productNumber) ? DBNull.Value : _productNumber;
                cmd.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                cmd.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                cmd.Parameters.Add("@Quantity", DbType.String).Value = _quantity;
                cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(_person) ? DBNull.Value : _person;
                cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(_regDate) ? DBNull.Value : _regDate;
                cmd.Parameters.Add("@Revision", DbType.String).Value = string.IsNullOrWhiteSpace(_revision) ? DBNull.Value : _revision;
                cmd.Parameters.Add("@SerialFirst", DbType.String).Value = string.IsNullOrWhiteSpace(_strSerialFirstNumber) ? DBNull.Value : _strSerialFirstNumber;
                cmd.Parameters.Add("@SerialLast", DbType.String).Value = string.IsNullOrWhiteSpace(_strSerialLastNumber) ? DBNull.Value : _strSerialLastNumber;
                cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(_comment) ? DBNull.Value : _comment;

                cmd.ExecuteNonQuery();

                // バックアップ作成
                BackupManager.CreateBackup();
                // ログ出力
                Logger.AppendLog($";[再印刷];注文番号[{_orderNumber}];製造番号[{_productNumber}];製品名[{ProductInfo.ProductType}];型式[{ProductInfo.ProductModel}];数量[{_quantity}];シリアル先頭[{_strSerialFirstNumber}];シリアル末尾[{_strSerialLastNumber}];Revision[{_revision}];登録日[{_regDate}];担当者[{_person}];");

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool FormCheck() {
            // 入力フォームのチェック
            var anyTextBoxEnabled = false;
            var allTextBoxesFilled = true;

            foreach (Control control in Controls) {
                if (control is TextBoxBase textBox && textBox.Enabled) {
                    anyTextBoxEnabled = true;
                    if (string.IsNullOrWhiteSpace(textBox.Text) && textBox.Name != "QrCodeTextBox") {
                        allTextBoxesFilled = false;
                        break;
                    }
                }
            }
            return !anyTextBoxEnabled
                ? throw new Exception("何も入力されていません")
                : !allTextBoxesFilled
                ? throw new Exception("空欄があります。")
                : ManufacturingNumberCheckBox.Checked && ManufacturingNumberMaskedTextBox.Text.Length != 15
                ? throw new Exception("製番を10桁+4桁で入力して下さい。")
                : QuantityCheckBox.Checked && int.Parse(QuantityTextBox.Text) <= 0 ? throw new Exception("1台以上入力して下さい。") : true;
        }
        private bool DataCheck() {
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
            return true;
        }
        // 印刷処理
        private bool PrintBarcode(int printFlg) {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();

                // PrintPageイベントハンドラの追加
                pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

                _intPageCnt = 1;

                switch (printFlg) {
                    case 1:
                        _labelProNumLabelsToPrint = _quantity;
                        _labelProPageNum = 1;
                        RePrintPrintDialog.Document = pd;
                        var r = RePrintPrintDialog.ShowDialog();

                        if (r == DialogResult.OK) {
                            // ローディング画面の表示
                            using var loadingForm = new LoadingForm();
                            // 別スレッドで印刷処理を実行
                            Task.Run(() => {
                                try {
                                    RePrintPrintDialog.Document.Print();
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
                        if (!FormCheck()) { return false; };
                        if (!DataCheck()) { return false; };
                        _labelProNumLabelsToPrint = _quantity;
                        _labelProPageNum = 1;
                        // 最大で表示
                        RePrintPrintPreviewDialog.Shown += (sender, e) => {
                            if (sender is Form form) {
                                form.WindowState = FormWindowState.Maximized;
                            }
                        };
                        RePrintPrintPreviewDialog.Document = pd;
                        RePrintPrintPreviewDialog.ShowDialog();
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
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            Point headerPos = new(0, 0);
            var headerString = string.Empty;
            Font headerFooterFont = new("ＭＳ Ｐ明朝", 5.25F);
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
                        break;
                    default:
                        break;
                }

                var startLineBarcode = (int)PrintPostionNumericUpDown.Value - 1;

                const decimal MM_PER_HUNDREDTH_INCH = 0.254M;

                if (!RePrintPrintDocument.PrintController.IsPreview) {
                    offsetX -= (decimal)e.PageSettings.HardMarginX * MM_PER_HUNDREDTH_INCH;
                    offsetY -= (decimal)e.PageSettings.HardMarginY * MM_PER_HUNDREDTH_INCH;
                    offset = _labelProPageNum == 1
                        ? new Point((int)((decimal)e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (startLineBarcode * (intervalY + sizeY))))
                        : new Point((int)((decimal)e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (0 * (intervalY + sizeY))));
                }
                else {
                    offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (0 * (intervalY + sizeY))));
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;

                headerPos.Offset(offset);
                e.Graphics.DrawString(headerString, headerFooterFont, Brushes.Black, headerPos);

                var barcodePageNum = 0;
                offset = barcodePageNum == 0
                    ? new Point((int)((decimal)e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (startLineBarcode * (intervalY + sizeY))))
                    : new Point((int)((decimal)e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)(((decimal)e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (0 * (intervalY + sizeY))));

                if (_labelProPageNum == 1) {
                    _remainingCount = serialCodePrintCopies;
                    _labelProNSerial = _serialFirstNumber;
                }
                if (_labelProPageNum >= 2) { startLineBarcode = 0; }

                var y = 0;
                for (y = startLineBarcode; y < maxY; y++) {
                    var x = 0;
                    for (x = 0; x < maxX; x++) {
                        var posX = (float)(offsetX + (x * (intervalX + sizeX)));
                        var posY = (float)(offsetY + (y * (intervalY + sizeY)));

                        // タイプ4で残り1の場合、最後のラベルに下線をつける
                        var fontUnderline = IsUnderlinePrint && _remainingCount == 1;

                        // シリアル生成、PrintTypeが9かつ最終行の場合は型式下4桁、それ以外はシリアルを生成
                        string generatedCode;
                        if (ProductInfo.PrintType != 9 || _remainingCount != 1) {
                            generatedCode = GenerateCode(_labelProNSerial); // シリアルコードを生成
                        }
                        else {
                            generatedCode = ProductInfo.ProductModel[^4..]; // 型式の下4桁を使用
                        }

                        using var labelImage = MakeLabelImage(generatedCode, (int)e.Graphics.DpiX, 1, fontUnderline);
                        e.Graphics.DrawImage(labelImage, posX, posY, (float)sizeX, (float)sizeY);

                        _remainingCount--;
                        if (_remainingCount <= 0) {
                            _labelProNSerial++;
                            _labelProNumLabelsToPrint--;
                            //印刷するラベルがなくなった場合の処理
                            if (_labelProNumLabelsToPrint <= 0) {
                                e.HasMorePages = false;
                                _labelProPageNum = 1;
                                _labelProNumLabelsToPrint = 0;
                                return;
                            }
                            _remainingCount = serialCodePrintCopies;
                        }
                    }
                }

                if (_labelProNumLabelsToPrint > 0) {
                    _labelProPageNum++;
                    _intPageCnt++;
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
                 .Replace("%M", _productNumber)
                 .Replace("%O", _orderNumber)
                 .Replace("%N", _quantity.ToString())
                 .Replace("%U", ProductInfo.Person);
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

            var outputCode = _serialType switch {
                "Label" => SettingsLabelPro.LabelProLabelSettings.Format,
                "Barcode" => SettingsBarcodePro.BarcodeProLabelSettings.Format,
                _ => string.Empty
            };

            outputCode = outputCode.Replace("%Y", DateTime.Parse(_regDate).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(_regDate).ToString("MM"))
                                    .Replace("%T", ProductInfo.Initial)
                                    .Replace("%R", _revision)
                                    .Replace("%M", monthCode[^1..])
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

                    var barWeight = resolution == _displayResolution ? 1 : (int)(1 * resolution / _displayResolution / _displayMagnitude);

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
                    break;
                case "ExtraCheckBox1":
                    ExtraTextBox2.Enabled = checkBox.Checked;
                    break;
                case "RevisionCheckBox":
                    RevisionTextBox.Enabled = checkBox.Checked;
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
        // QR入力処理
        private void QrInput() {
            try {
                if (string.IsNullOrWhiteSpace(QrCodeTextBox.Text)) { return; }
                string[] separator = ["//"];
                var arr = QrCodeTextBox.Text.Split(separator, StringSplitOptions.None);
                if (arr.Length != 4) {
                    MessageBox.Show("QRコードが正しくありません。");
                    return;
                }
                if (arr != null) {
                    ProductInfo.Proness1 = arr[0];
                    ProductInfo.Proness2 = arr[1];
                    ProductInfo.Proness4 = Convert.ToInt32(arr[2] ?? throw new Exception());
                    ProductInfo.Proness5 = arr[3];
                }
                OrderNumberTextBox.Text = ProductInfo.Proness5;
                ManufacturingNumberMaskedTextBox.Text = ProductInfo.Proness1;
                QuantityTextBox.Text = ProductInfo.Proness4.ToString();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RePrintWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void QrCodeButton_Click(object sender, EventArgs e) { QrInput(); }
        private void LabelPrintButton_Click(object sender, EventArgs e) {
            _serialType = "Label";
            RegisterCheck();
        }
        private void BarcodePrintButton_Click(object sender, EventArgs e) {
            _serialType = "Barcode";
            RegisterCheck();
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            var entries = new[]
                {
                    ("StrProductName", $"{ProductInfo.ProductName}"),
                    ("StrProductModel", $"{ProductInfo.ProductModel}"),
                    ("StrProductType", $"{ProductInfo.ProductType}"),
                    ("StrOrderNumber", $"{_orderNumber}"),
                    ("StrProductNumber", $"{_productNumber}"),
                    ("StrRevision", $"{_revision}"),
                    ("IntRegType", $"{ProductInfo.RegType}"),
                    ("StrRegDate", $"{_regDate}"),
                    ("StrPerson", $"{_person}"),
                    ("IntQuantity", $"{_quantity}"),
                    ("IntSerialFirstNumber", $"{_serialFirstNumber}"),
                    ("IntSerialLastNumber", $"{_serialLastNumber}"),
                    ("StrInitial", $"{ProductInfo.Initial}"),
                    ("IntPrintType", $"{ProductInfo.PrintType}"),
                    ("IntSerialDigit", $"{ProductInfo.SerialDigit}")
                };

            const int ColumnWidth = 15;
            var message = string.Join(Environment.NewLine,
                entries.Select(entry => $"{entry.Item1,-ColumnWidth}[{entry.Item2}]"));

            MessageBox.Show(message, "取得情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void シリアルラベル印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Label";
            PrintBarcode(2);
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
        private void ProductRegistration2PrintPreviewDialog_Load(object sender, EventArgs e) {
            var tool = (ToolStrip)RePrintPrintPreviewDialog.Controls[1];
            tool.Items[0].Visible = false;
        }
    }
}
