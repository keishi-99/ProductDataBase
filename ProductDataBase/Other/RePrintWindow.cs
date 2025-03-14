using GenCode128;
using ProductDatabase.Other;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class RePrintWindow : Form {

        public ProductPrintSettings ProductPrintSettings { get; set; } = new ProductPrintSettings();
        public string labelSettingFilePath = string.Empty;

        public ProductInformation ProductInfo { get; }

        private int _remainingCount;
        private int _labelProNSerial;
        private int _labelProNumLabelsToPrint;

        private int _pageCount = 1;
        private System.Drawing.Printing.PrintAction _printAction;

        private readonly double _displayResolution = 96.0;
        private readonly int _displayMagnitude = 3;

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

        public RePrintWindow(ProductInformation productInfo) {
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

                // TextBoxへ今日の年月日を入力
                var dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = dtNow.ToShortDateString();

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. ProductInfo.PersonList]);

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using (SQLiteConnection con = new(GetConnectionRegistration())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - [ProductName]テーブルの[SubstrateModel]列の[Revision]を取得
                    cmd.CommandText = $"""SELECT Revision FROM "{ProductInfo.CategoryName}_Product" WHERE ProductName = @ProductName AND RevisionGroup = @RevisionGroup ORDER BY "ID" DESC""";
                    cmd.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;
                    cmd.Parameters.Add("@RevisionGroup", DbType.String).Value = ProductInfo.RevisionGroup;
                    var result = cmd.ExecuteScalar();
                    RevisionTextBox.Text = result?.ToString() ?? "";
                }

                FirstSerialNumberCheckBox.Checked = true;
                LabelPrintButton.Enabled = IsLabelPrint;
                BarcodePrintButton.Enabled = IsBarcodePrint;

                // 印刷UI設定
                ConfigurePrintSettings();
                LoadSettings(labelSettingFilePath);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }
        private void LoadSettings(string labelSettingFilePath) {
            try {
                if (labelSettingFilePath != string.Empty) {
                    StreamReader? srLabel = new(labelSettingFilePath, new System.Text.UTF8Encoding(false));
                    System.Xml.Serialization.XmlSerializer serializerLabel = new(typeof(ProductPrintSettings));
                    if (serializerLabel.Deserialize(srLabel) is ProductPrintSettings result) { ProductPrintSettings = result; }
                    srLabel?.Close();
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
            ConfigureSerialLabelSettings();
            SetMenuOptions();
        }
        private void ConfigureSerialLabelSettings() {
            ProductPrintSettings = new ProductPrintSettings();
            labelSettingFilePath = Path.Combine(Environment.CurrentDirectory, "config", "Product", ProductInfo.CategoryName, ProductInfo.ProductName, $"SerialConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.xml");
            if (!File.Exists(labelSettingFilePath)) { throw new DirectoryNotFoundException($"ラベル印刷用設定ファイルがありません。"); }
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
                if (!Registeration()) { throw new Exception("登録できませんでした。"); }

                if (!PrintBarcode()) { throw new Exception("キャンセルしました。"); }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool Registeration() {
            try {
                // 再印刷登録テーブルへ追加
                using SQLiteConnection con = new(GetConnectionRegistration());
                con.Open();
                using var cmd = con.CreateCommand();
                cmd.CommandText =
                    $"""
                    INSERT INTO Reprint
                        (PrintType, OrderNumber, ProductName, ProductNumber, ProductType, ProductModel, Quantity, Person, RegDate, Revision, SerialFirst, SerialLast, Comment)
                    VALUES
                        (@PrintType, @OrderNumber, @ProductName, @ProductNumber, @ProductType, @ProductModel, @Quantity, @Person, @RegDate, @Revision, @SerialFirst, @SerialLast, @Comment)
                    """;

                // チェックボックスにチェックがない場合はNullを
                cmd.Parameters.Add("@PrintType", DbType.String).Value = string.IsNullOrWhiteSpace(_serialType) ? DBNull.Value : _serialType;
                cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                cmd.Parameters.Add("@ProductName", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductName) ? DBNull.Value : ProductInfo.ProductName;
                cmd.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                cmd.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                cmd.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                cmd.Parameters.Add("@Quantity", DbType.String).Value = ProductInfo.Quantity;
                cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                cmd.Parameters.Add("@Revision", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Revision) ? DBNull.Value : ProductInfo.Revision;
                cmd.Parameters.Add("@SerialFirst", DbType.String).Value = string.IsNullOrWhiteSpace(_strSerialFirstNumber) ? DBNull.Value : _strSerialFirstNumber;
                cmd.Parameters.Add("@SerialLast", DbType.String).Value = string.IsNullOrWhiteSpace(_strSerialLastNumber) ? DBNull.Value : _strSerialLastNumber;
                cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                cmd.ExecuteNonQuery();

                // バックアップ作成
                CommonUtils.BackupManager.CreateBackup();
                // ログ出力
                string[] logMessageArray = [
                    $"[再印刷]",
                    $"注文番号[{ProductInfo.OrderNumber}]",
                    $"製造番号[{ProductInfo.ProductNumber}]",
                    $"製品名[{ProductInfo.ProductName}]",
                    $"タイプ[{ProductInfo.ProductType}]",
                    $"型式[{ProductInfo.ProductModel}]",
                    $"数量[{ProductInfo.Quantity}]",
                    $"シリアル先頭[{_strSerialFirstNumber}]",
                    $"シリアル末尾[{_strSerialLastNumber}]",
                    $"Revision[{ProductInfo.Revision}]",
                    $"登録日[{ProductInfo.RegDate}]",
                    $"担当者[{ProductInfo.Person}]",
                    $"コメント[{ProductInfo.Comment}]"
                ];
                CommonUtils.Logger.AppendLog(logMessageArray);

                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
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

            ProductInfo.OrderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
            ProductInfo.ProductNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
            ProductInfo.Quantity = Convert.ToInt32(QuantityTextBox.Text ?? throw new Exception());
            ProductInfo.Person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
            ProductInfo.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateMaskedTextBox.Text : string.Empty;
            ProductInfo.Revision = RevisionCheckBox.Checked ? RevisionTextBox.Text : string.Empty;
            ProductInfo.Comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;

            ProductInfo.SerialFirstNumber = Convert.ToInt32(FirstSerialNumberTextBox.Text);
            ProductInfo.SerialLastNumber = ProductInfo.SerialFirstNumber + ProductInfo.Quantity - 1;

            _strSerialFirstNumber = GenerateCode(ProductInfo.SerialFirstNumber);
            _strSerialLastNumber = GenerateCode(ProductInfo.SerialLastNumber);
            return true;
        }
        // 印刷処理
        private bool PrintBarcode() {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();
                pd.BeginPrint += (sender, e) => _printAction = e.PrintAction;
                pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

                _labelProNumLabelsToPrint = ProductInfo.Quantity;
                _pageCount = 1;

                RePrintPrintDialog.Document = pd;

                if (RePrintPrintDialog.ShowDialog() == DialogResult.OK) {
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
                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private bool PreviewBarcode() {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();
                pd.BeginPrint += (sender, e) => _printAction = e.PrintAction;
                pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

                _labelProNumLabelsToPrint = ProductInfo.Quantity;
                _pageCount = 1;

                if (!FormCheck() || !DataCheck()) {
                    return false;
                }

                // 最大で表示
                RePrintPrintPreviewDialog.Shown += (sender, e) => {
                    if (sender is Form form) {
                        form.WindowState = FormWindowState.Maximized;
                    }
                };
                RePrintPrintPreviewDialog.PrintPreviewControl.Zoom = 3;
                RePrintPrintPreviewDialog.Document = pd;
                RePrintPrintPreviewDialog.ShowDialog();

                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            var headerString = string.Empty;
            Font headerFont = new("ＭＳ Ｐ明朝", 5.25F);
            var copiesPerLabel = 0;

            var labelCountX = 0;
            var labelCountY = 0;
            double labelWidth = 0;
            double labelHeight = 0;
            double marginX = 0;
            double marginY = 0;
            double intervalX = 0;
            double intervalY = 0;
            double headerPositionX = 0;
            double headerPositionY = 0;
            var startLine = 0;
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;

                switch (_serialType) {
                    case "Label":
                        if (ProductPrintSettings == null) { throw new Exception("ProductPrintSettingsがnullです。"); }
                        labelCountX = ProductPrintSettings.LabelPageSettings.LabelsPerColumn;
                        labelCountY = ProductPrintSettings.LabelPageSettings.LabelsPerRow;
                        labelWidth = ProductPrintSettings.LabelPageSettings.LabelWidth;
                        labelHeight = ProductPrintSettings.LabelPageSettings.LabelHeight;
                        marginX = ProductPrintSettings.LabelPageSettings.MarginX;
                        marginY = ProductPrintSettings.LabelPageSettings.MarginY;
                        intervalX = ProductPrintSettings.LabelPageSettings.IntervalX;
                        intervalY = ProductPrintSettings.LabelPageSettings.IntervalY;
                        headerPositionX = ProductPrintSettings.LabelPageSettings.HeaderPositionX;
                        headerPositionY = ProductPrintSettings.LabelPageSettings.HeaderPositionY;
                        headerString = ConvertHeaderString(ProductPrintSettings.LabelPageSettings.HeaderText);
                        headerFont = ProductPrintSettings.LabelPageSettings.HeaderFont;
                        copiesPerLabel = ProductPrintSettings.LabelLayoutSettings.CopiesPerLabel;
                        startLine = (int)PrintPostionNumericUpDown.Value - 1;
                        break;
                    case "Barcode":
                        if (ProductPrintSettings == null) { throw new Exception("ProductPrintSettingsがnullです。"); }
                        labelCountX = ProductPrintSettings.BarcodePageSettings.LabelsPerColumn;
                        labelCountY = ProductPrintSettings.BarcodePageSettings.LabelsPerRow;
                        labelWidth = ProductPrintSettings.BarcodePageSettings.LabelWidth;
                        labelHeight = ProductPrintSettings.BarcodePageSettings.LabelHeight;
                        marginX = ProductPrintSettings.BarcodePageSettings.MarginX;
                        marginY = ProductPrintSettings.BarcodePageSettings.MarginY;
                        intervalX = ProductPrintSettings.BarcodePageSettings.IntervalX;
                        intervalY = ProductPrintSettings.BarcodePageSettings.IntervalY;
                        headerPositionX = ProductPrintSettings.LabelPageSettings.HeaderPositionX;
                        headerPositionY = ProductPrintSettings.LabelPageSettings.HeaderPositionY;
                        headerString = ConvertHeaderString(ProductPrintSettings.BarcodePageSettings.HeaderText);
                        headerFont = ProductPrintSettings.BarcodePageSettings.HeaderFont;
                        copiesPerLabel = ProductPrintSettings.BarcodeLayoutSettings.CopiesPerLabel;
                        startLine = (int)PrintPostionNumericUpDown.Value - 1;
                        break;
                    default:
                        break;
                }

                if (labelCountX == 0 || labelCountY == 0 || copiesPerLabel == 0) { throw new Exception("印刷設定が異常です。"); }

                // ハードマージンをミリメートルに変換
                const double MM_PER_HUNDREDTH_INCH = 0.254;
                marginX -= e.PageSettings.HardMarginX * MM_PER_HUNDREDTH_INCH;
                marginY -= e.PageSettings.HardMarginY * MM_PER_HUNDREDTH_INCH;

                // 最初のページのみオフセットを調整
                var verticalOffset = _pageCount == 1 ? startLine * (intervalY + labelHeight) : 0;
                // ヘッダーの描画
                e.Graphics.DrawString(headerString, headerFont, Brushes.Black, (float)headerPositionX, (float)(verticalOffset + headerPositionY));

                if (_pageCount == 1) {
                    _remainingCount = copiesPerLabel;
                    _labelProNSerial = ProductInfo.SerialFirstNumber;
                }
                if (_pageCount >= 2) { startLine = 0; }

                var y = 0;
                for (y = startLine; y < labelCountY; y++) {
                    var x = 0;
                    for (x = 0; x < labelCountX; x++) {
                        var posX = (float)(marginX + (x * (intervalX + labelWidth)));
                        var posY = (float)(marginY + (y * (intervalY + labelHeight)));

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
                        e.Graphics.DrawImage(labelImage, posX, posY, (float)labelWidth, (float)labelHeight);

                        _remainingCount--;
                        if (_remainingCount <= 0) {
                            _labelProNSerial++;
                            _labelProNumLabelsToPrint--;
                            //印刷するラベルがなくなった場合の処理
                            if (_labelProNumLabelsToPrint <= 0) {
                                // 最終行の行番号を計算
                                var rowOffset = _serialType switch {
                                    "Label" => 2,
                                    "Barcode" => 1,
                                    _ => 0
                                };
                                // 最終行の行番号を表示
                                e.Graphics.DrawString((y + rowOffset).ToString(), ProductPrintSettings.LabelPageSettings.HeaderFont, Brushes.Black, 0, posY);

                                e.HasMorePages = false;
                                _pageCount = 1;
                                _labelProNumLabelsToPrint = 0;
                                return;
                            }
                            _remainingCount = copiesPerLabel;
                        }
                    }
                }

                if (_labelProNumLabelsToPrint > 0) {
                    _pageCount++;
                    e.HasMorePages = true;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        private string ConvertHeaderString(string s) {
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
                "Label" => ProductPrintSettings.LabelLayoutSettings.Format,
                "Barcode" => ProductPrintSettings.BarcodeLayoutSettings.Format,
                _ => string.Empty
            };

            outputCode = outputCode.Replace("%Y", DateTime.Parse(ProductInfo.RegDate).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(ProductInfo.RegDate).ToString("MM"))
                                    .Replace("%T", ProductInfo.Initial)
                                    .Replace("%R", ProductInfo.Revision)
                                    .Replace("%M", monthCode[^1..])
                                    .Replace("%S", Convert.ToInt32(serialCode).ToString($"D{ProductInfo.SerialDigit}"));

            return outputCode;
        }
        private Bitmap MakeLabelImage(string text, int resolution, int magnitude, bool fontUnderline) {
            Bitmap labelImage = new(1, 1);
            Graphics g;
            SizeF stringSize;
            double sizeX;
            double sizeY;
            double fontSize;
            float stringPosX;
            float stringPosY;
            Font fnt;
            // プレビューかどうかの判定
            var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;

            // サイズとフォント情報の設定
            void SetLabelProperties(double labelWidth, double labelHeight, double posY, double fontPointSize, string fontName, bool underlined) {
                sizeX = labelWidth / 25.4 * resolution * magnitude;
                sizeY = labelHeight / 25.4 * resolution * magnitude;
                stringPosY = (float)(posY / 25.4 * resolution * magnitude);
                fontSize = fontPointSize / 72.0 * resolution * magnitude;
                var style = underlined ? FontStyle.Underline : FontStyle.Regular;
                fnt = new Font(fontName, (float)fontSize, style);
            }

            switch (_serialType) {
                case "Label":
                    if (ProductPrintSettings == null) { throw new Exception("ProductPrintSettingsがnull"); }
                    SetLabelProperties(ProductPrintSettings.LabelPageSettings.LabelWidth,
                                        ProductPrintSettings.LabelPageSettings.LabelHeight,
                                        ProductPrintSettings.LabelLayoutSettings.TextPositionY,
                                        ProductPrintSettings.LabelLayoutSettings.TextFont.SizeInPoints,
                                        ProductPrintSettings.LabelLayoutSettings.TextFont.Name,
                                        fontUnderline);

                    labelImage = new((int)sizeX, (int)sizeY);
                    g = Graphics.FromImage(labelImage);
                    // アンチエイリアス処理を改善
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                    // StringFormat を使用して中心に配置
                    var sf = new StringFormat {
                        Alignment = ProductPrintSettings.LabelLayoutSettings.AlignTextXCenter ? StringAlignment.Center : StringAlignment.Near,
                        LineAlignment = ProductPrintSettings.LabelLayoutSettings.AlignTextYCenter ? StringAlignment.Center : StringAlignment.Near
                    };

                    var x = ProductPrintSettings.LabelLayoutSettings.AlignTextXCenter ? 0 : (float)(ProductPrintSettings.LabelLayoutSettings.TextPositionX / 25.4 * resolution * magnitude);
                    var y = ProductPrintSettings.LabelLayoutSettings.AlignTextYCenter ? 0 : stringPosY;

                    // 矩形領域を計算 (文字列を配置する領域)
                    var layoutRect = new RectangleF(x, y, labelImage.Width - x, labelImage.Height - y);
                    g.DrawString(text, fnt, Brushes.Black, layoutRect, sf);

                    // プレビュー時、黒枠を描画
                    if (isPreview) {
                        using var p = new Pen(Color.Black, 3);
                        g.DrawRectangle(p, 0, 0, labelImage.Width - 1, labelImage.Height - 1);
                    }

                    g.Dispose();
                    break;
                case "Barcode":
                    if (ProductPrintSettings == null) { throw new Exception("SettingsBarcodeProがnull"); }
                    SetLabelProperties(ProductPrintSettings.BarcodePageSettings.LabelWidth,
                                        ProductPrintSettings.BarcodePageSettings.LabelHeight,
                                        ProductPrintSettings.BarcodeLayoutSettings.TextPositionY,
                                        ProductPrintSettings.BarcodeLayoutSettings.TextFont.SizeInPoints,
                                        ProductPrintSettings.BarcodeLayoutSettings.TextFont.Name,
                                        false);

                    labelImage = new((int)sizeX, (int)sizeY);
                    g = Graphics.FromImage(labelImage);
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit; // アンチエイリアス処理を改善

                    stringSize = TextRenderer.MeasureText(text, fnt);

                    stringPosX = ProductPrintSettings.BarcodeLayoutSettings.AlignTextXCenter
                        ? (float)((labelImage.Width / 2f) - (stringSize.Width / 2f))
                        : (float)(ProductPrintSettings.BarcodeLayoutSettings.TextPositionX / 25.4 * resolution * magnitude);

                    g.DrawString(text, fnt, Brushes.Black, stringPosX, stringPosY);

                    var barWeight = resolution == _displayResolution ? 1 : (int)(1 * resolution / _displayResolution / _displayMagnitude);

                    using (var img = Code128Rendering.MakeBarcodeImage(text, barWeight, true)) {
                        var imageWidth = img.Width * ProductPrintSettings.BarcodeLayoutSettings.BarcodeMagnitude;

                        if (imageWidth > labelImage.Width) { MessageBox.Show($"バーコードの幅がラベル幅を超えています{imageWidth}>{labelImage.Width}"); }

                        var barCodePosX = ProductPrintSettings.BarcodeLayoutSettings.AlignBarcodeXCenter
                            ? (float)((labelImage.Width / 2) - (imageWidth / 2))
                            : (float)(ProductPrintSettings.BarcodeLayoutSettings.TextPositionX / 25.4 * resolution * magnitude);
                        var barCodePosY = (float)(ProductPrintSettings.BarcodeLayoutSettings.BarcodePositionY / 25.4 * resolution * magnitude);

                        var barcodeHeight = ProductPrintSettings.BarcodeLayoutSettings.BarcodeHeight / 25.4 * resolution * magnitude;

                        g.DrawImage(img, barCodePosX, barCodePosY, (float)imageWidth, (float)barcodeHeight);
                        // プレビュー時、黒枠を描画
                        if (isPreview) {
                            using var p = new Pen(Color.Black, 3);
                            g.DrawRectangle(p, 0, 0, labelImage.Width - 1, labelImage.Height - 1);
                        }

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
                var code = textToUpperCheckBox.Checked ? QrCodeTextBox.Text.ToUpper() : QrCodeTextBox.Text;
                var arr = code.Split(separator, StringSplitOptions.None);
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
                    ("StrOrderNumber", $"{ProductInfo.OrderNumber}"),
                    ("StrProductNumber", $"{ProductInfo.ProductNumber}"),
                    ("StrRevision", $"{ProductInfo.Revision}"),
                    ("IntRegType", $"{ProductInfo.RegType}"),
                    ("StrRegDate", $"{ProductInfo.RegDate}"),
                    ("StrPerson", $"{ProductInfo.Person}"),
                    ("IntQuantity", $"{ProductInfo.Quantity}"),
                    ("IntSerialFirstNumber", $"{ProductInfo.SerialFirstNumber}"),
                    ("IntSerialLastNumber", $"{ProductInfo.SerialLastNumber}"),
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
            PreviewBarcode();
        }
        private void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Barcode";
            PreviewBarcode();
        }
        private void シリアルラベル印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductPrintSettingsWindow ls = new();
            ls.ShowDialog(this);
            LoadSettings(labelSettingFilePath);
        }
        private void バーコード印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductBarcodeSettingsWindow ls = new();
            ls.ShowDialog(this);
            LoadSettings(labelSettingFilePath);
        }
        private void ProductRegistration2PrintPreviewDialog_Load(object sender, EventArgs e) {
            var tool = (ToolStrip)RePrintPrintPreviewDialog.Controls[1];
            tool.Items[0].Visible = false;
        }
        private void QrCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }
    }
}
