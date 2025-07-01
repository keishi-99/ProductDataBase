using ProductDatabase.Other;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using ZXing;
using ZXing.Windows.Compatibility;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class RePrintWindow : Form {

        public ProductPrintSettings ProductPrintSettings { get; set; } = new ProductPrintSettings();
        public LabelPageSettings LabelPageSettings => ProductPrintSettings.LabelPageSettings ?? new LabelPageSettings();
        public BarcodePageSettings BarcodePageSettings => ProductPrintSettings.BarcodePageSettings ?? new BarcodePageSettings();
        public LabelLayoutSettings LabelLayoutSettings => ProductPrintSettings.LabelLayoutSettings ?? new LabelLayoutSettings();
        public BarcodeLayoutSettings BarcodeLayoutSettings => ProductPrintSettings.BarcodeLayoutSettings ?? new BarcodeLayoutSettings();

        public string printSettingPath = string.Empty;

        public ProductInformation ProductInfo { get; }

        private int _remainingCount;
        private int _labelProNSerial;
        private int _labelProNumLabelsToPrint;

        private const float MmPerInch = 25.4f;
        private const float PointsPerInch = 72.0f;

        private int _pageCount = 1;
        private System.Drawing.Printing.PrintAction _printAction;
        private string _printerName = string.Empty;

        private string _serialType = string.Empty;
        private string _strSerialFirstNumber = string.Empty;
        private string _strSerialLastNumber = string.Empty;
        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "ExtraCheckBox1",
                    "RevisionCheckBox", "ExtraCheckBox2", "ExtraCheckBox3", "FirstSerialNumberCheckBox", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" ];

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
                FirstSerialNumberTextBox.Text = ProductInfo.SerialDigit switch {
                    3 => "000",
                    4 => "0000",
                    _ => string.Empty
                };

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
                // 印刷UI設定
                ConfigurePrintSettings();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }
        // 印刷UI設定
        private void ConfigurePrintSettings() {
            FirstSerialNumberCheckBox.Checked = true;
            LabelPrintButton.Enabled = ProductInfo.IsLabelPrint;
            BarcodePrintButton.Enabled = ProductInfo.IsBarcodePrint;

            シリアルラベル印刷プレビューToolStripMenuItem.Enabled = ProductInfo.IsLabelPrint;
            シリアルラベル印刷設定ToolStripMenuItem.Enabled = ProductInfo.IsLabelPrint;
            バーコード印刷プレビューToolStripMenuItem.Enabled = ProductInfo.IsBarcodePrint;
            バーコード印刷設定ToolStripMenuItem.Enabled = ProductInfo.IsBarcodePrint;

            LoadSettings();
        }
        private void LoadSettings() {
            try {
                ProductPrintSettings = new ProductPrintSettings();
                printSettingPath = Path.Combine(Environment.CurrentDirectory, "config", "Product", ProductInfo.CategoryName, ProductInfo.ProductName, $"PrintConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.json");
                if (!File.Exists(printSettingPath)) { throw new DirectoryNotFoundException($"ラベル印刷用設定ファイルがありません。"); }
                var jsonString = File.ReadAllText(printSettingPath);
                ProductPrintSettings = System.Text.Json.JsonSerializer.Deserialize<ProductPrintSettings>(jsonString) ?? new ProductPrintSettings();
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 登録処理
        private void RegisterCheck(bool print) {
            try {
                FormCheck();
                if (!DataCheck()) { return; }

                if (print) {
                    DialogResult result;
                    result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.Cancel) { return; }

                    result = MessageBox.Show("同一のシリアルラベルが複数存在しないようにして下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.Cancel) { return; }

                    ProductInfo.Person = PersonComboBox.Text;
                    if (!Registeration()) { throw new Exception("登録できませんでした。"); }
                }

                if (!Print(print)) { throw new Exception("キャンセルしました。"); }

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
                : true;
        }
        private bool DataCheck() {
            var revision = RevisionTextBox.Text.Trim();
            if (RevisionCheckBox.Checked) {
                // revision.Any(...) は、revision 内のいずれかの文字が条件を満たす場合に true を返します。
                // char.ToUpperInvariant(c) は、文字を大文字に変換し、比較を行います。
                if (revision.Any(c => "IO".Contains(char.ToUpperInvariant(c)))) {
                    MessageBox.Show("Revisionに I, O は使用できません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    RevisionTextBox.Focus();
                    return false;
                }
            }

            if (string.IsNullOrWhiteSpace(QuantityTextBox.Text)) {
                MessageBox.Show("数量を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                QuantityTextBox.Focus();
                return false;
            }
            if (!int.TryParse(QuantityTextBox.Text, out var quantity)) {
                MessageBox.Show("数量は有効な数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                QuantityTextBox.Focus();
                return false;
            }
            if (quantity <= 0) {
                MessageBox.Show("1台以上入力して下さい。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                QuantityTextBox.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(FirstSerialNumberTextBox.Text)) {
                MessageBox.Show("シリアル開始番号を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                FirstSerialNumberTextBox.Focus();
                return false;
            }
            if (!int.TryParse(FirstSerialNumberTextBox.Text, out var firstSerial)) {
                MessageBox.Show("シリアル開始番号が不正な形式です。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                FirstSerialNumberTextBox.Focus();
                return false;
            }
            if (firstSerial <= 0) {
                MessageBox.Show("1以上入力して下さい。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                FirstSerialNumberTextBox.Focus();
                return false;
            }

            // 最終シリアル番号計算
            var calculatedLastSerial = quantity + firstSerial - 1; // 数量と開始番号から最終シリアルを算出

            // シリアル番号の桁数に応じて、閾値とリセット値を設定
            (var minNumber, var maxNumber, var digit) = ProductInfo.SerialType switch {
                3 => (1, 999, 3),
                4 => (1, 9999, 4),
                101 => (1, 899, 3),
                102 => (901, 999, 3),
                _ => throw new InvalidOperationException("不明なシリアル桁数です。") // より具体的な例外
            };

            if (calculatedLastSerial > maxNumber || firstSerial < minNumber) {// あるいは firstSerialがminNumber未満の場合も対象に
                MessageBox.Show($"シリアルが範囲外になるため、{minNumber.ToString().PadLeft(digit, '0')}から開始します。", "シリアル番号リセット", MessageBoxButtons.OK, MessageBoxIcon.Information);
                FirstSerialNumberTextBox.Text = minNumber.ToString();
            }

            ProductInfo.OrderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
            ProductInfo.ProductNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
            ProductInfo.Quantity = quantity;
            ProductInfo.Person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
            ProductInfo.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateMaskedTextBox.Text : string.Empty;
            ProductInfo.Revision = RevisionCheckBox.Checked ? RevisionTextBox.Text : string.Empty;
            ProductInfo.Comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;

            ProductInfo.SerialFirstNumber = firstSerial;
            ProductInfo.SerialLastNumber = ProductInfo.SerialFirstNumber + ProductInfo.Quantity - 1;

            _strSerialFirstNumber = GenerateCode(ProductInfo.SerialFirstNumber);
            _strSerialLastNumber = GenerateCode(ProductInfo.SerialLastNumber);
            return true;
        }
        // 印刷処理
        private bool Print(bool print) {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();

                pd.BeginPrint += (sender, e) => _printAction = e.PrintAction;
                pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

                _printerName = pd.PrinterSettings.PrinterName;
                _labelProNumLabelsToPrint = ProductInfo.Quantity;
                _pageCount = 1;

                switch (print) {
                    case true:
                        //PrintDialogクラスの作成
                        var pdlg = new PrintDialog {
                            Document = pd
                        };
                        if (pdlg.ShowDialog() == DialogResult.OK) {
                            // ローディング画面の表示
                            using var loadingForm = new LoadingForm();
                            // 別スレッドで印刷処理を実行
                            Task.Run(() => {
                                try {
                                    pd.Print();
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
                    case false:
                        //PrintPreviewDialogオブジェクトの作成
                        var ppd = new PrintPreviewDialog();
                        ppd.Shown += (sender, e) => {
                            var tool = (ToolStrip)ppd.Controls[1];
                            tool.Items[0].Visible = false;
                            if (sender is Form form) {
                                form.WindowState = FormWindowState.Maximized;
                            }
                        };
                        ppd.PrintPreviewControl.Zoom = 3;
                        ppd.Document = pd;
                        ppd.ShowDialog();

                        return true;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            var headerString = string.Empty;
            var headerFont = SystemFonts.DefaultFont;
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
            // プレビューかどうかの判定
            var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                var dpiX = e.Graphics.DpiX;
                var dpiY = e.Graphics.DpiY;
                e.Graphics.PageUnit = GraphicsUnit.Millimeter;

                switch (_serialType) {
                    case "Label":
                        labelCountX = LabelPageSettings.LabelsPerColumn;
                        labelCountY = LabelPageSettings.LabelsPerRow;
                        labelWidth = LabelPageSettings.LabelWidth;
                        labelHeight = LabelPageSettings.LabelHeight;
                        marginX = LabelPageSettings.MarginX;
                        marginY = LabelPageSettings.MarginY;
                        intervalX = LabelPageSettings.IntervalX;
                        intervalY = LabelPageSettings.IntervalY;
                        headerPositionX = LabelPageSettings.HeaderPositionX;
                        headerPositionY = LabelPageSettings.HeaderPositionY;
                        headerString = ConvertHeaderString(LabelPageSettings.HeaderTextFormat);
                        headerFont = LabelPageSettings.HeaderFont;
                        copiesPerLabel = LabelLayoutSettings.CopiesPerLabel;
                        startLine = (int)PrintPostionNumericUpDown.Value - 1;
                        break;
                    case "Barcode":
                        labelCountX = BarcodePageSettings.LabelsPerColumn;
                        labelCountY = BarcodePageSettings.LabelsPerRow;
                        labelWidth = BarcodePageSettings.LabelWidth;
                        labelHeight = BarcodePageSettings.LabelHeight;
                        marginX = BarcodePageSettings.MarginX;
                        marginY = BarcodePageSettings.MarginY;
                        intervalX = BarcodePageSettings.IntervalX;
                        intervalY = BarcodePageSettings.IntervalY;
                        headerPositionX = BarcodePageSettings.HeaderPositionX;
                        headerPositionY = BarcodePageSettings.HeaderPositionY;
                        headerString = ConvertHeaderString(BarcodePageSettings.HeaderTextFormat);
                        headerFont = BarcodePageSettings.HeaderFont;
                        copiesPerLabel = BarcodeLayoutSettings.CopiesPerLabel;
                        startLine = (int)PrintPostionNumericUpDown.Value - 1;
                        break;
                    default:
                        break;
                }

                if (labelCountX == 0 || labelCountY == 0 || copiesPerLabel == 0) { throw new Exception("印刷設定が異常です。"); }

                // ハードマージンをミリメートルに変換
                var hardMarginX = 0f;
                var hardMarginY = 0f;
                if (!isPreview) {
                    (hardMarginX, hardMarginY) = _printerName switch {
                        "TOSHIBA Universal Printer 2" => (5.6f, 1.9f),
                        "検査課　複合機" => (5.6f, 1.9f),
                        _ => (e.PageSettings.HardMarginX * MmPerInch / 100, e.PageSettings.HardMarginY * MmPerInch / 100)
                    };
                }

                if (_pageCount == 1) {
                    _remainingCount = copiesPerLabel;
                    _labelProNSerial = ProductInfo.SerialFirstNumber;
                }
                if (_pageCount >= 2) { startLine = 0; }

                // 最初のページのみオフセットを調整
                var verticalOffset = _pageCount == 1 ? startLine * (intervalY + labelHeight) : 0;
                // ヘッダーの描画
                e.Graphics.DrawString(headerString, headerFont, Brushes.Gray, (float)headerPositionX, (float)(verticalOffset + headerPositionY - hardMarginY));

                var y = 0;
                for (y = startLine; y < labelCountY; y++) {
                    var x = 0;
                    for (x = 0; x < labelCountX; x++) {
                        var posX = (float)(marginX - hardMarginX + (x * (intervalX + labelWidth)));
                        var posY = (float)(marginY - hardMarginY + (y * (intervalY + labelHeight)));

                        // タイプ4で残り1の場合、最後のラベルに下線をつける
                        var fontUnderline = ProductInfo.IsUnderlinePrint && _remainingCount == 1;

                        // シリアル生成、PrintTypeが9かつ最終行の場合は型式下4桁、それ以外はシリアルを生成
                        string generatedCode;
                        if (ProductInfo.PrintType != 9 || _remainingCount != 1) {
                            generatedCode = GenerateCode(_labelProNSerial); // シリアルコードを生成
                        }
                        else {
                            generatedCode = ProductInfo.ProductModel[^4..]; // 型式の下4桁を使用
                        }

                        using var labelImage = MakeLabelImage(generatedCode, fontUnderline, dpiX, dpiY);
                        e.Graphics.DrawImage(labelImage, posX, posY, (float)labelWidth, (float)labelHeight);

                        _remainingCount--;
                        if (_remainingCount <= 0) {
                            _labelProNSerial++;
                            _labelProNumLabelsToPrint--;
                            //印刷するラベルがなくなった場合の処理
                            if (_labelProNumLabelsToPrint <= 0) {
                                // 最終行の行番号を表示
                                var sf = new StringFormat {
                                    Alignment = StringAlignment.Near,
                                    LineAlignment = StringAlignment.Center
                                };
                                var layoutRect = new RectangleF(0, posY, 0, (float)labelHeight);
                                var rowNumber = (y + 1).ToString();
                                e.Graphics.DrawString(rowNumber, headerFont, Brushes.Black, layoutRect, sf);

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
                "Label" => LabelLayoutSettings.TextFormat ?? string.Empty,
                "Barcode" => BarcodeLayoutSettings.TextFormat ?? string.Empty,
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
        private Bitmap MakeLabelImage(string text, bool fontUnderline, float dpiX, float dpiY) {

            // プレビューモードかどうかを判定
            var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;

            // --- 2. Bitmapの生成とGraphicsの準備 ---
            (var labelWidth, var labelHeight) = _serialType switch {
                "Label" => (LabelPageSettings.LabelWidth, LabelPageSettings.LabelHeight),
                "Barcode" => (BarcodePageSettings.LabelWidth, BarcodePageSettings.LabelHeight),
                _ => throw new InvalidOperationException("Unsupported serial type.")
            };

            // ビットマップのサイズをピクセル単位で計算
            var pixelWidth = (int)(labelWidth / MmPerInch * dpiX);
            var pixelHeight = (int)(labelHeight / MmPerInch * dpiY);

            var labelImage = new Bitmap(pixelWidth, pixelHeight);
            labelImage.SetResolution(dpiX, dpiY);

            // 'using'ステートメントでGraphicsオブジェクトを確実に破棄
            using (var g = Graphics.FromImage(labelImage)) {
                // すべての描画操作をミリメートル単位で行うように設定
                g.PageUnit = GraphicsUnit.Millimeter;

                // 高品質な描画設定
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

                // --- 3. 印刷タイプに応じた描画処理 ---

                switch (_serialType) {
                    case "Label":
                        DrawLabel(g, text, fontUnderline);
                        break;

                    case "Barcode":
                        DrawBarcode(g, text);
                        break;
                }

                // --- 4. プレビュー用の枠を描画 ---

                if (isPreview) {
                    // 0.1mmの黒いペンで枠線を描画
                    using var p = new Pen(Color.Black, 0.1f);
                    g.DrawRectangle(p, 0, 0, (float)labelWidth - 0.1f, (float)labelHeight - 0.1f);
                }
            }

            return labelImage;
        }
        private void DrawLabel(Graphics g, string text, bool fontUnderline) {

            // フォントサイズはポイント単位でそのまま使用
            var fontName = LabelLayoutSettings.TextFont.Name;
            var fontSize = LabelLayoutSettings.TextFont.SizeInPoints;
            var style = fontUnderline ? FontStyle.Underline : FontStyle.Regular;

            using var textFont = new Font(fontName, fontSize, style);
            // テキストの配置設定
            using var sf = new StringFormat {
                Alignment = LabelLayoutSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near,
                LineAlignment = LabelLayoutSettings.AlignTextCenterY ? StringAlignment.Center : StringAlignment.Near
            };
            // 描画領域をミリメートル単位で計算
            var textPosX = LabelLayoutSettings.AlignTextCenterX ? 0f : (float)LabelLayoutSettings.TextPositionX;
            var textPosY = LabelLayoutSettings.AlignTextCenterY ? 0f : (float)LabelLayoutSettings.TextPositionY;

            var pageWidth = (float)LabelPageSettings.LabelWidth;
            var pageHeight = (float)LabelPageSettings.LabelHeight;

            var layoutRect = new RectangleF(textPosX, textPosY, pageWidth - textPosX, pageHeight - textPosY);

            g.DrawString(text, textFont, Brushes.Black, layoutRect, sf);
        }
        private void DrawBarcode(Graphics g, string text) {

            var pageWidth = (float)BarcodePageSettings.LabelWidth;
            var pageHeight = (float)BarcodePageSettings.LabelHeight;

            // --- テキストの描画 ---
            // フォントサイズはポイント単位でそのまま使用
            var fontName = BarcodeLayoutSettings.TextFont.Name;
            var fontSize = BarcodeLayoutSettings.TextFont.SizeInPoints;

            using (var textFont = new Font(fontName, fontSize, FontStyle.Regular)) {
                // テキストの配置設定
                using var sf = new StringFormat {
                    Alignment = BarcodeLayoutSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near
                };

                var textPosX = BarcodeLayoutSettings.AlignTextCenterX ? 0f : (float)BarcodeLayoutSettings.TextPositionX;
                var textPosY = (float)BarcodeLayoutSettings.TextPositionY;

                var layoutRectString = new RectangleF(textPosX, textPosY, pageWidth - textPosX, pageHeight - textPosY);

                g.DrawString(text, textFont, Brushes.Black, layoutRectString, sf);
            }

            // --- バーコードの描画 ---

            // ZXingはピクセル単位で画像を生成するため、mmからpixelへの変換が必要
            // GraphicsオブジェクトのDPI（g.DpiX）を使用して計算する
            var barcodePixelWidth = (int)(BarcodeLayoutSettings.BarcodeWidth / MmPerInch * g.DpiX);
            var barcodePixelHeight = (int)(BarcodeLayoutSettings.BarcodeHeight / MmPerInch * g.DpiY);

            var writer = new BarcodeWriter<Bitmap> {
                Format = BarcodeFormat.CODE_128,
                Options = new ZXing.Common.EncodingOptions {
                    Height = barcodePixelHeight,
                    Width = barcodePixelWidth,
                    PureBarcode = true // テキストを含まないバーコードのみを生成
                },
                Renderer = new BitmapRenderer()
            };

            using var barcodeBitmap = writer.Write(text);
            var barcodePosX = (float)BarcodeLayoutSettings.BarcodePositionX;
            var barcodePosY = (float)BarcodeLayoutSettings.BarcodePositionY;
            var barcodeWidth = (float)BarcodeLayoutSettings.BarcodeWidth;
            var barcodeHeight = (float)BarcodeLayoutSettings.BarcodeHeight;

            // X座標を中央に調整 (ミリメートル単位で計算)
            if (BarcodeLayoutSettings.AlignBarcodeCenterX) {
                barcodePosX = (pageWidth / 2f) - (barcodeWidth / 2f);
            }

            var layoutRectBarcode = new RectangleF(barcodePosX, barcodePosY, barcodeWidth, barcodeHeight);
            // g.DrawImageでミリメートル単位の座標とサイズを指定して描画
            g.DrawImage(barcodeBitmap, layoutRectBarcode);
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
            RegisterCheck(true);
        }
        private void BarcodePrintButton_Click(object sender, EventArgs e) {
            _serialType = "Barcode";
            RegisterCheck(true);
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
            RegisterCheck(false);
        }
        private void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Barcode";
            RegisterCheck(false);
        }
        private void シリアルラベル印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductPrintSettingsWindow ls = new() {
                ProductInfo = ProductInfo
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void バーコード印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            ProductBarcodeSettingsWindow ls = new() {
                ProductInfo = ProductInfo
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void QrCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }

    }
}
