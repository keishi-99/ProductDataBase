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
        public string printSettingPath = string.Empty;

        public ProductInformation ProductInfo { get; }

        private int _remainingCount;
        private int _labelProNSerial;
        private int _labelProNumLabelsToPrint;

        private const float MmPerInch = 25.4F;
        private const float PointsPerInch = 72.0F;

        private int _pageCount = 1;
        private System.Drawing.Printing.PrintAction _printAction;

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

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;

                switch (_serialType) {
                    case "Label":
                        if (ProductPrintSettings == null) { throw new Exception("ProductPrintSettingsがnullです。"); }
                        if (ProductPrintSettings.LabelPageSettings != null) {
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
                            headerString = ConvertHeaderString(ProductPrintSettings.LabelPageSettings.HeaderTextFormat);
                            headerFont = ProductPrintSettings.LabelPageSettings.HeaderFont;
                        }
                        if (ProductPrintSettings.LabelLayoutSettings != null) {
                            copiesPerLabel = ProductPrintSettings.LabelLayoutSettings.CopiesPerLabel;
                        }
                        startLine = (int)PrintPostionNumericUpDown.Value - 1;
                        break;
                    case "Barcode":
                        if (ProductPrintSettings == null) { throw new Exception("ProductPrintSettingsがnullです。"); }
                        if (ProductPrintSettings.BarcodePageSettings != null) {
                            labelCountX = ProductPrintSettings.BarcodePageSettings.LabelsPerColumn;
                            labelCountY = ProductPrintSettings.BarcodePageSettings.LabelsPerRow;
                            labelWidth = ProductPrintSettings.BarcodePageSettings.LabelWidth;
                            labelHeight = ProductPrintSettings.BarcodePageSettings.LabelHeight;
                            marginX = ProductPrintSettings.BarcodePageSettings.MarginX;
                            marginY = ProductPrintSettings.BarcodePageSettings.MarginY;
                            intervalX = ProductPrintSettings.BarcodePageSettings.IntervalX;
                            intervalY = ProductPrintSettings.BarcodePageSettings.IntervalY;
                            headerPositionX = ProductPrintSettings.BarcodePageSettings.HeaderPositionX;
                            headerPositionY = ProductPrintSettings.BarcodePageSettings.HeaderPositionY;
                            headerString = ConvertHeaderString(ProductPrintSettings.BarcodePageSettings.HeaderTextFormat);
                            headerFont = ProductPrintSettings.BarcodePageSettings.HeaderFont;
                        }
                        if (ProductPrintSettings.BarcodeLayoutSettings != null) {
                            copiesPerLabel = ProductPrintSettings.BarcodeLayoutSettings.CopiesPerLabel;
                        }
                        startLine = (int)PrintPostionNumericUpDown.Value - 1;
                        break;
                    default:
                        break;
                }

                if (labelCountX == 0 || labelCountY == 0 || copiesPerLabel == 0) { throw new Exception("印刷設定が異常です。"); }

                // ハードマージンをミリメートルに変換
                const double MM_PER_HUNDREDTH_INCH = 0.254;
                var hardMarginX = isPreview ? 0 : e.PageSettings.HardMarginX * MM_PER_HUNDREDTH_INCH;
                var hardMarginY = isPreview ? 0 : e.PageSettings.HardMarginY * MM_PER_HUNDREDTH_INCH;

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
                        var dpiX = e.Graphics.DpiX;
                        using var labelImage = MakeLabelImage(generatedCode, dpiX, fontUnderline);
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
                "Label" => ProductPrintSettings.LabelLayoutSettings?.TextFormat ?? string.Empty,
                "Barcode" => ProductPrintSettings.BarcodeLayoutSettings?.TextFormat ?? string.Empty,
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
        private Bitmap MakeLabelImage(string text, float dpi, bool fontUnderline) {
            Bitmap labelImage = new(1, 1);
            Graphics g;
            SizeF stringSize;
            var sizeX = 0;
            var sizeY = 0;
            float fontSize;
            float textPosX = 0;
            float textPosY = 0;
            var textFont = SystemFonts.DefaultFont;
            // プレビューかどうかの判定
            var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;

            switch (_serialType) {
                case "Label":
                    if (ProductPrintSettings == null) {
                        throw new Exception("ProductPrintSettingsがnull");
                    }
                    if (ProductPrintSettings.LabelPageSettings != null) {
                        sizeX = (int)(ProductPrintSettings.LabelPageSettings.LabelWidth / MmPerInch * dpi);
                        sizeY = (int)(ProductPrintSettings.LabelPageSettings.LabelHeight / MmPerInch * dpi);
                    }
                    if (ProductPrintSettings.LabelLayoutSettings != null) {
                        textPosX = ProductPrintSettings.LabelLayoutSettings.TextPositionX / MmPerInch * dpi;
                        textPosY = ProductPrintSettings.LabelLayoutSettings.TextPositionY / MmPerInch * dpi;
                        fontSize = ProductPrintSettings.LabelLayoutSettings.TextFont.SizeInPoints / PointsPerInch * dpi;
                        var style = fontUnderline ? FontStyle.Underline : FontStyle.Regular;
                        textFont = new Font(ProductPrintSettings.LabelLayoutSettings.TextFont.Name, fontSize, style);
                    }

                    labelImage = new(sizeX, sizeY);

                    using (g = Graphics.FromImage(labelImage)) {
                        // アンチエイリアス処理を改善
                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                        // StringFormat を使用して中心に配置
                        var sf = new StringFormat {
                            Alignment = ProductPrintSettings?.LabelLayoutSettings?.AlignTextXCenter ?? false ? StringAlignment.Center : StringAlignment.Near,
                            LineAlignment = ProductPrintSettings?.LabelLayoutSettings?.AlignTextYCenter ?? false ? StringAlignment.Center : StringAlignment.Near
                        };

                        var x = ProductPrintSettings?.LabelLayoutSettings?.AlignTextXCenter ?? false ? 0 : textPosX;
                        var y = ProductPrintSettings?.LabelLayoutSettings?.AlignTextYCenter ?? false ? 0 : textPosY;
                        // 矩形領域を計算 (文字列を配置する領域)
                        var layoutRect = new RectangleF(x, y, labelImage.Width - x, labelImage.Height - y);
                        g.DrawString(text, textFont, Brushes.Black, layoutRect, sf);

                        // プレビュー時、黒枠を描画
                        if (isPreview) {
                            using var p = new Pen(Color.Black, 3);
                            g.DrawRectangle(p, 0, 0, labelImage.Width - 1, labelImage.Height - 1);
                        }
                    }
                    break;
                case "Barcode":
                    float barcodePosX = 0;
                    float barcodePosY = 0;
                    var barcodeHeight = 1;
                    var barcodeWidth = 1;
                    if (ProductPrintSettings == null) {
                        throw new Exception("SettingsBarcodeProがnull");
                    }
                    if (ProductPrintSettings.BarcodePageSettings != null) {
                        sizeX = (int)(ProductPrintSettings.BarcodePageSettings.LabelWidth / MmPerInch * dpi);
                        sizeY = (int)(ProductPrintSettings.BarcodePageSettings.LabelHeight / MmPerInch * dpi);
                    }
                    if (ProductPrintSettings.BarcodeLayoutSettings != null) {
                        textPosX = ProductPrintSettings.BarcodeLayoutSettings.TextPositionX / MmPerInch * dpi;
                        textPosY = ProductPrintSettings.BarcodeLayoutSettings.TextPositionY / MmPerInch * dpi;
                        fontSize = ProductPrintSettings.BarcodeLayoutSettings.TextFont.SizeInPoints / PointsPerInch * dpi;
                        textFont = new Font(ProductPrintSettings.BarcodeLayoutSettings.TextFont.Name, fontSize, FontStyle.Regular);
                        barcodePosX = ProductPrintSettings.BarcodeLayoutSettings.BarcodePositionX / MmPerInch * dpi;
                        barcodePosY = ProductPrintSettings.BarcodeLayoutSettings.BarcodePositionY / MmPerInch * dpi;
                        barcodeHeight = (int)(ProductPrintSettings.BarcodeLayoutSettings.BarcodeHeight / MmPerInch * dpi);
                        barcodeWidth = (int)(ProductPrintSettings.BarcodeLayoutSettings.BarcodeWidth / MmPerInch * dpi);
                    }

                    labelImage = new(sizeX, sizeY);

                    using (g = Graphics.FromImage(labelImage)) {
                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                        // テキストを描画
                        stringSize = TextRenderer.MeasureText(text, textFont);

                        // X座標の調整
                        if (ProductPrintSettings?.BarcodeLayoutSettings?.AlignTextXCenter ?? false) {
                            textPosX = (labelImage.Width / 2) - (stringSize.Width / 2);
                        }

                        g.DrawString(text, textFont, Brushes.Black, textPosX, textPosY);

                        // ZXingライブラリを使用してバーコードを描画
                        var writer = new BarcodeWriter<Bitmap> {
                            Format = BarcodeFormat.CODE_128,
                            Options = new ZXing.Common.EncodingOptions {
                                Height = barcodeHeight,
                                Width = barcodeWidth,
                                PureBarcode = true
                            },
                            Renderer = new BitmapRenderer()
                        };

                        using (var barcodeBitmap = writer.Write(text)) {
                            // X座標の調整
                            if (ProductPrintSettings?.BarcodeLayoutSettings?.AlignBarcodeXCenter ?? false) {
                                barcodePosX = (labelImage.Width / 2) - (barcodeBitmap.Width / 2);
                            }

                            g.DrawImage(barcodeBitmap, barcodePosX, barcodePosY, barcodeWidth, barcodeHeight);
                        }
                        // プレビュー時、黒枠を描画
                        if (isPreview) {
                            using var p = new Pen(Color.Black, 3);
                            g.DrawRectangle(p, 0, 0, labelImage.Width - 1, labelImage.Height - 1);
                        }
                    }
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
