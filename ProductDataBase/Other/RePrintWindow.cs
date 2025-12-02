using Microsoft.Data.Sqlite;
using ProductDatabase.Other;
using ProductDatabase.Print;
using static ProductDatabase.MainWindow;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase {
    public partial class RePrintWindow : Form {

        public DocumentPrintSettings ProductPrintSettings { get; set; } = new DocumentPrintSettings();
        public LabelPrintSettings LabelPrintSettings => this.ProductPrintSettings.LabelPrintSettings ?? new LabelPrintSettings();
        public BarcodePrintSettings BarcodePrintSettings => ProductPrintSettings.BarcodePrintSettings ?? new BarcodePrintSettings();
        public NameplatePrintSettings NameplatePrintSettings => ProductPrintSettings.NameplatePrintSettings ?? new NameplatePrintSettings();

        public string printSettingPath = string.Empty;

        public ProductInformation ProductInfo { get; }

        private string _serialType = string.Empty;
        private string _serialFirstNumber = string.Empty;
        private string _serialLastNumber = string.Empty;
        private readonly List<string> _serialList = [];
        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox",  "FirstSerialNumberCheckBox",
                    "RevisionCheckBox", "ExtraCheckBox1", "ExtraCheckBox2", "ExtraCheckBox3", "RegistrationDateCheckBox",
                    "PersonCheckBox", "CommentCheckBox" ];

        public RePrintWindow(ProductInformation productInfo) {
            InitializeComponent();
            ProductInfo = productInfo;
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new System.Drawing.Font(ProductInfo.FontName, ProductInfo.FontSize);

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

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. ProductInfo.PersonList]);

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using (SqliteConnection con = new(GetConnectionRegistration())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - [ProductName]テーブルの[SubstrateModel]列の[Revision]を取得
                    var productTableName = $"[T_{ProductInfo.CategoryName}_Product]";
                    cmd.CommandText = $"SELECT Revision FROM {productTableName} WHERE ProductName = @ProductName AND RevisionGroup = @RevisionGroup ORDER BY ID DESC;";
                    cmd.Parameters.Add("@ProductName", SqliteType.Text).Value = ProductInfo.ProductName;
                    cmd.Parameters.Add("@RevisionGroup", SqliteType.Text).Value = ProductInfo.RevisionGroup;
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
            NameplatePrintButton.Enabled = ProductInfo.IsNameplatePrint;

            シリアルラベル印刷プレビューToolStripMenuItem.Enabled = ProductInfo.IsLabelPrint;
            シリアルラベル印刷設定ToolStripMenuItem.Enabled = ProductInfo.IsLabelPrint;
            バーコード印刷プレビューToolStripMenuItem.Enabled = ProductInfo.IsBarcodePrint;
            バーコード印刷設定ToolStripMenuItem.Enabled = ProductInfo.IsBarcodePrint;
            銘版印刷設定ToolStripMenuItem.Enabled = ProductInfo.IsNameplatePrint;

            LoadSettings();
        }
        private void LoadSettings() {
            try {
                ProductPrintSettings = new DocumentPrintSettings();
                printSettingPath = Path.Combine(Environment.CurrentDirectory, "config", "Product", ProductInfo.CategoryName, ProductInfo.ProductName, $"PrintConfig_{ProductInfo.ProductName}_{ProductInfo.ProductModel}.json");
                if (!File.Exists(printSettingPath)) { throw new DirectoryNotFoundException($"ラベル印刷用設定ファイルがありません。"); }
                var jsonString = File.ReadAllText(printSettingPath);
                ProductPrintSettings = System.Text.Json.JsonSerializer.Deserialize<DocumentPrintSettings>(jsonString) ?? new DocumentPrintSettings();
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 登録処理
        private void RegisterCheck(bool isPrint) {
            try {
                FormCheck();
                if (!DataCheck()) { return; }

                if (isPrint) {
                    DialogResult result;
                    result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.Cancel) { return; }

                    result = MessageBox.Show("同一のシリアルラベルが複数存在しないようにして下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.Cancel) { return; }

                    ProductInfo.Person = PersonComboBox.Text;
                    if (!Registration()) { throw new Exception("登録できませんでした。"); }
                }

                switch (_serialType) {
                    case "Nameplate":
                        PrintManager.PrintUsingBPac(NameplatePrintSettings, _serialList);
                        break;
                    default:
                        if (!Print(isPrint)) { throw new Exception("キャンセルしました。"); }
                        break;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool Registration() {
            try {
                // 再印刷登録テーブルへ追加
                using SqliteConnection con = new(GetConnectionRegistration());
                con.Open();
                using var cmd = con.CreateCommand();
                cmd.CommandText =
                    $"""
                    INSERT INTO T_Reprint (
                        SerialPrintType, OrderNumber, ProductName, ProductNumber, ProductType, ProductModel, Quantity, Person, RegDate, Revision, SerialFirst, SerialLast, Comment
                        )
                    VALUES (
                        @SerialPrintType, @OrderNumber, @ProductName, @ProductNumber, @ProductType, @ProductModel, @Quantity, @Person, @RegDate, @Revision, @SerialFirst, @SerialLast, @Comment
                        )
                    ;
                    """;

                // チェックボックスにチェックがない場合はNullを
                cmd.Parameters.Add("@SerialPrintType", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_serialType) ? DBNull.Value : _serialType;
                cmd.Parameters.Add("@OrderNumber", SqliteType.Text).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                cmd.Parameters.Add("@ProductName", SqliteType.Text).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductName) ? DBNull.Value : ProductInfo.ProductName;
                cmd.Parameters.Add("@ProductNumber", SqliteType.Text).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                cmd.Parameters.Add("@ProductType", SqliteType.Text).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                cmd.Parameters.Add("@ProductModel", SqliteType.Text).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                cmd.Parameters.Add("@Quantity", SqliteType.Text).Value = ProductInfo.Quantity;
                cmd.Parameters.Add("@Person", SqliteType.Text).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                cmd.Parameters.Add("@RegDate", SqliteType.Text).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                cmd.Parameters.Add("@Revision", SqliteType.Text).Value = string.IsNullOrWhiteSpace(ProductInfo.Revision) ? DBNull.Value : ProductInfo.Revision;
                cmd.Parameters.Add("@SerialFirst", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_serialFirstNumber) ? DBNull.Value : _serialFirstNumber;
                cmd.Parameters.Add("@SerialLast", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_serialLastNumber) ? DBNull.Value : _serialLastNumber;
                cmd.Parameters.Add("@Comment", SqliteType.Text).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                cmd.ExecuteNonQuery();

                // バックアップ作成
                CommonUtils.BackupManager.CreateBackup();
                // ログ出力
                string[] logMessageArray = [
                    $"[再印刷]",
                    $"[{ProductInfo.CategoryName}]",
                    $"[]",
                    $"注文番号[{ProductInfo.OrderNumber}]",
                    $"製造番号[{ProductInfo.ProductNumber}]",
                    $"製品名[{ProductInfo.ProductName}]",
                    $"タイプ[{ProductInfo.ProductType}]",
                    $"型式[{ProductInfo.ProductModel}]",
                    $"数量[{ProductInfo.Quantity}]",
                    $"シリアル先頭[{_serialFirstNumber}]",
                    $"シリアル末尾[{_serialLastNumber}]",
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
            string manufacturingNumber = ManufacturingNumberMaskedTextBox.Text.Trim();
            return !anyTextBoxEnabled
                ? throw new Exception("何も入力されていません")
                : !allTextBoxesFilled
                ? throw new Exception("空欄があります。")
                : ManufacturingNumberCheckBox.Checked &&
                  !manufacturingNumber.StartsWith("R", StringComparison.OrdinalIgnoreCase) &&
                  manufacturingNumber.Length != 15
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

            if (calculatedLastSerial > maxNumber || firstSerial < minNumber) {// あるいは firstSerialが minNumber未満の場合も対象に
                MessageBox.Show($"シリアルが範囲外になるため、{minNumber.ToString().PadLeft(digit, '0')}から開始します。", "シリアル番号リセット", MessageBoxButtons.OK, MessageBoxIcon.Information);
                FirstSerialNumberTextBox.Text = minNumber.ToString();
            }

            ProductInfo.OrderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
            ProductInfo.ProductNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
            ProductInfo.Quantity = quantity;
            ProductInfo.Person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
            ProductInfo.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateTimePicker.Value.ToShortDateString() : string.Empty;
            ProductInfo.Revision = RevisionCheckBox.Checked ? RevisionTextBox.Text : string.Empty;
            ProductInfo.Comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;

            ProductInfo.SerialFirstNumber = firstSerial;
            ProductInfo.SerialLastNumber = ProductInfo.SerialFirstNumber + ProductInfo.Quantity - 1;

            _serialList.Clear();

            for (var i = 0; i < quantity; i++) {
                _serialList.Add(GenerateCode(ProductInfo.SerialFirstNumber + i));
            }

            _serialFirstNumber = GenerateCode(ProductInfo.SerialFirstNumber);
            _serialLastNumber = GenerateCode(ProductInfo.SerialLastNumber);
            return true;
        }

        // 印刷処理
        private bool Print(bool isPrint) {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();
                var isPreview = !isPrint;
                var serialType = _serialType;

                var startLine = (int)PrintPositionNumericUpDown.Value - 1;

                pd.BeginPrint += (sender, e) => {
                    PrintManager.Initialize(ProductInfo, ProductPrintSettings, _serialList);
                };
                pd.PrintPage += (sender, e) => {
                    bool hasMore = PrintManager.PrintSerialCommon(e, isPreview, startLine, serialType);
                    e.HasMorePages = hasMore;
                };

                switch (isPrint) {
                    case true:
                        // PrintDialogクラスの作成
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
                                    loadingForm.Invoke(new System.Action(() => loadingForm.Close()));
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
                        // PrintPreviewDialogオブジェクトの作成
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
        private string GenerateCode(int serialCode) {
            var monthCode = DateTime.Parse(ProductInfo.RegDate).ToString("MM");

            monthCode = monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => monthCode
            };

            var outputCode = _serialType switch {
                "Label" => LabelPrintSettings.TextFormat ?? string.Empty,
                "Barcode" => BarcodePrintSettings.TextFormat ?? string.Empty,
                "Nameplate" => NameplatePrintSettings.TextFormat ?? string.Empty,
                _ => string.Empty
            };

            var regDate = DateTime.Parse(ProductInfo.RegDate);

            var map = new Dictionary<string, string> {
                ["{T}"] = ProductInfo.Initial,
                ["{Y}"] = regDate.ToString("yy"),
                ["{MM}"] = regDate.ToString("MM"),
                ["{R}"] = ProductInfo.Revision,
                ["{M}"] = monthCode[^1..],
                ["{S}"] = Convert.ToInt32(serialCode).ToString($"D{ProductInfo.SerialDigit}")
            };

            foreach (var kv in map) {
                outputCode = outputCode.Replace(kv.Key, kv.Value);
            }

            return outputCode;
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
                    RNumberCheckBox.Enabled = checkBox.Checked;
                    break;
                case "RNumberCheckBox":
                    ManufacturingNumberMaskedTextBox.Mask = checkBox.Checked ? "R00000000000000" : ">LA00A00000-0000";
                    break;
                case "QuantityCheckBox":
                    QuantityTextBox.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox1":
                    ExtraTextBox1.Enabled = checkBox.Checked;
                    break;
                case "RevisionCheckBox":
                    RevisionTextBox.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox2":
                    ExtraTextBox2.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox3":
                    ExtraTextBox3.Enabled = checkBox.Checked;
                    break;
                case "FirstSerialNumberCheckBox":
                    FirstSerialNumberTextBox.Enabled = checkBox.Checked;
                    break;
                case "RegistrationDateCheckBox":
                    RegistrationDateTimePicker.Enabled = checkBox.Checked;
                    break;
                case "PersonCheckBox":
                    PersonComboBox.Enabled = checkBox.Checked;
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
                if (arr is not null) {
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
        // 取得情報表示
        private void ShowInfo() {
            var items = new Dictionary<string, string>
                {
                    {"ProductName", $"{ProductInfo.ProductName}"},
                    {"ProductModel", $"{ProductInfo.ProductModel}"},
                    {"ProductType", $"{ProductInfo.ProductType}"},
                    {"OrderNumber", $"{ProductInfo.OrderNumber}"},
                    {"ProductNumber", $"{ProductInfo.ProductNumber}"},
                    {"Revision", $"{ProductInfo.Revision}"},
                    {"RegType", $"{ProductInfo.RegType}"},
                    {"RegDate", $"{ProductInfo.RegDate}"},
                    {"Person", $"{ProductInfo.Person}"},
                    {"Quantity", $"{ProductInfo.Quantity}"},
                    {"SerialFirstNumber", $"{ProductInfo.SerialFirstNumber}"},
                    {"SerialLastNumber", $"{ProductInfo.SerialLastNumber}"},
                    {"Initial", $"{ProductInfo.Initial}"},
                    {"SerialPrintType", $"{ProductInfo.SerialPrintType}"},
                    {"SheetPrintType", $"{ProductInfo.SheetPrintType}"},
                    {"SerialDigit", $"{ProductInfo.SerialDigit}"}
                };

            var form = new Form {
                Text = "取得情報",
                Width = 300,
                Height = 400,
                AutoSize = true,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                ShowInTaskbar = false,
                MaximizeBox = false,
                MinimizeBox = false
            };

            // ListView作成
            var listView = new ListView {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Font = new Font("PlemolJP", ProductInfo.FontSize),
            };

            // 列の追加
            listView.Columns.Add("", 0);   // 値列の幅（調整可）
            listView.Columns.Add("項目", 200, HorizontalAlignment.Right);  // 項目列の幅
            listView.Columns.Add("値", 360);   // 値列の幅（調整可）

            // データを追加
            foreach (var kvp in items) {
                var item = new ListViewItem("");  // ダミー1列目
                item.SubItems.Add(kvp.Key);
                item.SubItems.Add(kvp.Value);
                listView.Items.Add(item);
            }
            form.Controls.Add(listView);

            // フォームのイベントハンドラ
            form.Shown += (_, _) => {
                listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            };
            form.ShowDialog();
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
        private void NamePlatePrintButton_Click(object sender, EventArgs e) {
            _serialType = "Nameplate";
            RegisterCheck(true);
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) { ShowInfo(); }
        private void シリアルラベル印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Label";
            RegisterCheck(false);
        }
        private void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _serialType = "Barcode";
            RegisterCheck(false);
        }
        private void シリアルラベル印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            PrintSettingsWindow ls = new() {
                ProductInfo = ProductInfo,
                serialType = "Label"
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void バーコード印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            PrintSettingsWindow ls = new() {
                ProductInfo = ProductInfo,
                serialType = "Barcode"
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void 銘版印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            PrintSettingsWindow ls = new() {
                ProductInfo = ProductInfo,
                serialType = "Nameplate"
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void QrCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }

    }
}
