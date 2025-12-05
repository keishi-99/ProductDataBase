using Microsoft.Data.Sqlite;
using ProductDatabase.ExcelService;
using ProductDatabase.Other;
using ProductDatabase.Print;
using static ProductDatabase.MainWindow;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase {
    public partial class SubstrateRegistrationWindow : Form {

        public ProductInformation ProductInfo { get; }

        public DocumentPrintSettings SubstratePrintSettings { get; set; } = new DocumentPrintSettings();
        public LabelPrintSettings LabelPrintSettings => SubstratePrintSettings.LabelPrintSettings ?? new LabelPrintSettings();
        public readonly string printSettingPath = Path.Combine(Environment.CurrentDirectory, "config", "Substrate", "SubstrateConfig.json");

        private readonly List<string> _serialList = [];
        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "DefectQuantityCheckBox",
                    "ExtraCheckBox7", "ExtraCheckBox1", "ExtraCheckBox2", "ExtraCheckBox3", "RegistrationDateCheckBox",
                    "PersonCheckBox", "CommentCheckBox"
        ];

        // プロパティ設定
        private bool IsLabelPrint => ProductInfo.SerialPrintType is 1;
        private bool IsRegistration => ProductInfo.RegType is 0 or 1;

        public SubstrateRegistrationWindow(ProductInformation productInfo) {
            InitializeComponent();
            ProductInfo = productInfo;
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);

                ProductNameLabel2.Text = ProductInfo.ProductName;

                SubstrateModelLabel2.Text = $"{ProductInfo.SubstrateName} - {ProductInfo.SubstrateModel}";

                var stockQuantity = GetStockQuantity();
                StockLabel2.Text = stockQuantity;

                OrderNumberTextBox.Text = ProductInfo.Proness5;
                ManufacturingNumberMaskedTextBox.Text = !string.IsNullOrEmpty(ProductInfo.Proness1) ? ProductInfo.Proness1 : ManufacturingNumberMaskedTextBox.Text;
                QuantityTextBox.Text = (ProductInfo.Proness4 != 0) ? ProductInfo.Proness4.ToString() : string.Empty;

                RegisterButton.Enabled = true;

                CommentComboBox.Items.Add("[Rev.UP]変更点番号:");

                // 変数[check_bin]のビットに応じてCheckboxにチェックを入れる
                for (var i = 0; i < _checkBoxNames.Count; i++) {
                    if (Controls[_checkBoxNames[i]] is CheckBox checkBox) {
                        // i番目のビットが1かどうかをチェック
                        checkBox.Checked = (ProductInfo.CheckBin & (1 << i)) != 0;
                    }
                }

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. ProductInfo.PersonList]);

                // 印刷しない場合は関連コントロール非表示に
                if (IsLabelPrint == false) {
                    PrintRowLabel.Visible = false;
                    PrintPositionNumericUpDown.Visible = false;
                    PrintOnlyCheckBox.Visible = false;
                    PrintButton.Visible = false;
                }

                if (File.Exists(printSettingPath) == false) { throw new Exception("印刷設定ファイルが見つかりませんでした"); }
                SubstratePrintSettings = new DocumentPrintSettings();
                LoadSettings(printSettingPath);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            } finally {
            }
        }
        private void LoadSettings(string settingFilePath) {
            try {
                var jsonString = File.ReadAllText(settingFilePath);
                SubstratePrintSettings = System.Text.Json.JsonSerializer.Deserialize<DocumentPrintSettings>(jsonString) ?? new DocumentPrintSettings();
            } catch (Exception ex) {
                MessageBox.Show($"設定ファイルの読み込みに失敗しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // 登録処理
        private void RegisterCheck(bool isPrint) {
            try {

                FormCheck();
                if (!DataCheck()) { return; }

                RegisterButton.Enabled = false;

                // 印刷のみにチェックがある場合登録処理をしない
                if (!PrintOnlyCheckBox.Checked) {
                    if (isPrint && !Registration()) {
                        return;
                    }
                }

                if (IsLabelPrint) {
                    if (QuantityCheckBox.Checked) {
                        if (isPrint) { MessageBox.Show("登録完了 続けて印刷します。"); }
                        PrintStart(isPrint);
                    }
                    if (DefectQuantityCheckBox.Checked) {
                        MessageBox.Show("登録完了");
                    }
                }
                else {
                    MessageBox.Show("登録完了");
                }
                Close();

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                RegisterButton.Enabled = true;
            }
        }
        private bool Registration() {
            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();

            using var transaction = con.BeginTransaction();
            try {
                var orderNumber = ProductInfo.OrderNumber;
                var substrateNumber = ProductInfo.ProductNumber;
                var quantity = ProductInfo.Quantity;
                var defectQuantity = string.IsNullOrWhiteSpace(DefectQuantityTextBox.Text) ? 0 : Convert.ToInt32(DefectQuantityTextBox.Text);
                var registrationDate = ProductInfo.RegDate;
                var person = ProductInfo.Person;
                var comment = ProductInfo.Comment;
                var rowId = string.Empty;
                var commandText = string.Empty;

                // 製番が新規かチェック
                if (IsRegistration) {
                    var substrateName = string.Empty;
                    commandText =
                        $"""
                        SELECT
                            StockName,SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                        FROM
                            {Constants.VSubstrateTableName}
                        WHERE
                            StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                        GROUP BY
                            StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                        ORDER BY
                            MIN(ID)
                        LIMIT 1
                        ;
                        """;
                    using var dr = ExecuteReader(con, commandText,
                        ("@StockName", ProductInfo.StockName),
                        ("@SubstrateModel", ProductInfo.SubstrateModel),
                        ("@SubstrateNumber", substrateNumber)
                    );
                    while (dr.Read()) {
                        substrateName = $"{dr["SubstrateName"]}";
                    }

                    if (substrateName != string.Empty) {
                        if (ProductInfo.SubstrateName == substrateName) {
                            if (QuantityCheckBox.Checked && !DefectQuantityCheckBox.Checked) {
                                var result = MessageBox.Show($"[{substrateNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                                if (result == DialogResult.No) { return false; }
                            }
                        }
                    }
                }

                // 不良処理時在庫チェック
                if (DefectQuantityCheckBox.Checked) {
                    commandText =
                        $"""
                        SELECT
                            StockName,SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                        FROM
                            {Constants.VSubstrateTableName}
                        WHERE
                            StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                        GROUP BY
                            StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                        ORDER BY
                            MIN(ID)
                        LIMIT 1
                        ;
                        """;
                    using var dr = ExecuteReader(con, commandText,
                        ("@StockName", ProductInfo.StockName),
                        ("@SubstrateModel", ProductInfo.SubstrateModel),
                        ("@SubstrateNumber", substrateNumber)
                    );
                    if (dr.Read()) {
                        if (Convert.ToInt32(dr["Stock"]) < defectQuantity) {
                            throw new Exception($"[{substrateNumber}]は在庫が[{dr["Stock"]}]です。");
                        }
                    }
                    else {
                        // データが見つからなかった場合の処理
                        throw new Exception($"[{substrateNumber}]は登録がありません。");
                    }
                }

                // 基板登録テーブルへ追加
                commandText =
                    $"""
                    INSERT INTO {Constants.TSubstrateTableName} (
                        SubstrateID,SubstrateNumber,OrderNumber,Increase,Defect,Person,RegDate,Comment
                        )
                    VALUES (
                        @SubstrateID,@SubstrateNumber,@OrderNumber,@Increase,@Defect,@Person,@RegDate,@Comment
                        )
                    ;
                    """;
                ExecuteNonQuery(con, commandText,
                    ("@SubstrateID", ProductInfo.SubstrateID),
                    ("@SubstrateNumber", string.IsNullOrWhiteSpace(substrateNumber) ? DBNull.Value : substrateNumber),
                    ("@OrderNumber", string.IsNullOrWhiteSpace(orderNumber) ? DBNull.Value : orderNumber),
                    ("@Increase", QuantityCheckBox.Checked ? quantity : DBNull.Value),
                    ("@Defect", DefectQuantityCheckBox.Checked ? $"-{defectQuantity}" : DBNull.Value),
                    ("@RegDate", string.IsNullOrWhiteSpace(registrationDate) ? DBNull.Value : registrationDate),
                    ("@Person", string.IsNullOrWhiteSpace(person) ? DBNull.Value : person),
                    ("@Comment", string.IsNullOrWhiteSpace(comment) ? DBNull.Value : comment)
                    );
                commandText = $@"SELECT MAX(ID) FROM {Constants.VSubstrateTableName};";
                rowId = ExecuteScalar(con, commandText).ToString() ?? string.Empty;

                // ログ出力
                var logQuantity = QuantityCheckBox.Checked ? quantity.ToString() : string.Empty;
                var logDefectQuantity = DefectQuantityCheckBox.Checked ? (0 - defectQuantity).ToString() : string.Empty;

                string[] logMessageArray = [
                    $"[基板登録]",
                    $"[{ProductInfo.CategoryName}]",
                    $"ID[{rowId}]",
                    $"注文番号[{orderNumber}]",
                    $"製造番号[{substrateNumber}]",
                    $"製品名[{ProductInfo.ProductName}]",
                    $"基板名[{ProductInfo.SubstrateName}]",
                    $"型式[{ProductInfo.SubstrateModel}]",
                    $"追加数[{logQuantity}]",
                    $"使用数[]",
                    $"減少数[{logDefectQuantity}]",
                    $"[]",
                    $"登録日[{registrationDate}]",
                    $"担当者[{person}]",
                    $"コメント[{comment}]"
                ];
                CommonUtils.Logger.AppendLog(logMessageArray);

                transaction.Commit();

                // バックアップ作成
                CommonUtils.BackupManager.CreateBackup();

                return true;
            } catch (Exception ex) {
                if (transaction.Connection is not null) { // 接続が開いているか確認する。
                    transaction.Rollback();
                }
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

            if (PersonComboBox.SelectedIndex == -1) {
                throw new Exception("担当者が選択されていません。");
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

            var quantity = 0;

            if (QuantityCheckBox.Checked) {
                if (!int.TryParse(QuantityTextBox.Text, out quantity) || quantity <= 0) {
                    MessageBox.Show("数量は1以上の有効な数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    QuantityTextBox.Focus();
                    return false;
                }
            }
            if (DefectQuantityCheckBox.Checked) {
                if (!int.TryParse(DefectQuantityTextBox.Text, out var defectQuantity) || defectQuantity <= 0) {
                    MessageBox.Show("数量は1以上の有効な数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DefectQuantityTextBox.Focus();
                    return false;
                }
            }

            ProductInfo.OrderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
            ProductInfo.ProductNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
            ProductInfo.Quantity = quantity;
            ProductInfo.Person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
            ProductInfo.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateTimePicker.Value.ToShortDateString() : string.Empty;
            ProductInfo.Comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;

            _serialList.Clear();
            var substrateNumber = ManufacturingNumberMaskedTextBox.Text.Substring(5, 5);

            var monthCode = DateTime.Parse(ProductInfo.RegDate).ToString("MM");

            monthCode = monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => monthCode
            };
            var regDate = DateTime.Parse(ProductInfo.RegDate);
            var map = new Dictionary<string, string> {
                ["{T}"] = ProductInfo.Initial,
                ["{Y}"] = regDate.ToString("yy"),
                ["{MM}"] = regDate.ToString("MM"),
                ["{R}"] = ProductInfo.Revision,
                ["{M}"] = monthCode[^1..],
                ["{S}"] = substrateNumber
            };

            var outputCodes = LabelPrintSettings.TextFormat ?? string.Empty;
            foreach (var kv in map) {
                outputCodes = outputCodes.Replace(kv.Key, kv.Value);
            }

            for (var i = 0; i < quantity; i++) {
                _serialList.Add(outputCodes);
            }

            return true;
        }

        private static void ExecuteNonQuery(SqliteConnection con, string commandText, params (string, object?)[] parameters) {
            var command = con.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                // 空の文字列の場合にNULLを設定
                var sqlValue = value is string strValue && string.IsNullOrEmpty(strValue) ? DBNull.Value : value ?? DBNull.Value;
                command.Parameters.Add(name, SqliteType.Text).Value = sqlValue;
            }

            command.ExecuteNonQuery();
        }
        private static object ExecuteScalar(SqliteConnection con, string commandText, params (string, object)[] parameters) {
            var command = con.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                command.Parameters.Add(name, SqliteType.Text).Value = value ?? DBNull.Value;
            }

            return command.ExecuteScalar() ?? 0;
        }
        private static SqliteDataReader ExecuteReader(SqliteConnection con, string commandText, params (string, object)[] parameters) {
            var command = con.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                command.Parameters.Add(name, SqliteType.Text).Value = value ?? DBNull.Value;
            }

            return command.ExecuteReader();
        }
        // 印刷処理
        private void PrintStart(bool isPrint) {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();

                var isPreview = !isPrint;
                var startLine = (int)PrintPositionNumericUpDown.Value - 1;
                var serialType = "Substrate";

                pd.BeginPrint += (sender, e) => {
                    PrintManager.Initialize(ProductInfo, SubstratePrintSettings, _serialList);
                };
                pd.PrintPage += (sender, e) => {
                    var hasMore = PrintManager.PrintSerialCommon(e, isPreview, startLine, serialType);
                    e.HasMorePages = hasMore;
                };

                ProductInfo.Quantity = int.TryParse(QuantityTextBox.Text, out var quantity) ? quantity : 0;
                if (ProductInfo.Quantity == 0) {
                    throw new Exception("数量が入力されていません。");
                }

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
                        break;
                    case false:
                        ProductInfo.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateTimePicker.Value.ToShortDateString() : string.Empty;
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
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 在庫数取得
        private string GetStockQuantity() {

            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();

            using var transaction = con.BeginTransaction();

            var commandText =
                $"""
                SELECT
                    StockName,SubstrateName,SubstrateModel,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                FROM
                    {Constants.VSubstrateTableName}
                WHERE
                    StockName = @StockName AND SubstrateModel = @SubstrateModel
                GROUP BY
                    StockName, SubstrateName, SubstrateModel
                ;
                """;
            using var dr = ExecuteReader(con, commandText,
                ("@StockName", ProductInfo.StockName),
                ("@SubstrateModel", ProductInfo.SubstrateModel)
            );

            transaction.Commit();

            if (dr.Read()) {
                return dr["Stock"].ToString() ?? "0";
            }
            else {
                return "0";
            }
        }

        // コメント用テンプレート
        private void TemplateComment() {
            var templateWord = CommentComboBox.SelectedIndex switch {
                0 => "[Rev.UP]変更点番号:",
                _ => string.Empty
            };
            CommentTextBox.Text = $"{CommentTextBox.Text}{templateWord}";
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
                    if (checkBox.Checked) { DefectQuantityCheckBox.Checked = false; }
                    break;
                case "DefectQuantityCheckBox":
                    DefectQuantityTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) { QuantityCheckBox.Checked = false; }
                    break;
                case "ExtraCheckBox7":
                    ExtraTextBox7.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox1":
                    ExtraTextBox1.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox2":
                    ExtraTextBox2.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox3":
                    ExtraTextBox3.Enabled = checkBox.Checked;
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
                case "PrintOnlyCheckBox":
                    PrintButton.Enabled = checkBox.Checked;
                    RegisterButton.Enabled = !checkBox.Checked;
                    if (PrintOnlyCheckBox.Checked) { MessageBox.Show($"登録せずに印刷だけ行います。", "注意", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                    break;
                default:
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
                throw new Exception($"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー{Environment.NewLine}{ex.Message}");
            }
        }
        // 取得情報表示
        private void ShowInfo() {
            var items = new Dictionary<string, string>
                {
                    {"Proness1", $"{ProductInfo.Proness1}"},
                    {"Proness2", $"{ProductInfo.Proness2}"},
                    {"Proness3", $"{ProductInfo.Proness3}"},
                    {"Proness4", $"{ProductInfo.Proness4}"},
                    {"Proness5", $"{ProductInfo.Proness5}"},
                    {"ProductName", $"{ProductInfo.ProductName}"},
                    {"StockName", $"{ProductInfo.StockName}"},
                    {"SubstrateName", $"{ProductInfo.SubstrateName}"},
                    {"SubstrateModel", $"{ProductInfo.SubstrateModel}"},
                    {"ProductNumber", $"{ProductInfo.ProductNumber}"},
                    {"Initial", $"{ProductInfo.Initial}"},
                    {"RegType", $"{ProductInfo.RegType}"},
                    {"SerialPrintType", $"{ProductInfo.SerialPrintType}"},
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
                Font = new Font("PlemolJP", ProductInfo.FontSize), // 等幅フォント
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
        // 基板設定を開く
        private void OpenSubstrateInformation() {
            try {
                ExcelServiceClosedXml.SubstrateInformationClosedXml.OpenSubstrateInformationClosedXml(ProductInfo);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SubstrateRegistrationWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void QrCodeButton_Click(object sender, EventArgs e) { QrInput(); }
        private void RegisterButton_Click(object sender, EventArgs e) { RegisterCheck(true); }
        private void PrintButton_Click(object sender, EventArgs e) { RegisterCheck(true); }
        private void OpenSubstrateInformationButton_Click(object sender, EventArgs e) { OpenSubstrateInformation(); }
        private void TemplateButton_Click(object sender, EventArgs e) { TemplateComment(); }
        private void NumberCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBoxChecked(sender, e); }
        private void QuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void DefectQuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void 印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) { RegisterCheck(false); }
        private void 印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            PrintSettingsWindow ls = new() {
                ProductInfo = ProductInfo,
                serialType = "Substrate"
            };
            ls.ShowDialog(this);
            LoadSettings(printSettingPath);
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) { ShowInfo(); }
        private void QrCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }
    }
}
