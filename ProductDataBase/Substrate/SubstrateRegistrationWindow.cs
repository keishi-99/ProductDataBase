using ProductDatabase.Other;
using ProductDatabase.Product;
using ProductDatabase.Substrate;
using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SubstrateRegistrationWindow : Form {

        public ProductInformation ProductInfo { get; }

        public SubstratePrintSettings SubstratePrintSettings { get; set; } = new SubstratePrintSettings();
        public Substrate.LabelPageSettings LabelPageSettings => SubstratePrintSettings.LabelPageSettings ?? new Substrate.LabelPageSettings();
        public Substrate.LabelLayoutSettings LabelLayoutSettings => SubstratePrintSettings.LabelLayoutSettings ?? new Substrate.LabelLayoutSettings();
        private readonly string _printSettingPath = SubstratePrintSettingsWindow.s_substratePrintSettingFilePath;

        private string _labelSubNSerial = string.Empty;

        private int _labelSubNumLabelsToPrint;

        private const float MmPerInch = 25.4f;

        private int _pageCount;
        private System.Drawing.Printing.PrintAction _printAction;

        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "DefectNumberCheckBox",
                    "RevisionCheckBox", "ExtraCheckBox1", "ExtraCheckBox2", "ExtraCheckBox3", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" ];

        // プロパティ設定
        private bool IsLabelPrint => ProductInfo.PrintType is 1;
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
                    ExtraCheckBox7.Visible = false;
                    PrintRowLabel.Visible = false;
                    PrintPostionNumericUpDown.Visible = false;
                    PrintOnlyCheckBox.Visible = false;
                    PrintButton.Visible = false;
                }

                if (File.Exists(_printSettingPath) == false) { throw new Exception("印刷設定ファイルが見つかりませんでした"); }
                SubstratePrintSettings = new SubstratePrintSettings();
                LoadSettings(_printSettingPath);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            } finally {
            }
        }
        private void LoadSettings(string settingFilePath) {
            try {
                var jsonString = File.ReadAllText(settingFilePath);
                SubstratePrintSettings = System.Text.Json.JsonSerializer.Deserialize<SubstratePrintSettings>(jsonString) ?? new SubstratePrintSettings();
            } catch (Exception ex) {
                MessageBox.Show($"設定ファイルの読み込みに失敗しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // 登録処理
        private void RegisterCheck() {
            try {
                // 印刷のみにチェックがある場合処理をしない
                if (PrintOnlyCheckBox.Checked) { return; }

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
                if (!anyTextBoxEnabled) { throw new Exception("何も入力されていません"); }
                if (!allTextBoxesFilled || (PersonCheckBox.Checked && string.IsNullOrWhiteSpace(PersonComboBox.Text))) { throw new Exception("空欄があります。"); }

                if (ManufacturingNumberCheckBox.Checked && ManufacturingNumberMaskedTextBox.Text.Length != 15) { throw new Exception("製番を10桁+4桁で入力して下さい。"); }

                if (QuantityCheckBox.Checked) {
                    if (string.IsNullOrWhiteSpace(QuantityTextBox.Text)) {
                        MessageBox.Show("数量を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        QuantityTextBox.Focus();
                        return;
                    }
                    if (!int.TryParse(QuantityTextBox.Text, out var quantity)) {
                        MessageBox.Show("数量は有効な数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        QuantityTextBox.Focus();
                        return;
                    }
                    if (quantity <= 0) {
                        MessageBox.Show("1台以上入力して下さい。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        QuantityTextBox.Focus();
                        return;
                    }
                }

                if (DefectNumberCheckBox.Checked) {
                    if (string.IsNullOrWhiteSpace(DefectNumberTextBox.Text)) {
                        MessageBox.Show("数量を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        DefectNumberTextBox.Focus();
                        return;
                    }
                    if (!int.TryParse(DefectNumberTextBox.Text, out var quantity)) {
                        MessageBox.Show("数量は有効な数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        DefectNumberTextBox.Focus();
                        return;
                    }
                    if (quantity <= 0) {
                        MessageBox.Show("1台以上入力して下さい。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        DefectNumberTextBox.Focus();
                        return;
                    }
                }

                if (!QuantityCheckBox.Checked && !DefectNumberCheckBox.Checked) { throw new Exception("追加量か減少数を入力してください。"); }

                if (!DefectNumberCheckBox.Checked && string.IsNullOrEmpty(PrintPostionNumericUpDown.Text)) { PrintPostionNumericUpDown.Text = "1"; }

                var result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) {
                    return;
                }

                RegisterButton.Enabled = false;

                if (!Registration()) { return; }

                ProductInfo.Person = PersonComboBox.Text;

                if (IsLabelPrint) {
                    if (QuantityCheckBox.Checked) {
                        MessageBox.Show("登録完了 続けて印刷します。");
                        PrintStart();
                        MessageBox.Show("印刷完了");
                        Close();
                    }
                    if (DefectNumberCheckBox.Checked) {
                        MessageBox.Show("登録完了");
                        Close();
                    }
                }
                else {
                    MessageBox.Show("登録完了");
                    Close();
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                RegisterButton.Enabled = true;
            }
        }
        private bool Registration() {
            using SQLiteConnection connection = new(GetConnectionRegistration());
            connection.Open();
            using var transaction = connection.BeginTransaction();
            try {
                var substrateTableName = $"[{ProductInfo.CategoryName}_Substrate]";
                var substrateNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
                var orderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
                var quantity = string.IsNullOrWhiteSpace(QuantityTextBox.Text) ? 0 : Convert.ToInt32(QuantityTextBox.Text);
                var defectNumber = string.IsNullOrWhiteSpace(DefectNumberTextBox.Text) ? 0 : Convert.ToInt32(DefectNumberTextBox.Text);
                var registrationDate = RegistrationDateCheckBox.Checked ? RegistrationDateTimePicker.Value.ToShortDateString() : string.Empty;
                var person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
                var comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;
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
                            {substrateTableName}
                        WHERE
                            StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                        GROUP BY
                            StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                        ORDER BY
                            MIN(ID)
                        LIMIT 1
                        ;
                        """;
                    using var dr = ExecuteReader(connection, commandText,
                        ("@StockName", ProductInfo.StockName),
                        ("@SubstrateModel", ProductInfo.SubstrateModel),
                        ("@SubstrateNumber", substrateNumber)
                    );
                    while (dr.Read()) {
                        substrateName = $"{dr["SubstrateName"]}";
                    }

                    if (substrateName != string.Empty) {
                        if (ProductInfo.SubstrateName == substrateName) {
                            if (QuantityCheckBox.Checked && !DefectNumberCheckBox.Checked) {
                                var result = MessageBox.Show($"[{substrateNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                                if (result == DialogResult.No) { return false; }
                            }
                        }
                    }
                }

                // 不良処理時在庫チェック
                if (DefectNumberCheckBox.Checked) {
                    commandText =
                        $"""
                        SELECT
                            StockName,SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                        FROM
                            {substrateTableName}
                        WHERE
                            StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                        GROUP BY
                            StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                        ORDER BY
                            MIN(ID)
                        LIMIT 1
                        ;
                        """;
                    using var dr = ExecuteReader(connection, commandText,
                        ("@StockName", ProductInfo.StockName),
                        ("@SubstrateModel", ProductInfo.SubstrateModel),
                        ("@SubstrateNumber", substrateNumber)
                    );
                    if (dr.Read()) {
                        if (Convert.ToInt32(dr["Stock"]) < defectNumber) {
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
                    INSERT INTO {substrateTableName} (
                        StockName,SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,Increase,Defect,Person,RegDate,Comment
                        )
                    VALUES (
                        @StockName,@SubstrateName,@SubstrateModel,@SubstrateNumber,@OrderNumber,@Increase,@Defect,@Person,@RegDate,@Comment
                        )
                    ;
                    """;
                ExecuteNonQuery(connection, commandText,
                    ("@StockName", ProductInfo.StockName),
                    ("@SubstrateName", ProductInfo.SubstrateName),
                    ("@SubstrateModel", ProductInfo.SubstrateModel),
                    ("@SubstrateNumber", string.IsNullOrWhiteSpace(substrateNumber) ? DBNull.Value : substrateNumber),
                    ("@OrderNumber", string.IsNullOrWhiteSpace(orderNumber) ? DBNull.Value : orderNumber),
                    ("@Increase", QuantityCheckBox.Checked ? quantity : DBNull.Value),
                    ("@Defect", DefectNumberCheckBox.Checked ? $"-{defectNumber}" : DBNull.Value),
                    ("@RegDate", string.IsNullOrWhiteSpace(registrationDate) ? DBNull.Value : registrationDate),
                    ("@Person", string.IsNullOrWhiteSpace(person) ? DBNull.Value : person),
                    ("@Comment", string.IsNullOrWhiteSpace(comment) ? DBNull.Value : comment)
                    );
                commandText = $@"SELECT MAX(ID) FROM {substrateTableName};";
                rowId = ExecuteScalar(connection, commandText).ToString() ?? string.Empty;

                // ログ出力
                var number = QuantityCheckBox.Checked ? quantity : 0 - defectNumber;

                string[] logMessageArray = [
                    $"[基板登録]",
                    $"[{ProductInfo.CategoryName}]",
                    $"ID[{rowId}]",
                    $"注文番号[{orderNumber}]",
                    $"製造番号[{substrateNumber}]",
                    $"製品名[{ProductInfo.ProductName}]",
                    $"基板名[{ProductInfo.SubstrateName}]",
                    $"型式[{ProductInfo.SubstrateModel}]",
                    $"数量[{number}]",
                    $"[]",
                    $"[]",
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
                if (transaction.Connection != null) { //接続が開いているか確認する。
                    transaction.Rollback();
                }
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private static void ExecuteNonQuery(SQLiteConnection con, string commandText, params (string, object?)[] parameters) {
            using var command = con.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                // 空の文字列の場合にNULLを設定
                var sqlValue = value is string strValue && string.IsNullOrEmpty(strValue) ? DBNull.Value : value ?? DBNull.Value;
                command.Parameters.Add(name, DbType.String).Value = sqlValue;
            }

            command.ExecuteNonQuery();
        }
        private static object ExecuteScalar(SQLiteConnection con, string commandText, params (string, object)[] parameters) {
            using var command = con.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                command.Parameters.Add(name, DbType.String).Value = value ?? DBNull.Value;
            }

            return command.ExecuteScalar() ?? 0;
        }
        private static SQLiteDataReader ExecuteReader(SQLiteConnection con, string commandText, params (string, object)[] parameters) {
            using var command = con.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                command.Parameters.Add(name, DbType.String).Value = value ?? DBNull.Value;
            }

            return command.ExecuteReader();
        }
        // 印刷処理
        private void PrintStart() {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();
                pd.BeginPrint += (sender, e) => _printAction = e.PrintAction;
                pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

                _labelSubNumLabelsToPrint = int.TryParse(QuantityTextBox.Text, out var quantity) ? quantity : 0;
                if (_labelSubNumLabelsToPrint == 0) {
                    throw new Exception("数量が入力されていません。");
                }
                _pageCount = 1;

                SubstrateRegistrationPrintDialog.Document = pd;

                if (SubstrateRegistrationPrintDialog.ShowDialog() == DialogResult.OK) {
                    // ローディング画面の表示
                    using var loadingForm = new LoadingForm();
                    // 別スレッドで印刷処理を実行
                    Task.Run(() => {
                        try {
                            SubstrateRegistrationPrintDialog.Document.Print();
                        } finally {
                            // 印刷が終了したらローディング画面を閉じる
                            loadingForm.Invoke(new Action(() => loadingForm.Close()));
                        }
                    });

                    // ローディング画面をモーダルとして表示
                    loadingForm.ShowDialog();
                }
                else {
                    return;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool PreviewPrint() {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();
                pd.BeginPrint += (sender, e) => _printAction = e.PrintAction;
                pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

                _labelSubNumLabelsToPrint = int.TryParse(QuantityTextBox.Text, out var quantity) ? quantity : 0;
                if (_labelSubNumLabelsToPrint == 0) {
                    throw new Exception("数量が入力されていません。");
                }

                _pageCount = 1;

                // 最大で表示
                SubstrateRegistrationPrintPreviewDialog.Shown += (sender, e) => {
                    if (sender is Form form) {
                        form.WindowState = FormWindowState.Maximized;
                    }
                };
                SubstrateRegistrationPrintPreviewDialog.PrintPreviewControl.Zoom = 3;
                SubstrateRegistrationPrintPreviewDialog.Document = pd;
                SubstrateRegistrationPrintPreviewDialog.ShowDialog();

                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }
                var dpiX = e.Graphics.DpiX;
                var dpiY = e.Graphics.DpiY;
                e.Graphics.PageUnit = GraphicsUnit.Millimeter;

                // プレビューかどうかの判定
                var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;

                var labelWidth = (float)LabelPageSettings.LabelWidth;
                var labelHeight = (float)LabelPageSettings.LabelHeight;
                var marginX = LabelPageSettings.MarginX;
                var marginY = LabelPageSettings.MarginY;
                var intervalX = LabelPageSettings.IntervalX;
                var intervalY = LabelPageSettings.IntervalY;
                var headerPositionX = LabelPageSettings.HeaderPositionX;
                var headerPositionY = LabelPageSettings.HeaderPositionY;
                var headerFont = LabelPageSettings.HeaderFont;
                var startLine = (int)PrintPostionNumericUpDown.Value - 1;

                // ハードマージンをミリメートルに変換
                var hardMarginX = isPreview ? 0 : e.PageSettings.HardMarginX * MmPerInch / 100;
                var hardMarginY = isPreview ? 0 : e.PageSettings.HardMarginY * MmPerInch / 100;

                var headerString = ConvertHeaderString(LabelPageSettings.HeaderTextFormat);

                _labelSubNSerial = ManufacturingNumberMaskedTextBox.Text;

                if (_pageCount >= 2) {
                    startLine = 0;
                }

                // 最初のページのみオフセットを調整
                var verticalOffset = _pageCount == 1 ? startLine * (intervalY + labelHeight) : 0;
                // ヘッダーの描画
                e.Graphics.DrawString(headerString, headerFont, Brushes.Gray, (float)headerPositionX, (float)(verticalOffset + headerPositionY - hardMarginY));

                var labelCountX = LabelPageSettings.LabelsPerColumn;
                var labelCountY = LabelPageSettings.LabelsPerRow;
                int y;
                var copiesPerLabel = LabelLayoutSettings.CopiesPerLabel;
                if (labelCountX == 0 || labelCountY == 0 || copiesPerLabel == 0) { throw new Exception("印刷設定が異常です。"); }
                for (y = startLine; y < labelCountY; y++) {
                    int x;
                    for (x = 0; x < labelCountX; x++) {
                        var posX = (float)(marginX - hardMarginX + (x * (intervalX + labelWidth)));
                        var posY = (float)(marginY - hardMarginY + (y * (intervalY + labelHeight)));

                        var generatedCode = GenerateCode(_labelSubNSerial);
                        var labelImage = MakeLabelImage(generatedCode, dpiX, dpiY);
                        e.Graphics.DrawImage(labelImage, posX, posY, labelWidth, labelHeight);

                        _labelSubNumLabelsToPrint--;

                        if (_labelSubNumLabelsToPrint <= 0) {
                            copiesPerLabel--;
                            if (copiesPerLabel <= 0) {
                                // 最終行の行番号を描画
                                var sf = new StringFormat {
                                    Alignment = StringAlignment.Near,
                                    LineAlignment = StringAlignment.Center
                                };
                                var layoutRect = new RectangleF(0, posY, 0, labelHeight);
                                var rowNumber = (y + 1).ToString();
                                e.Graphics.DrawString(rowNumber, headerFont, Brushes.Black, layoutRect, sf);
                                // 次のページがあるかどうかの判定
                                e.HasMorePages = false;
                                _pageCount = 0;
                                _labelSubNumLabelsToPrint = 0;
                                return;
                            }
                            else {
                                _labelSubNumLabelsToPrint += x + 1;
                                break;
                            }
                        }

                        if (x >= labelCountX - 1) {
                            copiesPerLabel--;
                            if (copiesPerLabel <= 0) {
                                copiesPerLabel = LabelLayoutSettings.CopiesPerLabel;
                            }
                            else if (copiesPerLabel > 0) {
                                _labelSubNumLabelsToPrint += x + 1;
                                break;
                            }
                        }
                    }
                }

                if (_labelSubNumLabelsToPrint > 0) {
                    _pageCount++;
                    e.HasMorePages = true;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        private string ConvertHeaderString(string s) {
            s = s.Replace("%P", ProductInfo.SubstrateName)
                 .Replace("%T", ProductInfo.SubstrateModel)
                 .Replace("%D", DateTime.Today.ToShortDateString())
                 .Replace("%M", ManufacturingNumberMaskedTextBox.Text)
                 .Replace("%O", OrderNumberTextBox.Text)
                 .Replace("%N", QuantityTextBox.Text)
                 .Replace("%U", ProductInfo.Person);
            return s;
        }
        private string GenerateCode(string serial) {
            var monthCode = RegistrationDateTimePicker.Value.Month.ToString("MM");

            monthCode = monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => monthCode
            };

            var outputCode = LabelLayoutSettings.TextFormat;
            var serialCode = serial.Substring(5, 5);
            outputCode = outputCode.Replace("%T", ProductInfo.Initial)
                                    .Replace("%Y", RegistrationDateTimePicker.Value.Year.ToString("yy"))
                                    .Replace("%MM", RegistrationDateTimePicker.Value.Month.ToString("MM"))
                                    .Replace("%M", string.IsNullOrEmpty(monthCode) ? string.Empty : monthCode[^1..])
                                    .Replace("%S", serialCode);

            return outputCode;
        }
        private Bitmap MakeLabelImage(string text, float dpiX, float dpiY) {

            var labelWidth = (float)LabelPageSettings.LabelWidth;
            var labelHeight = (float)LabelPageSettings.LabelHeight;

            // ビットマップのサイズをピクセル単位で計算
            var pixelWidth = (int)(labelWidth / MmPerInch * dpiX);
            var pixelHeight = (int)(labelHeight / MmPerInch * dpiY);

            var labelImage = new Bitmap(pixelWidth, pixelHeight);
            labelImage.SetResolution(dpiX, dpiY);

            //Bitmap labelImage = new(sizeX, sizeY);
            using (var g = Graphics.FromImage(labelImage)) {
                g.PageUnit = GraphicsUnit.Millimeter;
                // アンチエイリアス処理を改善
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

                var fontName = LabelLayoutSettings.TextFont.Name;
                var fontSize = LabelLayoutSettings.TextFont.SizeInPoints;
                using Font fnt = new(fontName, fontSize);

                // StringFormat を使用して中心に配置
                var sf = new StringFormat {
                    Alignment = LabelLayoutSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near,
                    LineAlignment = LabelLayoutSettings.AlignTextCenterY ? StringAlignment.Center : StringAlignment.Near
                };

                var stringPosX = LabelLayoutSettings.AlignTextCenterX ? 0f : (float)LabelLayoutSettings.TextPositionX;
                var stringPosY = LabelLayoutSettings.AlignTextCenterY ? 0f : (float)LabelLayoutSettings.TextPositionY;

                var pageWidth = (float)LabelPageSettings.LabelWidth;
                var pageHeight = (float)LabelPageSettings.LabelHeight;

                // 矩形領域を計算 (文字列を配置する領域)
                var layoutRect = new RectangleF(stringPosX, stringPosY, pageWidth - stringPosX, pageHeight - stringPosY);
                g.DrawString(text, fnt, Brushes.Black, layoutRect, sf);

                // プレビューかどうかの判定
                var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;
                // プレビュー時、黒枠を描画
                if (isPreview) {
                    // 0.1mmの黒いペンで枠線を描画
                    using var p = new Pen(Color.Black, 0.1f);
                    g.DrawRectangle(p, 0, 0, labelWidth - 0.1f, labelHeight - 0.1f);
                }
            }

            return labelImage;
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
                    if (checkBox.Checked) { DefectNumberCheckBox.Checked = false; }
                    break;
                case "DefectNumberCheckBox":
                    DefectNumberTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) { QuantityCheckBox.Checked = false; }
                    break;
                case "RevisionCheckBox":
                    //RevisionTextBox.Enabled = checkBox.Checked;
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
                    {"PrintType", $"{ProductInfo.PrintType}"}
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
                //var item = new ListViewItem(kvp.Key);
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

        private void SubstrateRegistrationWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void QrCodeButton_Click(object sender, EventArgs e) { QrInput(); }
        private void RegisterButton_Click(object sender, EventArgs e) { RegisterCheck(); }
        private void PrintButton_Click(object sender, EventArgs e) { PrintStart(); }
        private void SubstrateRegistrationPrintDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) { PrintDocumentPrintPage(sender, e); }
        private void TemplateButton_Click(object sender, EventArgs e) { TemplateComment(); }
        private void NumberCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBoxChecked(sender, e); }
        private void QuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void DefectNumberTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void 印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) { PreviewPrint(); }
        private void 印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            SubstratePrintSettingsWindow ls = new();
            ls.ShowDialog(this);
            LoadSettings(_printSettingPath);
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) { ShowInfo(); }
        private void SubstrateRegistrationPrintPreviewDialog_Load(object sender, EventArgs e) {
            var tool = (ToolStrip)SubstrateRegistrationPrintPreviewDialog.Controls[1];
            tool.Items[0].Visible = false;
        }
        private void QrCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }
    }
}
