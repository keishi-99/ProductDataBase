using ProductDatabase.Other;
using ProductDatabase.Substrate;
using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SubstrateRegistrationWindow : Form {

        public ProductInformation ProductInfo { get; }

        public SubstratePrintSettings SubstratePrintSettings { get; set; } = new SubstratePrintSettings();
        private readonly string _printSettingPath = SubstratePrintSettingsWindow.s_substratePrintSettingFilePath;

        private string _labelSubNSerial = string.Empty;

        private int _labelSubNumLabelsToPrint;

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

                // TextBoxへ今日の年月日を入力
                var dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = dtNow.ToShortDateString();

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. ProductInfo.PersonList]);

                // 在庫管理する基板はDB2へ接続し対象製品の在庫取得
                //if (IsRegistration) {
                //    using SQLiteConnection con = new(GetConnectionRegistration());
                //    con.Open();
                //    using var cmd = con.CreateCommand();

                //    // テーブル検索SQL - [[ProductName]_Substrate]テーブルの最新の[Revison]を取得
                //    cmd.CommandText = $"""SELECT Revision FROM "{ProductInfo.CategoryName}_Substrate" WHERE SubstrateModel = @SubstrateModel AND Revision IS NOT NULL ORDER BY ID DESC""";
                //    cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = ProductInfo.SubstrateModel;
                //    var result = cmd.ExecuteScalar();
                //    RevisionTextBox.Text = result?.ToString() ?? "";
                //}

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

                if (QuantityCheckBox.Checked && int.Parse(QuantityTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

                if (DefectNumberCheckBox.Checked && int.Parse(DefectNumberTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

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
                var substrateNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
                var orderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
                var quantity = string.IsNullOrWhiteSpace(QuantityTextBox.Text) ? 0 : Convert.ToInt32(QuantityTextBox.Text);
                var defectNumber = string.IsNullOrWhiteSpace(DefectNumberTextBox.Text) ? 0 : Convert.ToInt32(DefectNumberTextBox.Text);
                var registrationDate = RegistrationDateCheckBox.Checked ? RegistrationDateMaskedTextBox.Text : string.Empty;
                var person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
                var comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;
                var rowId = string.Empty;
                var commandText = string.Empty;

                // 製番が新規かチェック
                if (IsRegistration) {
                    var substrateName = string.Empty;
                    commandText = $@"
                            SELECT
                                StockName,SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                            FROM
                                {ProductInfo.CategoryName}_Substrate
                            WHERE
                                StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                            GROUP BY
                                StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                            ORDER BY
                                MIN(ID)
                            LIMIT 1;
                            ";
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
                        $@"
                        SELECT
                            StockName,SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                        FROM
                            {ProductInfo.CategoryName}_Substrate
                        WHERE
                            StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                        GROUP BY
                            StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                        ORDER BY
                            MIN(ID)
                        LIMIT 1;
                        ";
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
                    $@"
                    INSERT INTO ""{ProductInfo.CategoryName}_Substrate""
                        (StockName,SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,Increase,Defect,Person,RegDate,Comment)
                    VALUES
                        (@StockName,@SubstrateName,@SubstrateModel,@SubstrateNumber,@OrderNumber,@Increase,@Defect,@Person,@RegDate,@Comment);";
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
                commandText = $@"SELECT MAX(ID) FROM ""{ProductInfo.CategoryName}_Substrate"";";
                rowId = ExecuteScalar(connection, commandText).ToString() ?? string.Empty;

                // バックアップ作成
                CommonUtils.BackupManager.CreateBackup();

                // ログ出力
                var number = QuantityCheckBox.Checked ? quantity : 0 - defectNumber;

                string[] logMessageArray = [
                    $"[基板登録]",
                    $"ID[{ProductInfo.CategoryName}_{rowId}]",
                    $"注文番号[{orderNumber}]",
                    $"製造番号[{substrateNumber}]",
                    $"製品名[{ProductInfo.ProductName}]",
                    $"基板名[{ProductInfo.SubstrateName}]",
                    $"型式[{ProductInfo.SubstrateModel}]",
                    $"数量[{number}]",
                    $"登録日[{registrationDate}]",
                    $"担当者[{person}]",
                    $"コメント[{comment}]"
                ];
                CommonUtils.Logger.AppendLog(logMessageArray);

                transaction.Commit();

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
                if (SubstratePrintSettings == null || e.Graphics == null) { return; }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                // プレビューかどうかの判定
                var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;

                var labelWidth = (float)SubstratePrintSettings.LabelPageSettings.LabelWidth;
                var labelHeight = (float)SubstratePrintSettings.LabelPageSettings.LabelHeight;
                var marginX = SubstratePrintSettings.LabelPageSettings.MarginX;
                var marginY = SubstratePrintSettings.LabelPageSettings.MarginY;
                var intervalX = SubstratePrintSettings.LabelPageSettings.IntervalX;
                var intervalY = SubstratePrintSettings.LabelPageSettings.IntervalY;
                var headerPositionX = SubstratePrintSettings.LabelPageSettings.HeaderPositionX;
                var headerPositionY = SubstratePrintSettings.LabelPageSettings.HeaderPositionY;
                var headerFont = SubstratePrintSettings.LabelPageSettings.HeaderFont;
                var startLine = (int)PrintPostionNumericUpDown.Value - 1;

                // ハードマージンをミリメートルに変換
                const double MM_PER_HUNDREDTH_INCH = 0.254;
                marginX -= e.PageSettings.HardMarginX * MM_PER_HUNDREDTH_INCH;
                marginY -= e.PageSettings.HardMarginY * MM_PER_HUNDREDTH_INCH;

                var headerString = ConvertHeaderString(SubstratePrintSettings.LabelPageSettings.HeaderTextFormat);

                // 最初のページのみオフセットを調整
                var verticalOffset = _pageCount == 1 ? startLine * (intervalY + labelHeight) : 0;
                // ヘッダーの描画
                e.Graphics.DrawString(headerString, headerFont, Brushes.DarkGray, (float)headerPositionX, (float)(verticalOffset + headerPositionY));

                _labelSubNSerial = ManufacturingNumberMaskedTextBox.Text;

                if (_pageCount >= 2) {
                    startLine = 0;
                }

                var labelCountX = SubstratePrintSettings.LabelPageSettings.LabelsPerColumn;
                var labelCountY = SubstratePrintSettings.LabelPageSettings.LabelsPerRow;
                int y;
                var copiesPerLabel = SubstratePrintSettings.LabelLayoutSettings.CopiesPerLabel;
                if (labelCountX == 0 || labelCountY == 0 || copiesPerLabel == 0) { throw new Exception("印刷設定が異常です。"); }
                for (y = startLine; y < labelCountY; y++) {
                    int x;
                    for (x = 0; x < labelCountX; x++) {
                        var posX = (float)(marginX + (x * (intervalX + labelWidth)));
                        var posY = (float)(marginY + (y * (intervalY + labelHeight)));

                        var generatedCode = GenerateCode(_labelSubNSerial);
                        var labelImage = MakeLabelImage(generatedCode, (int)e.Graphics.DpiX, 1);
                        e.Graphics.DrawImage(labelImage, posX, posY, labelWidth, labelHeight);

                        _labelSubNumLabelsToPrint--;

                        if (_labelSubNumLabelsToPrint <= 0) {
                            copiesPerLabel--;
                            if (copiesPerLabel <= 0) {
                                // 最終行の行番号を表示
                                var rowNumber = (y + 2).ToString();
                                e.Graphics.DrawString(rowNumber, SubstratePrintSettings.LabelPageSettings.HeaderFont, Brushes.Black, 0, posY);
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
                                copiesPerLabel = SubstratePrintSettings.LabelLayoutSettings.CopiesPerLabel;
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
            var monthCode = DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("MM");

            monthCode = monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => monthCode
            };

            var outputCode = SubstratePrintSettings.LabelLayoutSettings.TextFormat;
            var serialCode = serial.Substring(5, 5);
            outputCode = outputCode.Replace("%T", ProductInfo.Initial)
                                    .Replace("%Y", DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("MM"))
                                    .Replace("%M", string.IsNullOrEmpty(monthCode) ? string.Empty : monthCode[^1..])
                                    .Replace("%S", serialCode);

            return outputCode;
        }
        private Bitmap MakeLabelImage(string text, int resolution, int magnitude) {
            if (SubstratePrintSettings is null) { throw new Exception(); }
            var sizeX = (decimal)SubstratePrintSettings.LabelPageSettings.LabelWidth / 25.4M * resolution * magnitude;
            var sizeY = (decimal)SubstratePrintSettings.LabelPageSettings.LabelHeight / 25.4M * resolution * magnitude;

            Bitmap labelImage = new((int)sizeX, (int)sizeY);
            using (var g = Graphics.FromImage(labelImage)) {
                // アンチエイリアス処理を改善
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                var fontSize = (decimal)SubstratePrintSettings.LabelLayoutSettings.TextFont.SizeInPoints / 72.0M * resolution * magnitude;
                using (Font fnt = new(SubstratePrintSettings.LabelLayoutSettings.TextFont.Name, (float)fontSize)) {

                    // StringFormat を使用して中心に配置
                    var sf = new StringFormat {
                        Alignment = SubstratePrintSettings.LabelLayoutSettings.AlignTextXCenter ? StringAlignment.Center : StringAlignment.Near,
                        LineAlignment = SubstratePrintSettings.LabelLayoutSettings.AlignTextYCenter ? StringAlignment.Center : StringAlignment.Near
                    };

                    var stringPosX = SubstratePrintSettings.LabelLayoutSettings.AlignTextXCenter ? 0 : (float)(SubstratePrintSettings.LabelLayoutSettings.TextPositionX / 25.4F * resolution * magnitude);
                    var stringPosY = SubstratePrintSettings.LabelLayoutSettings.AlignTextYCenter ? 0 : (float)(SubstratePrintSettings.LabelLayoutSettings.TextPositionY / 25.4F * resolution * magnitude);

                    // 矩形領域を計算 (文字列を配置する領域)
                    var layoutRect = new RectangleF(stringPosX, stringPosY, labelImage.Width - stringPosX, labelImage.Height - stringPosY);
                    g.DrawString(text, fnt, Brushes.Black, layoutRect, sf);

                    // プレビューかどうかの判定
                    var isPreview = _printAction == System.Drawing.Printing.PrintAction.PrintToPreview;
                    // プレビュー時、黒枠を描画
                    if (isPreview) {
                        using var p = new Pen(Color.Black, 3);
                        g.DrawRectangle(p, 0, 0, labelImage.Width - 1, labelImage.Height - 1);
                    }
                }

                g.Dispose();
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
                throw new Exception($"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー{Environment.NewLine}{ex.Message}");
            }
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
        private void RegistrationDateMaskedTextBox_TypeValidationCompleted(object sender, TypeValidationEventArgs e) { RegistrationDateCheck(sender, e); }
        private void 印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) { PreviewPrint(); }
        private void 印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            SubstratePrintSettingsWindow ls = new();
            ls.ShowDialog(this);
            LoadSettings(_printSettingPath);
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            var entries = new[]
                {
                    ("StrProness1", $"{ProductInfo.Proness1}"),
                    ("StrProness2", $"{ProductInfo.Proness2}"),
                    ("StrProness3", $"{ProductInfo.Proness3}"),
                    ("StrProness4", $"{ProductInfo.Proness4}"),
                    ("StrProness5", $"{ProductInfo.Proness5}"),
                    ("StrProductName", $"{ProductInfo.ProductName}"),
                    ("StrStockName", $"{ProductInfo.StockName}"),
                    ("StrSubstrateName", $"{ProductInfo.SubstrateName}"),
                    ("StrSubstrateModel", $"{ProductInfo.SubstrateModel}"),
                    ("StrProductNumber", $"{ProductInfo.ProductNumber}"),
                    ("StrInitial", $"{ProductInfo.Initial}"),
                    ("IntRegType", $"{ProductInfo.RegType}"),
                    ("IntPrintType", $"{ProductInfo.PrintType}")
                };

            const int ColumnWidth = 15;
            var message = string.Join(Environment.NewLine,
                entries.Select(entry => $"{entry.Item1,-ColumnWidth}[{entry.Item2}]"));

            MessageBox.Show(message, "取得情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void SubstrateRegistrationPrintPreviewDialog_Load(object sender, EventArgs e) {
            var tool = (ToolStrip)SubstrateRegistrationPrintPreviewDialog.Controls[1];
            tool.Items[0].Visible = false;
        }
        private void QrCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }
    }
}
