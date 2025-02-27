using ProductDatabase.Other;
using ProductDatabase.Substrate;
using System.Data;
using System.Data.SQLite;
using System.Drawing.Printing;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SubstrateRegistrationWindow : Form {

        public ProductInformation ProductInfo { get; }

        public CSettingsLabelSub SettingsLabelSub { get; set; } = new CSettingsLabelSub();

        private readonly string _settingFilePath = Path.Combine(Environment.CurrentDirectory, "config", "Substrate", "SubstrateConfig.xml");

        private string _labelSubNSerial = string.Empty;

        private int _labelSubPageNum;
        private int _labelSubNumLabelsToPrint;
        private int _intPageCnt = 1;

        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "DefectNumberCheckBox",
                    "RevisionCheckBox", "ExtraCheckBox1", "ExtraCheckBox2", "ExtraCheckBox3", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" ];

        // プロパティ設定
        private bool IsLabelPrint => ProductInfo.PrintType is 1;
        private bool IsRegistration => ProductInfo.RegType is 1;

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

                // DB1へ接続し担当者取得
                using (SQLiteConnection con = new(GetConnectionInformation())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        PersonComboBox.Items.Add($"{dr["PersonName"]}");
                    }
                }

                // 在庫管理する基板はDB2へ接続し対象製品の在庫取得
                //if (IsRegistration) {
                //    using SQLiteConnection con = new(GetConnectionRegistration());
                //    con.Open();
                //    using var cmd = con.CreateCommand();

                //    // テーブル検索SQL - [[ProductName]_Substrate]テーブルの最新の[Revison]を取得
                //    cmd.CommandText = $"""SELECT Revision FROM "{ProductInfo.CategoryName}_Substrate" WHERE SubstrateModel = @SubstrateModel AND Revision IS NOT NULL ORDER BY _rowid_ DESC""";
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

                if (File.Exists(_settingFilePath) == false) { throw new Exception("印刷設定ファイルが見つかりませんでした"); }
                SettingsLabelSub = new CSettingsLabelSub();
                LoadSettings(_settingFilePath);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            } finally {
            }
        }
        private void LoadSettings(string strSettingFilePath) {
            try {
                StreamReader sr = new(strSettingFilePath, new System.Text.UTF8Encoding(false));
                System.Xml.Serialization.XmlSerializer serializer = new(typeof(CSettingsLabelSub));
                if (serializer.Deserialize(sr) is CSettingsLabelSub result) { SettingsLabelSub = result; }
                sr.Close();
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
                        PrintBarcode(1);
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
            try {
                var substrateNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
                var orderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
                var quantity = string.IsNullOrWhiteSpace(QuantityTextBox.Text) ? 0 : Convert.ToInt32(QuantityTextBox.Text);
                var defectNumber = string.IsNullOrWhiteSpace(DefectNumberTextBox.Text) ? 0 : Convert.ToInt32(DefectNumberTextBox.Text);
                var registrationDate = RegistrationDateCheckBox.Checked ? RegistrationDateMaskedTextBox.Text : string.Empty;
                var person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
                var comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;

                using SQLiteConnection con = new(GetConnectionRegistration());
                con.Open();

                // 製番が新規かチェック
                if (IsRegistration) {
                    var substrateName = string.Empty;
                    using (var cmd = con.CreateCommand()) {
                        //cmd.CommandText = $"""SELECT * FROM "{ProductInfo.StockName}_StockView" WHERE SubstrateNumber = @SubstrateNumber LIMIT 1""";
                        cmd.CommandText = $"""
                            SELECT
                                StockName,
                                SubstrateName,
                                SubstrateModel,
                                SubstrateNumber,
                                OrderNumber,
                                SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                            FROM {ProductInfo.CategoryName}_Substrate
                            WHERE StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                            GROUP BY StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                            ORDER BY MIN(_rowid_)
                            LIMIT 1;
                            """;

                        cmd.Parameters.Add("@StockName", DbType.String).Value = ProductInfo.StockName;
                        cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = ProductInfo.SubstrateModel;
                        cmd.Parameters.Add("@SubstrateNumber", DbType.String).Value = substrateNumber;
                        using var dr = cmd.ExecuteReader();
                        while (dr.Read()) {
                            substrateName = $"{dr["SubstrateName"]}";
                        }
                    }

                    if (substrateName != string.Empty) {
                        if (ProductInfo.SubstrateName == substrateName) {
                            if (QuantityCheckBox.Checked && !DefectNumberCheckBox.Checked) {
                                var result = MessageBox.Show($"[{substrateNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                                if (result == DialogResult.No) { return false; }
                            }
                        }
                        //else { throw new Exception($"[{substrateNumber}]は[{substrateName}]として在庫があります。確認してください。"); }
                    }
                }

                // 不良処理時在庫チェック
                if (DefectNumberCheckBox.Checked) {
                    using var cmd = con.CreateCommand();
                    //cmd.CommandText = $"""SELECT * FROM "{ProductInfo.StockName}_StockView" WHERE SubstrateNumber = @SubstrateNumber LIMIT 1""";
                    cmd.CommandText = $"""
                            SELECT
                                StockName,
                                SubstrateName,
                                SubstrateModel,
                                SubstrateNumber,
                                OrderNumber,
                                SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                            FROM {ProductInfo.CategoryName}_Substrate
                            WHERE StockName = @StockName AND SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                            GROUP BY StockName, SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                            ORDER BY MIN(_rowid_)
                            LIMIT 1;
                            """;

                    cmd.Parameters.Add("@StockName", DbType.String).Value = ProductInfo.StockName;
                    cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = ProductInfo.SubstrateModel;
                    cmd.Parameters.Add("@SubstrateNumber", DbType.String).Value = substrateNumber;
                    using var dr = cmd.ExecuteReader();
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
                using (var cmd = con.CreateCommand()) {
                    cmd.CommandText =
                        $"""
                        INSERT INTO "{ProductInfo.CategoryName}_Substrate"
                            (
                            StockName,
                            SubstrateName,
                            SubstrateModel,
                            SubstrateNumber,
                            OrderNumber,
                            Increase,
                            Defect,
                            Person,
                            RegDate,
                            Comment
                            )
                        VALUES
                            (
                            @StockName,
                            @SubstrateName,
                            @SubstrateModel,
                            @SubstrateNumber,
                            @OrderNumber,
                            @Increase,
                            @Defect,
                            @Person,
                            @RegDate,
                            @Comment
                            )
                        """;

                    // チェックボックスにチェックがない場合はNullを
                    cmd.Parameters.Add("@StockName", DbType.String).Value = ProductInfo.StockName;
                    cmd.Parameters.Add("@SubstrateName", DbType.String).Value = ProductInfo.SubstrateName;
                    cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = ProductInfo.SubstrateModel;
                    cmd.Parameters.Add("@SubstrateNumber", DbType.String).Value = string.IsNullOrWhiteSpace(substrateNumber) ? DBNull.Value : substrateNumber;
                    cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(orderNumber) ? DBNull.Value : orderNumber;
                    cmd.Parameters.Add("@Increase", DbType.String).Value = QuantityCheckBox.Checked ? quantity : DBNull.Value;
                    cmd.Parameters.Add("@Defect", DbType.String).Value = DefectNumberCheckBox.Checked ? $"-{defectNumber}" : DBNull.Value;
                    cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(registrationDate) ? DBNull.Value : registrationDate;
                    cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(person) ? DBNull.Value : person;
                    cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(comment) ? DBNull.Value : comment;

                    cmd.ExecuteNonQuery();
                }

                // バックアップ作成
                CommonUtils.BackupManager.CreateBackup();
                // ログ出力
                var number = QuantityCheckBox.Checked ? quantity : 0 - defectNumber;
                CommonUtils.Logger.AppendLog($"[基板登録];注文番号[{orderNumber}];製造番号[{substrateNumber}];製品名[{ProductInfo.ProductName}];基板名[{ProductInfo.SubstrateName}];型式[{ProductInfo.SubstrateModel}];数量[{number}];登録日[{registrationDate}];担当者[{person}];");

                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        // 印刷処理
        private void PrintBarcode(int printFlg) {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();

                //// PrintPageイベントハンドラの追加
                pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

                _labelSubNumLabelsToPrint = int.TryParse(QuantityTextBox.Text, out var quantity) ? quantity : 0;
                if (_labelSubNumLabelsToPrint == 0) {
                    throw new Exception("数量が入力されていません。");
                }
                _labelSubPageNum = 0;

                switch (printFlg) {
                    case 1:
                        SubstrateRegistrationPrintDialog.Document = pd;
                        var r = SubstrateRegistrationPrintDialog.ShowDialog();

                        if (r == DialogResult.OK) {
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
                        break;
                    case 2:
                        SubstrateRegistrationPrintPreviewDialog.PrintPreviewControl.Zoom = 3;
                        SubstrateRegistrationPrintPreviewDialog.Document = pd;
                        SubstrateRegistrationPrintPreviewDialog.ShowDialog();
                        break;
                    default:
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            try {
                if (SettingsLabelSub == null || e.Graphics == null) { return; }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                var startLine = (int)PrintPostionNumericUpDown.Value - 1;
                var sizeX = (float)SettingsLabelSub.LabelSubPageSettings.SizeX;
                var sizeY = (float)SettingsLabelSub.LabelSubPageSettings.SizeY;
                var offsetX = SettingsLabelSub.LabelSubPageSettings.OffsetX;
                var offsetY = SettingsLabelSub.LabelSubPageSettings.OffsetY;
                var intervalX = SettingsLabelSub.LabelSubPageSettings.IntervalX;
                var intervalY = SettingsLabelSub.LabelSubPageSettings.IntervalY;
                var headerPos = SettingsLabelSub.LabelSubPageSettings.HeaderPos;
                var offset = new System.Drawing.Point(0, 0);
                const double MM_PER_HUNDREDTH_INCH = 0.254;

                var pd = (PrintDocument)sender;
                var bPrintMode = pd.PrintController.IsPreview;

                if (!bPrintMode) {
                    offsetX -= e.PageSettings.HardMarginX * 0.254;
                    offsetY -= e.PageSettings.HardMarginY * 0.254;
                    offset = _labelSubPageNum == 0
                        ? new System.Drawing.Point((int)(e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)((e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (startLine * (intervalY + sizeY))))
                        : new System.Drawing.Point((int)(e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)((e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (0 * (intervalY + sizeY))));
                }
                else {
                    offset = new System.Drawing.Point(0, 0);
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;

                var headerString = ConvertHeaderFooterString(SettingsLabelSub.LabelSubPageSettings.HeaderString);
                headerPos.Offset(offset);
                e.Graphics.DrawString(headerString, SettingsLabelSub.LabelSubPageSettings.HeaderFooterFont, Brushes.Black, headerPos);
                _labelSubNSerial = ManufacturingNumberMaskedTextBox.Text;

                if (_labelSubPageNum >= 1) {
                    startLine = 0;
                }

                var maxX = SettingsLabelSub.LabelSubPageSettings.NumLabelsX;
                var maxY = SettingsLabelSub.LabelSubPageSettings.NumLabelsY;
                int y;
                var serialCodePrintCopies = SettingsLabelSub.LabelSubLabelSettings.NumLabels;
                if (maxX == 0 || maxY == 0 || serialCodePrintCopies == 0) { throw new Exception("印刷設定が異常です。"); }
                for (y = startLine; y < maxY; y++) {
                    int x;
                    for (x = 0; x < maxX; x++) {
                        var posX = (float)(offsetX + (x * (intervalX + sizeX)));
                        var posY = (float)(offsetY + (y * (intervalY + sizeY)));

                        var generatedCode = GenerateCode(_labelSubNSerial);
                        var labelImage = MakeLabelImage(generatedCode, (int)e.Graphics.DpiX, 1);
                        e.Graphics.DrawImage(labelImage, posX, posY, sizeX, sizeY);

                        _labelSubNumLabelsToPrint--;

                        if (_labelSubNumLabelsToPrint <= 0) {
                            serialCodePrintCopies--;
                            if (serialCodePrintCopies <= 0) {
                                e.HasMorePages = false;
                                _labelSubPageNum = 0;
                                var txtNumPublish = 0;
                                _labelSubNumLabelsToPrint = txtNumPublish;
                                return;
                            }
                            else {
                                _labelSubNumLabelsToPrint += x + 1;
                                break;
                            }
                        }

                        if (x >= maxX - 1) {
                            serialCodePrintCopies--;
                            if (serialCodePrintCopies <= 0) {
                                serialCodePrintCopies = SettingsLabelSub.LabelSubLabelSettings.NumLabels;
                            }
                            else if (serialCodePrintCopies > 0) {
                                _labelSubNumLabelsToPrint += x + 1;
                                break;
                            }
                        }
                    }
                }

                if (_labelSubNumLabelsToPrint > 0) {
                    _labelSubPageNum++;
                    _intPageCnt++;
                    e.HasMorePages = true;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        private string ConvertHeaderFooterString(string s) {
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

            var outputCode = SettingsLabelSub.LabelSubLabelSettings.Format;
            var serialCode = serial.Substring(5, 5);
            outputCode = outputCode.Replace("%T", ProductInfo.Initial)
                                    .Replace("%Y", DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("MM"))
                                    .Replace("%M", string.IsNullOrEmpty(monthCode) ? string.Empty : monthCode[^1..])
                                    .Replace("%S", serialCode);

            return outputCode;
        }
        private Bitmap MakeLabelImage(string text, int resolution, int magnitude) {
            if (SettingsLabelSub is null) { throw new Exception(); }
            var sizeX = (decimal)SettingsLabelSub.LabelSubPageSettings.SizeX / 25.4M * resolution * magnitude;
            var sizeY = (decimal)SettingsLabelSub.LabelSubPageSettings.SizeY / 25.4M * resolution * magnitude;

            Bitmap labelImage = new((int)sizeX, (int)sizeY);
            using (var g = Graphics.FromImage(labelImage)) {
                var fontSize = (decimal)SettingsLabelSub.LabelSubLabelSettings.Font.SizeInPoints / 72.0M * resolution * magnitude;
                using (Font fnt = new(SettingsLabelSub.LabelSubLabelSettings.Font.Name, (float)fontSize)) {
                    //var stringSize = g.MeasureString(text, fnt);
                    var stringSize = TextRenderer.MeasureText(text, fnt);

                    var stringPosX = (labelImage.Width / 2f) - (stringSize.Width / 2f);
                    var stringPosY = (int)((decimal)SettingsLabelSub.LabelSubLabelSettings.StringPosY / 25.4M * resolution * magnitude);

                    g.DrawString(text, fnt, Brushes.Black, stringPosX, stringPosY);
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
        private void PrintButton_Click(object sender, EventArgs e) { PrintBarcode(1); }
        private void SubstrateRegistrationPrintDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) { PrintDocumentPrintPage(sender, e); }
        private void TemplateButton_Click(object sender, EventArgs e) { TemplateComment(); }
        private void NumberCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBoxChecked(sender, e); }
        private void QuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void DefectNumberTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void RegistrationDateMaskedTextBox_TypeValidationCompleted(object sender, TypeValidationEventArgs e) { RegistrationDateCheck(sender, e); }
        private void 印刷ToolStripMenuItem_Click(object sender, EventArgs e) { PrintBarcode(1); }
        private void 印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) { PrintBarcode(2); }
        private void 印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            SubstrateLabelSettingsWindow ls = new();
            ls.ShowDialog(this);

            try {
                StreamWriter? sw = null;

                sw = new StreamWriter(_settingFilePath, false, new System.Text.UTF8Encoding(false));
                System.Xml.Serialization.XmlSerializer serializer = new(typeof(CSettingsLabelSub));
                serializer.Serialize(sw, SettingsLabelSub);
                sw?.Close();

                LoadSettings(_settingFilePath);
            } catch (Exception ex) {
                MessageBox.Show($"設定ファイルの保存に失敗しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
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
    }
}
