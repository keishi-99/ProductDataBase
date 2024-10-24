using ProductDatabase.Substrate;
using System.Data;
using System.Data.SQLite;
using System.Drawing.Printing;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SubstrateRegistrationWindow : Form {

        public ProductInfomation ProductInfo { get; set; } = new ProductInfomation();

        public CSettingsLabelSub SettingsLabelSub { get; set; } = new CSettingsLabelSub();

        private readonly string _settingFilePath = "./config/SubstrateConfig.xml";


        private string _labelSubNSerial = String.Empty;

        private int _labelSubPageNum;
        private int _labelSubNumLabelsToPrint;
        private int _intPageCnt = 1;

        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "DefectNumberCheckBox",
                    "RevisionCheckBox", "ExtraCheckBox1", "ExtraCheckBox2", "ExtraCheckBox3", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" ];

        public SubstrateRegistrationWindow() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);

                ProductNameLabel2.Text = ProductInfo.ProductName;

                SubstrateModelLabel2.Text = $"{ProductInfo.SubstrateName} - {ProductInfo.SubstrateModel}";

                OrderNumberTextBox.Text = ProductInfo.Proness5;
                ManufacturingNumberMaskedTextBox.Text = !String.IsNullOrEmpty(ProductInfo.Proness1) ? ProductInfo.Proness1 : ManufacturingNumberMaskedTextBox.Text;
                QuantityTextBox.Text = (ProductInfo.Proness4 != 0) ? ProductInfo.Proness4.ToString() : String.Empty;

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
                using (SQLiteConnection con = new(GetConnectionString1())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        PersonComboBox.Items.Add($"{dr["col_Person_Name"]}");
                    }
                }

                // 在庫管理する基盤はDB2へ接続し対象製品の在庫取得
                if (ProductInfo.RegType != 0) {
                    using SQLiteConnection con = new(GetConnectionString2());
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - [Product_Name]_stockテーブルの[col_Substrate_Model]列の[col_Stock]の合計を取得
                    cmd.CommandText = $"""SELECT total(col_stock) FROM "Stock_{ProductInfo.StockName}" WHERE col_Substrate_Model = @col_Substrate_Model""";
                    cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ProductInfo.SubstrateModel;
                    StockLabel2.Text = cmd.ExecuteScalar().ToString();

                    // テーブル検索SQL - [Substrate_Reg_[Product_Name]]テーブルの最新の[col_Revison]を取得
                    cmd.CommandText = $"""SELECT col_Revision FROM "Substrate_Reg_{ProductInfo.StockName}" WHERE col_Substrate_Model = @col_Substrate_Model AND col_Revision IS NOT NULL ORDER BY _rowid_ DESC""";
                    cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ProductInfo.SubstrateModel;
                    var result = cmd.ExecuteScalar();
                    RevisionTextBox.Text = result?.ToString() ?? "";
                }
                else { StockLabel2.Text = "---"; }

                if (File.Exists(_settingFilePath) == false) { throw new Exception("設定ファイルが見つかりませんでした"); }
                SettingsLabelSub = new CSettingsLabelSub();
                LoadSettings(_settingFilePath);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // クロージングイベント
        private static void ClosingEvents() {
            //try {
            //    StreamWriter? _sw = null;

            //    _sw = new StreamWriter(StrSettingFilePath, false, new System.Text.UTF8Encoding(false));
            //    System.Xml.Serialization.XmlSerializer _serializer = new(typeof(CSettingsLabelSub));
            //    _serializer.Serialize(_sw, SettingsLabelSub);
            //    _sw?.Close();
            //} catch (Exception ex) {
            //    MessageBox.Show($"設定ファイルの保存に失敗しました。{Environment.NewLine}{ex.Message}");
            //} finally {
            //    Dispose();
            //}
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
                        if (String.IsNullOrWhiteSpace(textBox.Text)) {
                            allTextBoxesFilled = false;
                            break;
                        }
                    }
                }
                if (!anyTextBoxEnabled) { throw new Exception("何も入力されていません"); }
                if (!allTextBoxesFilled) { throw new Exception("空欄があります。"); }

                if (ManufacturingNumberCheckBox.Checked && ManufacturingNumberMaskedTextBox.Text.Length != 15) { throw new Exception("製番を10桁+4桁で入力して下さい。"); }

                if (QuantityCheckBox.Checked && Int32.Parse(QuantityTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

                if (DefectNumberCheckBox.Checked && Int32.Parse(DefectNumberTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

                if (!QuantityCheckBox.Checked && !DefectNumberCheckBox.Checked) { throw new Exception("数量か不良数を入力してください。"); }

                if (!DefectNumberCheckBox.Checked && String.IsNullOrEmpty(PrintPostionNumericUpDown.Text)) { PrintPostionNumericUpDown.Text = "1"; }

                var result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) {
                    return;
                }

                RegisterButton.Enabled = false;

                if (!Registration()) { return; }

                switch (ProductInfo.PrintType) {
                    case 0:
                        MessageBox.Show("登録完了");
                        break;
                    case 1:
                        if (QuantityCheckBox.Checked) {
                            MessageBox.Show("登録完了 続けて印刷します。");
                            PrintBarcode(1);
                            PrintButton.Enabled = true;
                            break;
                        }
                        if (DefectNumberCheckBox.Checked) {
                            MessageBox.Show("登録完了");
                            break;
                        }
                        break;
                    default:
                        break;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                RegisterButton.Enabled = true;
            }
        }
        private bool Registration() {
            try {
                using SQLiteConnection con = new(GetConnectionString2());
                con.Open();

                // 製番が新規かチェック
                if (ProductInfo.RegType != 0) {
                    var substrateName = String.Empty;
                    using (var cmd = con.CreateCommand()) {
                        cmd.CommandText = $"""SELECT * FROM "Stock_{ProductInfo.StockName}" WHERE col_Substrate_Num = @col_Substrate_Num ORDER BY _rowid_ DESC LIMIT 1""";
                        cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                        using var dr = cmd.ExecuteReader();
                        while (dr.Read()) {
                            substrateName = $"{dr["col_Substrate_Name"]}";
                        }
                    }

                    if (substrateName != String.Empty) {
                        if (ProductInfo.SubstrateName == substrateName) {
                            if (ManufacturingNumberMaskedTextBox.Text != "宮崎在庫") {
                                var result = MessageBox.Show($"[{ManufacturingNumberMaskedTextBox.Text}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                                if (result == DialogResult.No) { return false; }
                            }
                        }
                        else { throw new Exception($"[{ManufacturingNumberMaskedTextBox.Text}]は[{substrateName}]として在庫があります。確認してください。"); }
                    }
                }

                // 在庫管理する基盤のみ
                if (ProductInfo.RegType != 0) {
                    //在庫追加
                    if (QuantityCheckBox.Checked && !DefectNumberCheckBox.Checked) {
                        // 基板在庫テーブルへ追加_製番が一致した行を更新する、製番がない場合は新規追加
                        using var cmd = con.CreateCommand();
                        cmd.CommandText =
                            $"""
                            INSERT INTO "Stock_{ProductInfo.StockName}"
                                (
                                col_Flg, col_Substrate_Name, col_Substrate_Model, col_Substrate_Num, col_Order_Num, col_Stock
                                )
                            VALUES
                                (
                                @col_Flg, @col_Substrate_Name, @col_Substrate_Model, @col_Substrate_Num, @col_Order_Num, @col_Stock
                                )
                                on conflict(col_Substrate_Num)
                            DO UPDATE
                                set col_Flg = 1, col_Stock = col_Stock + excluded.col_Stock
                            """;

                        cmd.Parameters.Add("@col_Flg", DbType.String).Value = 1;
                        cmd.Parameters.Add("@col_Substrate_Name", DbType.String).Value = ProductInfo.SubstrateName;
                        cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ProductInfo.SubstrateModel;
                        cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                        cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = OrderNumberTextBox.Text;
                        cmd.Parameters.Add("@col_Stock", DbType.String).Value = QuantityTextBox.Text;

                        cmd.ExecuteNonQuery();
                    }
                    //不良処理
                    else if (!QuantityCheckBox.Checked && DefectNumberCheckBox.Checked) {
                        using var cmd = con.CreateCommand();
                        cmd.CommandText = $"""SELECT col_Stock FROM "Stock_{ProductInfo.StockName}" WHERE col_Substrate_Num = @col_Substrate_Num""";
                        cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                        var intStock = Convert.ToInt32(cmd.ExecuteScalar());

                        if (intStock == 0) { throw new Exception("該当する製番の在庫がありません。"); }

                        var intStockFlg = (intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text)) > 0
                            ? 1
                            : (intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text)) == 0 ? 0 : throw new Exception("不良数が在庫より多く入力されています。");
                        cmd.CommandText =
                            $"""
                            UPDATE "Stock_{ProductInfo.StockName}" SET
                                col_Flg = @col_Flg,
                                col_Stock = @col_Stock,
                                col_History = ifnull(col_History, '') || @col_History
                            WHERE
                                col_Substrate_Num = @col_Substrate_Num
                            """;
                        cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;

                        cmd.Parameters.Add("@col_Flg", DbType.String).Value = intStockFlg;
                        cmd.Parameters.Add("@col_Stock", DbType.String).Value = intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text);
                        cmd.Parameters.Add("@col_History", DbType.String).Value = $"[不良]{ManufacturingNumberMaskedTextBox.Text}";

                        cmd.ExecuteNonQuery();
                    }
                }

                // 基板登録テーブルへ追加
                using (var cmd = con.CreateCommand()) {
                    cmd.CommandText =
                        $"""
                        INSERT INTO "Substrate_Reg_{ProductInfo.ProductName}"
                            (
                            col_Substrate_Name,
                            col_Substrate_Model,
                            col_Substrate_Num,
                            col_Order_Num,
                            col_Increase,
                            col_Defect,
                            col_Person,
                            col_RegDate,
                            col_Revision,
                            col_Comment
                            )
                        VALUES
                            (
                            @col_Substrate_Name,
                            @col_Substrate_Model,
                            @col_Substrate_Num,
                            @col_Order_Num,
                            @col_Increase,
                            @col_Defect,
                            @col_Person,
                            @col_RegDate,
                            @col_Revision,
                            @col_Comment
                            )
                        """;

                    // チェックボックスにチェックがない場合はNullを
                    cmd.Parameters.Add("@col_Substrate_Name", DbType.String).Value = ProductInfo.SubstrateName;
                    cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = ProductInfo.SubstrateModel;
                    cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                    cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = OrderNumberTextBox.Text;
                    cmd.Parameters.Add("@col_Increase", DbType.String).Value = QuantityCheckBox.Checked ? QuantityTextBox.Text : DBNull.Value;
                    cmd.Parameters.Add("@col_Defect", DbType.String).Value = DefectNumberCheckBox.Checked ? DefectNumberTextBox.Text : DBNull.Value;
                    cmd.Parameters.Add("@col_RegDate", DbType.String).Value = RegistrationDateMaskedTextBox.Text;
                    cmd.Parameters.Add("@col_Person", DbType.String).Value = PersonComboBox.Text;
                    cmd.Parameters.Add("@col_Revision", DbType.String).Value = RevisionTextBox.Text;
                    cmd.Parameters.Add("@col_Comment", DbType.String).Value = CommentTextBox.Text;

                    cmd.ExecuteNonQuery();
                }

                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                _labelSubNumLabelsToPrint = Int32.TryParse(QuantityTextBox.Text, out var quantity) ? quantity : 0;
                if (_labelSubNumLabelsToPrint == 0) {
                    throw new Exception("数量が入力されていません。");
                }
                _labelSubPageNum = 0;

                switch (printFlg) {
                    case 1:
                        SubstrateRegistrationPrintDialog.Document = pd;
                        var r = SubstrateRegistrationPrintDialog.ShowDialog();

                        if (r == DialogResult.OK) {
                            SubstrateRegistrationPrintDialog.Document.Print();

                            if (_intPageCnt >= 2) {
                                MessageBox.Show($"{_intPageCnt}枚印刷されます。2枚目以降は1行目から印刷されます。");
                            }
                        }
                        else {
                            return;
                        }
                        break;
                    case 2:
                        SubstrateRegistrationPrintPreviewDialog.Document = pd;
                        SubstrateRegistrationPrintPreviewDialog.ShowDialog();
                        break;
                    default:
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            try {
                if (SettingsLabelSub == null || e.Graphics == null) { return; }

                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                var startLine = (int)PrintPostionNumericUpDown.Value - 1;
                var sizeY = (float)SettingsLabelSub.LabelSubPageSettings.SizeY;
                var offsetX = SettingsLabelSub.LabelSubPageSettings.OffsetX;
                var offsetY = SettingsLabelSub.LabelSubPageSettings.OffsetY;
                var intervalY = SettingsLabelSub.LabelSubPageSettings.IntervalY;
                var headerPos = SettingsLabelSub.LabelSubPageSettings.HeaderPos;
                var offset = new Point(0, 0);
                const double MM_PER_HUNDREDTH_INCH = 0.254;

                var pd = (PrintDocument)sender;
                var bPrintMode = pd.PrintController.IsPreview;

                if (!bPrintMode) {
                    offsetX -= e.PageSettings.HardMarginX * 0.254;
                    offsetY -= e.PageSettings.HardMarginY * 0.254;
                    offset = _labelSubPageNum == 0
                        ? new Point((int)(e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)((e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (startLine * (intervalY + sizeY))))
                        : new Point((int)(e.PageSettings.HardMarginX * -MM_PER_HUNDREDTH_INCH), (int)((e.PageSettings.HardMarginY * -MM_PER_HUNDREDTH_INCH) + (0 * (intervalY + sizeY))));
                }
                else {
                    offset = new Point(0, 0);
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

                var maxY = SettingsLabelSub.LabelSubPageSettings.NumLabelsY;
                int y;
                var intNumLabels = SettingsLabelSub.LabelSubLabelSettings.NumLabels;
                for (y = startLine; y < maxY; y++) {

                    var maxX = SettingsLabelSub.LabelSubPageSettings.NumLabelsX;

                    int x;
                    for (x = 0; x < maxX; x++) {
                        var generatedCode = GenerateCode(_labelSubNSerial);

                        var sizeX = (float)SettingsLabelSub.LabelSubPageSettings.SizeX;

                        var intervalX = SettingsLabelSub.LabelSubPageSettings.IntervalX;
                        var posX = (float)(offsetX + (x * (intervalX + sizeX)));
                        var posY = (float)(offsetY + (y * (intervalY + sizeY)));
                        e.Graphics.DrawImage(MakeLabelImage(generatedCode, (int)e.Graphics.DpiX, 1), posX, posY, sizeX, sizeY);

                        _labelSubNumLabelsToPrint--;

                        if (_labelSubNumLabelsToPrint <= 0) {
                            intNumLabels--;
                            if (intNumLabels <= 0) {
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
                            intNumLabels--;
                            if (intNumLabels <= 0) {
                                intNumLabels = SettingsLabelSub.LabelSubLabelSettings.NumLabels;
                            }
                            else if (intNumLabels > 0) {
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
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                 .Replace("%U", PersonComboBox.Text);
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
                                    .Replace("%R", RevisionTextBox.Text)
                                    .Replace("%Y", DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("MM"))
                                    .Replace("%M", String.IsNullOrEmpty(monthCode) ? String.Empty : monthCode[^1..])
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
                    var stringSize = g.MeasureString(text, fnt);

                    var stringPosX = (int)((labelImage.Width / 2) - (stringSize.Width / 2));
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
                _ => String.Empty
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
                    RevisionTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) { MessageBox.Show("変更する場合は理由を記載して下さい。"); }
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

        private void SubstrateRegistrationWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void SubstrateRegistrationWindow_FormClosing(object sender, FormClosingEventArgs e) { ClosingEvents(); }
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
            SubstratePrintSetting ls = new();
            ls.ShowDialog(this);

            try {
                StreamWriter? sw = null;

                sw = new StreamWriter(_settingFilePath, false, new System.Text.UTF8Encoding(false));
                System.Xml.Serialization.XmlSerializer serializer = new(typeof(CSettingsLabelSub));
                serializer.Serialize(sw, SettingsLabelSub);
                sw?.Close();

                LoadSettings(_settingFilePath);
            } catch (Exception ex) {
                MessageBox.Show($"設定ファイルの保存に失敗しました。{Environment.NewLine}{ex.Message}");
            } finally {
            }
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            var message = String.Join(Environment.NewLine,
                $"StrProness1\t\t[{ProductInfo.Proness1}]",
                $"StrProness2\t\t[{ProductInfo.Proness2}]",
                $"StrProness3\t\t[{ProductInfo.Proness3}]",
                $"StrProness4\t\t[{ProductInfo.Proness4}]",
                $"StrProness5\t\t[{ProductInfo.Proness5}]",
                $"StrProductName\t\t[{ProductInfo.ProductName}]",
                $"StrStockName\t\t[{ProductInfo.StockName}]",
                $"StrSubstrateName\t\t[{ProductInfo.SubstrateName}]",
                $"StrSubstrateModel\t\t[{ProductInfo.SubstrateModel}]",
                $"StrInitial\t\t\t[{ProductInfo.Initial}]",
                $"IntRegType\t\t[{ProductInfo.RegType}]",
                $"IntPrintType\t\t[{ProductInfo.PrintType}]"
            );
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void SubstrateRegistrationPrintPreviewDialog_Load(object sender, EventArgs e) {
            var tool = (ToolStrip)SubstrateRegistrationPrintPreviewDialog.Controls[1];
            tool.Items[0].Visible = false;
        }
    }
}
