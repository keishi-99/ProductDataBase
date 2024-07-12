using ProductDatabase.Substrate;
using System.Data;
using System.Data.SQLite;

namespace ProductDatabase {
    public partial class SubstrateRegistrationWindow : Form {

        public CSettingsLabelSub SettingsLabelSub { get; set; } = new CSettingsLabelSub();
        public string StrSettingFilePath { get; set; } = string.Empty;

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public string StrProductName { get; set; } = string.Empty;
        public string StrStockName { get; set; } = string.Empty;
        public string StrSubstrateName { get; set; } = string.Empty;
        public string StrSubstrateModel { get; set; } = string.Empty;
        public string StrInitial { get; set; } = string.Empty;
        public int IntRegType { get; set; }
        public int IntPrintType { get; set; }
        public int IntCheckBin { get; set; }

        public Bitmap LabelSubBmp { get; } = new(1, 1);
        public int LabelSubPageNum { get; set; }
        public string LabelSubNSerial { get; set; } = string.Empty;
        public int LabelSubNLabel { get; set; }
        public int LabelSubNumLabelsToPrint { get; set; }

        public decimal DisplayResolution { get; } = 96.0m;
        public int DisplayMagnitude { get; } = 3;
        public int IntPageCnt { get; set; } = 1;

        public string StrProness1 { get; set; } = string.Empty;
        public string StrProness2 { get; } = string.Empty;
        public string StrProness3 { get; } = string.Empty;
        public int StrProness4 { get; set; }
        public string StrProness5 { get; set; } = string.Empty;

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
                Font = new Font(StrFontName, IntFontSize);

                ProductNameLabel2.Text = StrProductName;
                SubstrateModelLabel2.Text = $"{StrSubstrateName} - {StrSubstrateModel}";
                StockLabel2.Text = "0";

                OrderNumberTextBox.Text = StrProness5;
                ManufacturingNumberMaskedTextBox.Text = !string.IsNullOrEmpty(StrProness1) ? StrProness1 : ManufacturingNumberMaskedTextBox.Text;
                QuantityTextBox.Text = (StrProness4 != 0) ? StrProness4.ToString() : string.Empty;

                RegisterButton.Enabled = true;

                // TextBoxへ今日の年月日を入力
                var dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = dtNow.ToShortDateString();

                // DB1へ接続し担当者取得
                using (SQLiteConnection con = new(MainWindow.GetConnectionString1())) {
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
                if (IntRegType != 0) {
                    using SQLiteConnection con = new(MainWindow.GetConnectionString2());
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - [Product_Name]_stockテーブルの[col_Substrate_Model]列の[col_Stock]の合計を取得
                    cmd.CommandText = $"SELECT total(col_stock) FROM Stock_{StrStockName} WHERE col_Substrate_Model = @col_Substrate_Model";
                    cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = StrSubstrateModel;
                    StockLabel2.Text = cmd.ExecuteScalar().ToString();

                    // テーブル検索SQL - [Substrate_Reg_[Product_Name]]テーブルの最新の[col_Revison]を取得
                    cmd.CommandText = $"SELECT col_Revision FROM Substrate_Reg_{StrStockName} WHERE col_Substrate_Model = @col_Substrate_Model AND col_Revision IS NOT NULL ORDER BY _rowid_ DESC";
                    cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = StrSubstrateModel;
                    var result = cmd.ExecuteScalar();
                    RevisionTextBox.Text = result?.ToString() ?? "";
                }
                else { StockLabel2.Text = "---"; }

                // 変数[check_bin]の値に応じてCheckboxにチェックを入れる
                foreach (var checkBoxName in _checkBoxNames) {
                    if (Controls[checkBoxName] is CheckBox checkBox) {
                        checkBox.Checked = (IntCheckBin & 0x1) == 1;
                        IntCheckBin >>= 1;
                    }
                }

                CommentComboBox.Items.Add("[Rev.UP]変更点番号:");

                StrSettingFilePath = "./config/SubstrateConfig.xml";
                if (File.Exists(StrSettingFilePath) == false) { throw new Exception("設定ファイルが見つかりませんでした"); }
                SettingsLabelSub = new CSettingsLabelSub();
                LoadSettings(StrSettingFilePath);
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
                        if (string.IsNullOrWhiteSpace(textBox.Text)) {
                            allTextBoxesFilled = false;
                            break;
                        }
                    }
                }
                if (!anyTextBoxEnabled) { throw new Exception("何も入力されていません"); }
                if (!allTextBoxesFilled) { throw new Exception("空欄があります。"); }

                if (ManufacturingNumberCheckBox.Checked && ManufacturingNumberMaskedTextBox.Text.Length != 15) { throw new Exception("製番を10桁+4桁で入力して下さい。"); }

                if (QuantityCheckBox.Checked && int.Parse(QuantityTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

                if (DefectNumberCheckBox.Checked && int.Parse(DefectNumberTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

                if (!QuantityCheckBox.Checked && !DefectNumberCheckBox.Checked) { throw new Exception("数量か不良数を入力してください。"); }

                if (!DefectNumberCheckBox.Checked && string.IsNullOrEmpty(PrintPostionNumericUpDown.Text)) { PrintPostionNumericUpDown.Text = "1"; }

                var result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) {
                    return;
                }

                RegisterButton.Enabled = false;

                if (!Registration()) { return; }

                switch (IntPrintType) {
                    case 0:
                        MessageBox.Show("登録完了");
                        break;
                    case 1:
                        if (QuantityCheckBox.Checked) {
                            MessageBox.Show("登録完了 続けて印刷します。");
                            PrintBarcode(1);
                            PrintButton.Enabled = true;
                        }
                        else if (DefectNumberCheckBox.Checked) {
                            MessageBox.Show("登録完了");
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
                using SQLiteConnection con = new(MainWindow.GetConnectionString2());
                con.Open();

                // 製番が新規かチェック
                if (IntRegType != 0) {
                    var substrateName = string.Empty;
                    using (var cmd = con.CreateCommand()) {
                        cmd.CommandText = $"SELECT * FROM Stock_{StrStockName} WHERE col_Substrate_Num = @col_Substrate_Num ORDER BY _rowid_ DESC LIMIT 1";
                        cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                        using var dr = cmd.ExecuteReader();
                        while (dr.Read()) {
                            substrateName = $"{dr["col_Substrate_Name"]}";
                        }
                    }

                    if (substrateName != string.Empty) {
                        if (StrSubstrateName == substrateName) {
                            var result = MessageBox.Show($"[{ManufacturingNumberMaskedTextBox.Text}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                            if (result == DialogResult.No) { return false; }
                        }
                        else { throw new Exception($"[{ManufacturingNumberMaskedTextBox.Text}]は[{substrateName}]として在庫があります。確認してください。"); }
                    }
                }

                // 在庫管理する基盤のみ
                if (IntRegType != 0) {
                    //在庫追加
                    if (QuantityCheckBox.Checked && !DefectNumberCheckBox.Checked) {
                        // 基板在庫テーブルへ追加_製番が一致した行を更新する、製番がない場合は新規追加
                        using var cmd = con.CreateCommand();
                        cmd.CommandText =
                            $"""
                            INSERT INTO Stock_{StrStockName}
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
                        cmd.Parameters.Add("@col_Substrate_Name", DbType.String).Value = StrSubstrateName;
                        cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = StrSubstrateModel;
                        cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                        cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = OrderNumberTextBox.Text;
                        cmd.Parameters.Add("@col_Stock", DbType.String).Value = QuantityTextBox.Text;

                        cmd.ExecuteNonQuery();
                    }
                    //不良処理
                    else if (!QuantityCheckBox.Checked && DefectNumberCheckBox.Checked) {
                        using var cmd = con.CreateCommand();
                        cmd.CommandText = $"SELECT col_Stock FROM Stock_{StrStockName} WHERE col_Substrate_Num = @col_Substrate_Num";
                        cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                        var intStock = Convert.ToInt32(cmd.ExecuteScalar());

                        if (intStock == 0) { throw new Exception("該当する製番の在庫がありません。"); }

                        var intStockFlg = (intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text)) > 0
                            ? 1
                            : (intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text)) == 0 ? 0 : throw new Exception("不良数が在庫より多く入力されています。");
                        cmd.CommandText =
                            $"""
                            UPDATE Stock_{StrStockName} SET
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
                        INSERT INTO Substrate_Reg_{StrProductName}
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
                    cmd.Parameters.Add("@col_Substrate_Name", DbType.String).Value = StrSubstrateName;
                    cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = StrSubstrateModel;
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
            // PrintDocumentオブジェクトの作成
            using System.Drawing.Printing.PrintDocument pd = new();

            //// PrintPageイベントハンドラの追加
            pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

            LabelSubNumLabelsToPrint = int.Parse(QuantityTextBox.Text);
            LabelSubPageNum = 0;

            switch (printFlg) {
                case 1:
                    SubstrateRegistrationPrintDialog.Document = pd;
                    var r = SubstrateRegistrationPrintDialog.ShowDialog();

                    if (r == DialogResult.OK) {
                        SubstrateRegistrationPrintDialog.Document.Print();

                        if (IntPageCnt >= 2) {
                            MessageBox.Show($"{IntPageCnt}枚印刷されます。2枚目以降は1行目から印刷されます。");
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
                Point offset;
                if (!SubstrateRegistrationPrintDocument.PrintController.IsPreview) {
                    offsetX -= e.PageSettings.HardMarginX * 0.254;
                    offsetY -= e.PageSettings.HardMarginY * 0.254;
                    offset = LabelSubPageNum == 0
                        ? new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (startLine * (intervalY + sizeY))))
                        : new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (intervalY + sizeY))));
                }
                else {
                    offset = new Point(0, 0);
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;
                var headerPos = SettingsLabelSub.LabelSubPageSettings.HeaderPos;
                headerPos.Offset(offset);

                var headerString = ConvertHeaderFooterString(SettingsLabelSub.LabelSubPageSettings.HeaderString);
                e.Graphics.DrawString(headerString, SettingsLabelSub.LabelSubPageSettings.HeaderFooterFont, Brushes.Black, headerPos);
                LabelSubNSerial = ManufacturingNumberMaskedTextBox.Text;

                if (LabelSubPageNum >= 1) {
                    startLine = 0;
                }

                var maxY = SettingsLabelSub.LabelSubPageSettings.NumLabelsY;
                int y;
                var intNumLabels = SettingsLabelSub.LabelSubLabelSettings.NumLabels;
                for (y = startLine; y < maxY; y++) {

                    var maxX = SettingsLabelSub.LabelSubPageSettings.NumLabelsX;

                    int x;
                    for (x = 0; x < maxX; x++) {
                        var s = GenerateCode(LabelSubNSerial);

                        var sizeX = (float)SettingsLabelSub.LabelSubPageSettings.SizeX;

                        var intervalX = SettingsLabelSub.LabelSubPageSettings.IntervalX;
                        var posX = (float)(offsetX + (x * (intervalX + sizeX)));
                        var posY = (float)(offsetY + (y * (intervalY + sizeY)));
                        e.Graphics.DrawImage(MakeLabelImage(s, (int)e.Graphics.DpiX, 1), posX, posY, sizeX, sizeY);

                        LabelSubNLabel = 0;
                        LabelSubNumLabelsToPrint--;

                        if (LabelSubNumLabelsToPrint <= 0) {
                            intNumLabels--;
                            if (intNumLabels <= 0) {
                                e.HasMorePages = false;
                                LabelSubPageNum = 0;
                                var txtNumPublish = 0;
                                LabelSubNumLabelsToPrint = txtNumPublish;
                                return;
                            }
                            else {
                                LabelSubNumLabelsToPrint += x + 1;
                                break;
                            }
                        }

                        if (x >= maxX - 1) {
                            intNumLabels--;
                            if (intNumLabels <= 0) {
                                intNumLabels = SettingsLabelSub.LabelSubLabelSettings.NumLabels;
                            }
                            else if (intNumLabels > 0) {
                                LabelSubNumLabelsToPrint += x + 1;
                                break;
                            }
                        }
                    }
                }

                if (LabelSubNumLabelsToPrint > 0) {
                    LabelSubPageNum++;
                    IntPageCnt++;
                    e.HasMorePages = true;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        private string ConvertHeaderFooterString(string s) {
            s = s.Replace("%P", StrSubstrateName)
                 .Replace("%T", StrSubstrateModel)
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
            outputCode = outputCode.Replace("%T", StrInitial)
                                    .Replace("%R", RevisionTextBox.Text)
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
                    if (checkBox.Checked) {
                        DefectNumberCheckBox.Checked = false;
                    }

                    break;
                case "DefectNumberCheckBox":
                    DefectNumberTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) {
                        QuantityCheckBox.Checked = false;
                    }

                    break;
                case "RevisionCheckBox":
                    RevisionTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) {
                        MessageBox.Show("変更する場合は理由を記載して下さい。");
                    }

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

                sw = new StreamWriter(StrSettingFilePath, false, new System.Text.UTF8Encoding(false));
                System.Xml.Serialization.XmlSerializer serializer = new(typeof(CSettingsLabelSub));
                serializer.Serialize(sw, SettingsLabelSub);
                sw?.Close();

                LoadSettings(StrSettingFilePath);
            } catch (Exception ex) {
                MessageBox.Show($"設定ファイルの保存に失敗しました。{Environment.NewLine}{ex.Message}");
            } finally {
            }
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            var message = string.Join(Environment.NewLine,
                $"StrProness1\t\t[{StrProness1}]",
                $"StrProness2\t\t[{StrProness2}]",
                $"StrProness3\t\t[{StrProness3}]",
                $"StrProness4\t\t[{StrProness4}]",
                $"StrProness5\t\t[{StrProness5}]",
                $"StrProductName\t\t[{StrProductName}]",
                $"StrStockName\t\t[{StrStockName}]",
                $"StrSubstrateName\t\t[{StrSubstrateName}]",
                $"StrSubstrateModel\t\t[{StrSubstrateModel}]",
                $"StrInitial\t\t\t[{StrInitial}]",
                $"IntRegType\t\t[{IntRegType}]",
                $"IntPrintType\t\t[{IntPrintType}]"
            );
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void SubstrateRegistrationPrintPreviewDialog_Load(object sender, EventArgs e) {
            var tool = (ToolStrip)SubstrateRegistrationPrintPreviewDialog.Controls[1];
            tool.Items[0].Visible = false;
        }
    }
}
