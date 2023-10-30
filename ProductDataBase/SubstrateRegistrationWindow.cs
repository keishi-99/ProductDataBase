using LabelGenerator;
using System.Data;
using System.Data.SQLite;

namespace ProductDataBase {
    public partial class SubstrateRegistrationWindow : Form {

        public CSettingsLabelSub? SettingsLabelSub { get; set; }
        public string? StrSettingFilePath { get; set; } = string.Empty;

        public string? StrFontName { get; set; }
        public int IntFontSize { get; set; }

        public string? StrProductName { get; set; }
        public string? StrStockName { get; set; }
        public string? StrSubstrateName { get; set; }
        public string? StrSubstrateModel { get; set; }
        public string? StrInitial { get; set; }
        public int IntRegType { get; set; }
        public int IntPrintType { get; set; }
        public int IntCheckBin { get; set; }

        public Bitmap? LabelSubBmp { get; }
        public int LabelSubPageNum { get; set; }
        public string? LabelSubNSerial { get; set; }
        public int LabelSubNLabel { get; set; }
        public int LabelSubNumLabelsToPrint { get; set; }

        public decimal LabelSubDisplayResolution { get; } = 96.0m;
        public int LabelSubDisplayMagnitude { get; } = 3;
        public int IntPageCnt { get; set; } = 1;

        public string? StrProness1 { get; }
        public string? StrProness2 { get; }
        public string? StrProness3 { get; }
        public int StrProness4 { get; }
        public string? StrProness5 { get; }

        public SubstrateRegistrationWindow(MainWindow mainWindow) {
            StrFontName = mainWindow.StrFontName;
            IntFontSize = mainWindow.IntFontSize;
            StrProductName = mainWindow.StrProductName;
            StrStockName = mainWindow.StrStockName;
            StrSubstrateName = mainWindow.StrSubstrateName;
            StrSubstrateModel = mainWindow.StrSubstrateModel;
            StrInitial = mainWindow.StrInitial;
            IntRegType = mainWindow.IntRegType;
            IntPrintType = mainWindow.IntPrintType;
            IntCheckBin = mainWindow.IntCheckBin;

            StrProness1 = mainWindow.StrProness1;
            StrProness2 = mainWindow.StrProness2;
            StrProness3 = mainWindow.StrProness3;
            StrProness4 = mainWindow.StrProness4;
            StrProness5 = mainWindow.StrProness5;
        }

        public SubstrateRegistrationWindow() {
            InitializeComponent();
        }

        private void LoadEvents() {
            try {
                if (StrFontName != null) { Font = new Font(StrFontName, IntFontSize); }

                ProductNameLabel2.Text = StrProductName;
                SubstrateModelLabel2.Text = $"{StrSubstrateName} - {StrSubstrateModel}";
                StockLabel2.Text = "0";

                if (StrProness5 != null) OrderNumberTextBox.Text = StrProness5;
                if (StrProness1 != null) ManufacturingNumberMaskedTextBox.Text = StrProness1;
                if (StrProness4 != 0) QuantityTextBox.Text = StrProness4.ToString();

                RegisterButton.Enabled = true;

                // TextBoxへ今日の年月日を入力
                DateTime dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = dtNow.ToShortDateString();

                // DB1へ接続し担当者取得
                using (SQLiteConnection? con = new(MainWindow.GetConnectionString1())) {
                    con.Open();
                    using SQLiteCommand cmd = con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using SQLiteDataReader dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        PersonComboBox.Items.Add($"{dr["col_Person_Name"]}");
                    }
                }

                // 在庫管理する基盤はDB2へ接続し対象製品の在庫取得
                if (IntRegType != 0) {
                    string stock = string.Empty;
                    using SQLiteConnection con = new(MainWindow.GetConnectionString2());
                    con.Open();
                    using SQLiteCommand cmd = con.CreateCommand();
                    // テーブル検索SQL - [Product_Name]_stockテーブルの[col_Substrate_Model]列の[col_Stock]の合計を取得
                    cmd.CommandText = $"SELECT total(col_stock) FROM 'Stock_{StrStockName}' WHERE col_Substrate_Model = '{StrSubstrateModel}'";
                    StockLabel2.Text = cmd.ExecuteScalar().ToString();

                    // テーブル検索SQL - [Substrate_Reg_[Product_Name]]テーブルの最新の[col_Revison]を取得
                    cmd.CommandText = $"SELECT col_Revision FROM 'Substrate_Reg_{StrStockName}' WHERE col_Substrate_Model = '{StrSubstrateModel}' and col_Revision IS NOT NULL ORDER BY _rowid_ DESC";
                    object result = cmd.ExecuteScalar();
                    RevisionTextBox.Text = result?.ToString() ?? "";
                }
                else { StockLabel2.Text = "---"; }

                // 変数[check_bin]の値に応じてCheckboxにチェックを入れる
                List<string> checkBoxNames = new() {
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "DefectNumberCheckBox",
                    "RevisionCheckBox", "ExtraCheckBox1", "ExtraCheckBox2", "ExtraCheckBox3", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" };

                foreach (string checkBoxName in checkBoxNames) {
                    if (Controls[checkBoxName] is CheckBox checkBox) {
                        checkBox.Checked = (IntCheckBin & 0x1) == 1;
                        IntCheckBin >>= 1;
                    }
                }

                StrSettingFilePath = "./config/SubstrateConfig.xml";
                if (System.IO.File.Exists(StrSettingFilePath) == false) { throw new Exception("設定ファイルが見つかりませんでした"); }
                SettingsLabelSub = new CSettingsLabelSub();
                LoadSettings(StrSettingFilePath);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        private void LoadSettings(string StrSettingFilePath) {
            System.Xml.Serialization.XmlSerializer serializer = new(typeof(CSettingsLabelSub));

            try {
                StreamReader? sr = new(StrSettingFilePath, new System.Text.UTF8Encoding(false));
                SettingsLabelSub = serializer.Deserialize(sr) as CSettingsLabelSub;
                sr?.Close();
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        private void ClosingEvents() {
            try {
                System.Xml.Serialization.XmlSerializer serializer = new(typeof(CSettingsLabelSub));
                System.IO.StreamWriter? sw = null;

                if (StrSettingFilePath == null) { throw new Exception("StrSettingFilePathがnull"); }
                sw = new System.IO.StreamWriter(StrSettingFilePath, false, new System.Text.UTF8Encoding(false));
                serializer.Serialize(sw, SettingsLabelSub);
                sw?.Close();
            } catch (Exception ex) {
                MessageBox.Show($"設定ファイルの保存に失敗しました。{Environment.NewLine}{ex.Message}");
            } finally {
                Dispose();
            }
        }
        // 登録処理
        private void RegisterCheck() {
            try {
                // 印刷のみにチェックがある場合処理をしない
                if (PrintOnlyCheckBox.Checked) { return; }

                // 入力フォームのチェック
                bool anyTextBoxEnabled = false;
                bool allTextBoxesFilled = true;

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

                DialogResult result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) return;

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
                    string Substrate_Name = string.Empty;
                    using (SQLiteCommand cmd = con.CreateCommand()) {
                        cmd.CommandText = $"SELECT * FROM 'Stock_{StrStockName}' WHERE col_Substrate_Num = '{ManufacturingNumberMaskedTextBox.Text}' ORDER BY _rowid_ DESC LIMIT 1";
                        using SQLiteDataReader dr = cmd.ExecuteReader();
                        while (dr.Read()) {
                            Substrate_Name = $"{dr["col_Substrate_Name"]}";
                        }
                    }

                    if (Substrate_Name != string.Empty) {
                        if (StrSubstrateName == Substrate_Name) {
                            DialogResult dr = MessageBox.Show($"[{ManufacturingNumberMaskedTextBox.Text}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                            if (dr == DialogResult.No) { return false; }
                        }
                        else { throw new Exception($"[{ManufacturingNumberMaskedTextBox.Text}]は[{Substrate_Name}]として在庫があります。確認してください。"); }
                    }
                }

                // 在庫管理する基盤のみ
                if (IntRegType != 0) {
                    //在庫追加
                    if (QuantityCheckBox.Checked && !DefectNumberCheckBox.Checked) {
                        // 基板在庫テーブルへ追加_製番が一致した行を更新する、製番がない場合は新規追加
                        using SQLiteCommand cmd = con.CreateCommand();
                        cmd.CommandText =
                            $"INSERT INTO 'Stock_{StrStockName}'" +
                            $"(col_Flg, col_Substrate_Name, col_Substrate_Model, col_Substrate_Num, col_Order_Num, col_Stock)" +
                            $" VALUES" +
                            $"(@col_Flg, @col_Substrate_Name, @col_Substrate_Model, @col_Substrate_Num, @col_Order_Num, @col_Stock)" +
                            $" on conflict(col_Substrate_Num)" +
                            $" do update" +
                            $" set col_Flg = 1, col_Stock = col_Stock + excluded.col_Stock";

                        cmd.Parameters.Add("@col_Flg", DbType.AnsiString).Value = 1;
                        cmd.Parameters.Add("@col_Substrate_Name", DbType.AnsiString).Value = StrSubstrateName;
                        cmd.Parameters.Add("@col_Substrate_Model", DbType.AnsiString).Value = StrSubstrateModel;
                        cmd.Parameters.Add("@col_Substrate_Num", DbType.AnsiString).Value = ManufacturingNumberMaskedTextBox.Text;
                        cmd.Parameters.Add("@col_Order_Num", DbType.AnsiString).Value = OrderNumberTextBox.Text;
                        cmd.Parameters.Add("@col_Stock", DbType.AnsiString).Value = QuantityTextBox.Text;

                        cmd.ExecuteNonQuery();
                    }
                    //不良処理
                    else if (!QuantityCheckBox.Checked && DefectNumberCheckBox.Checked) {
                        int intStock_Flg;
                        int intStock;

                        using SQLiteCommand cmd = con.CreateCommand();
                        cmd.CommandText = $"SELECT col_Stock FROM 'Stock_{StrStockName}' WHERE col_Substrate_Num = '{ManufacturingNumberMaskedTextBox.Text}'";
                        intStock = Convert.ToInt32(cmd.ExecuteScalar());

                        if (intStock == 0) { throw new Exception("該当する製番の在庫がありません。"); }

                        if ((intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text)) > 0) { intStock_Flg = 1; }
                        else if ((intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text)) == 0) { intStock_Flg = 0; }
                        else { throw new Exception("不良数が在庫より多く入力されています。"); }

                        cmd.CommandText =
                            $"UPDATE 'Stock_{StrStockName}'SET" +
                            $" col_Flg = @col_Flg," +
                            $"col_Stock = @col_Stock," +
                            $"col_History = ifnull(col_History, '') || @col_History" +
                            $" WHERE" +
                            $" col_Substrate_Num = '{ManufacturingNumberMaskedTextBox.Text}'";


                        cmd.Parameters.Add("@col_Flg", DbType.AnsiString).Value = intStock_Flg;
                        cmd.Parameters.Add("@col_Stock", DbType.AnsiString).Value = intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text);
                        cmd.Parameters.Add("@col_History", DbType.AnsiString).Value = $"[不良]{ManufacturingNumberMaskedTextBox.Text}";

                        cmd.ExecuteNonQuery();
                    }
                }

                // 基板登録テーブルへ追加
                using (SQLiteCommand cmd = con.CreateCommand()) {
                    cmd.CommandText =
                        $"INSERT INTO 'Substrate_Reg_{StrProductName}'" +
                        $"(col_Substrate_Name, col_Substrate_Model, col_Substrate_Num, col_Order_Num, col_Increase, col_Defect, col_Person, col_RegDate, col_Revision, col_Comment)" +
                        $"VALUES" +
                        $"(@col_Substrate_Name, @col_Substrate_Model, @col_Substrate_Num, @col_Order_Num, @col_Increase, @col_Defect, @col_Person, @col_RegDate, @col_Revision, @col_Comment)";

                    // チェックボックスにチェックがない場合はNullを
                    cmd.Parameters.Add("@col_Substrate_Name", DbType.AnsiString).Value = StrSubstrateName;
                    cmd.Parameters.Add("@col_Substrate_Model", DbType.AnsiString).Value = StrSubstrateModel;
                    cmd.Parameters.Add("@col_Substrate_Num", DbType.AnsiString).Value = ManufacturingNumberMaskedTextBox.Text;
                    cmd.Parameters.Add("@col_Order_Num", DbType.AnsiString).Value = OrderNumberTextBox.Text;
                    cmd.Parameters.Add("@col_Increase", DbType.AnsiString).Value = QuantityCheckBox.Checked ? QuantityTextBox.Text : DBNull.Value;
                    cmd.Parameters.Add("@col_Defect", DbType.AnsiString).Value = DefectNumberCheckBox.Checked ? DefectNumberTextBox.Text : DBNull.Value;
                    cmd.Parameters.Add("@col_RegDate", DbType.AnsiString).Value = RegistrationDateMaskedTextBox.Text;
                    cmd.Parameters.Add("@col_Person", DbType.AnsiString).Value = PersonComboBox.Text;
                    cmd.Parameters.Add("@col_Revision", DbType.AnsiString).Value = RevisionTextBox.Text;
                    cmd.Parameters.Add("@col_Comment", DbType.AnsiString).Value = CommentTextBox.Text;

                    cmd.ExecuteNonQuery();
                }
                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        // 印刷処理
        private void PrintBarcode(int PrintFlg) {
            // PrintDocumentオブジェクトの作成
            System.Drawing.Printing.PrintDocument pd = new();

            // PrintPageイベントハンドラの追加
            pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

            LabelSubNumLabelsToPrint = int.Parse(QuantityTextBox.Text);
            LabelSubPageNum = 0;

            switch (PrintFlg) {
                case 1:
                    SubstrateRegistrationPrintDialog.Document = pd;
                    DialogResult r = SubstrateRegistrationPrintDialog.ShowDialog();

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
            }
        }
        private void PrintDocumentPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e) {
            try {
                if (SettingsLabelSub == null || e.Graphics == null) { return; }
                int TxtNumPublish = 0;
                int MaxX = SettingsLabelSub.LabelSubPageSettings.NumLabelsX;
                int MaxY = SettingsLabelSub.LabelSubPageSettings.NumLabelsY;
                float SizeX = (float)SettingsLabelSub.LabelSubPageSettings.SizeX;
                float SizeY = (float)SettingsLabelSub.LabelSubPageSettings.SizeY;
                double OffsetX = SettingsLabelSub.LabelSubPageSettings.OffsetX;
                double OffsetY = SettingsLabelSub.LabelSubPageSettings.OffsetY;
                double IntervalX = SettingsLabelSub.LabelSubPageSettings.IntervalX;
                double IntervalY = SettingsLabelSub.LabelSubPageSettings.IntervalY;
                float PosX, PosY;
                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                Point Offset;
                int StartLine = (int)PrintPostionNumericUpDown.Value - 1;

                if (!SubstrateRegistrationPrintDocument.PrintController.IsPreview) {
                    OffsetX -= e.PageSettings.HardMarginX * 0.254;
                    OffsetY -= e.PageSettings.HardMarginY * 0.254;
                    if (LabelSubPageNum == 0) {
                        Offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (StartLine * (IntervalY + SizeY))));
                    }
                    else {
                        Offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (IntervalY + SizeY))));
                    }
                }
                else {
                    Offset = new Point(0, 0);
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;

                string HeaderString = ConvertHeaderFooterString(SettingsLabelSub.LabelSubPageSettings.HeaderString);
                Point HeaderPos = SettingsLabelSub.LabelSubPageSettings.HeaderPos;
                HeaderPos.Offset(Offset);
                e.Graphics.DrawString(HeaderString, SettingsLabelSub.LabelSubPageSettings.HeaderFooterFont, Brushes.Black, HeaderPos);

                int x, y;
                string s;

                LabelSubNSerial = ManufacturingNumberMaskedTextBox.Text;

                if (LabelSubPageNum >= 1) {
                    StartLine = 0;
                }

                int IntNumLabels = SettingsLabelSub.LabelSubLabelSettings.NumLabels;

                for (y = StartLine; y < MaxY; y++) {
                    for (x = 0; x < MaxX; x++) {
                        s = GenerateCode(LabelSubNSerial);
                        PosX = (float)(OffsetX + (x * (IntervalX + SizeX)));
                        PosY = (float)(OffsetY + (y * (IntervalY + SizeY)));
                        e.Graphics.DrawImage(MakeLabelImage(s, (int)e.Graphics.DpiX, 1), PosX, PosY, SizeX, SizeY);

                        LabelSubNLabel = 0;
                        LabelSubNumLabelsToPrint--;

                        if (LabelSubNumLabelsToPrint <= 0) {
                            IntNumLabels--;
                            if (IntNumLabels <= 0) {
                                e.HasMorePages = false;
                                LabelSubPageNum = 0;
                                LabelSubNumLabelsToPrint = TxtNumPublish;
                                return;
                            }
                            else {
                                LabelSubNumLabelsToPrint += x + 1;
                                break;
                            }
                        }

                        if (x >= MaxX - 1) {
                            IntNumLabels--;
                            if (IntNumLabels <= 0) {
                                IntNumLabels = SettingsLabelSub.LabelSubLabelSettings.NumLabels;
                            }
                            else if (IntNumLabels > 0) {
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
                 .Replace("%N", PersonComboBox.Text)
                 .Replace("%U", PersonComboBox.Text);
            return s;
        }
        private string GenerateCode(string Serial) {
            string SerialCode = Serial.Substring(5, 5);
            string OutputCode = string.Empty;
            string MonthCode = DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("MM");

            MonthCode = MonthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => string.Empty
            };

            if (SettingsLabelSub != null) {
                OutputCode = SettingsLabelSub.LabelSubLabelSettings.Format;
                OutputCode = OutputCode.Replace("%T", StrInitial);
                OutputCode = OutputCode.Replace("%R", RevisionTextBox.Text);
                OutputCode = OutputCode.Replace("%Y", DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("yy"));
                OutputCode = OutputCode.Replace("%MM", DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("MM"));
                OutputCode = OutputCode.Replace("%M", MonthCode[^1..]);
                OutputCode = OutputCode.Replace("%S", SerialCode);
            }
            return OutputCode;
        }
        private Bitmap MakeLabelImage(string Text, int Resolution, int Magnitude) {
            if (SettingsLabelSub is null) { throw new Exception(); }
            decimal SizeX = (decimal)SettingsLabelSub.LabelSubPageSettings.SizeX / 25.4M * Resolution * Magnitude;
            decimal SizeY = (decimal)SettingsLabelSub.LabelSubPageSettings.SizeY / 25.4M * Resolution * Magnitude;

            Bitmap LabelImage = new((int)SizeX, (int)SizeY);
            Graphics g = Graphics.FromImage(LabelImage);

            decimal FontSize = (decimal)SettingsLabelSub.LabelSubLabelSettings.Font.SizeInPoints / 72.0M * Resolution * Magnitude;
            Font fnt = new(SettingsLabelSub.LabelSubLabelSettings.Font.Name, (float)FontSize);

            SizeF stringSize = g.MeasureString(Text, fnt);

            int StringPosX = (int)((LabelImage.Width / 2) - (stringSize.Width / 2));
            int StringPosY = (int)((decimal)SettingsLabelSub.LabelSubLabelSettings.StringPosY / 25.4M * Resolution * Magnitude);

            g.DrawString(Text, fnt, Brushes.Black, StringPosX, StringPosY);

            g.Dispose();

            return LabelImage;
        }

        private void TemplateComment() {
            string templateWord = CommentComboBox.SelectedIndex switch {
                1 => "[Rev.UP]変更点番号:",
                _ => string.Empty
            };
            CommentTextBox.Text = $"{CommentTextBox.Text}{templateWord}";
        }

        private void CheckBoxChecked(object sender, EventArgs e) {
            CheckBox checkBox = (CheckBox)sender;

            switch (checkBox.Name) {
                case "OrderNumberCheckBox":
                    OrderNumberTextBox.Enabled = checkBox.Checked;
                    break;
                case "ManufacturingNumberCheckBox":
                    ManufacturingNumberMaskedTextBox.Enabled = checkBox.Checked;
                    break;
                case "QuantityCheckBox":
                    QuantityTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) DefectNumberCheckBox.Checked = false;
                    break;
                case "DefectNumberCheckBox":
                    DefectNumberTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) QuantityCheckBox.Checked = false;
                    break;
                case "RevisionCheckBox":
                    RevisionTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) MessageBox.Show("変更する場合は理由を記載して下さい。");
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
                    break;
            }
        }

        private void NumericOnly(object sender, KeyPressEventArgs e) {
            // 0～9と、バックスペース以外の時は、イベントをキャンセルする
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b') {
                e.Handled = true;
            }
        }

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
        }

        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            MessageBox.Show($"" +
                $"StrProness1\t\t[{StrProness1}]\r\n" +
                $"StrProness2\t\t[{StrProness2}]\r\n" +
                $"StrProness3\t\t[{StrProness3}]\r\n" +
                $"StrProness4\t\t[{StrProness4}]\r\n" +
                $"StrProness5\t\t[{StrProness5}]\r\n" +
                $"StrProductName\t\t[{StrProductName}]\r\n" +
                $"StrStockName\t\t[{StrStockName}]\r\n" +
                $"StrSubstrateName\t\t[{StrSubstrateName}]\r\n" +
                $"StrSubstrateModel\t\t[{StrSubstrateModel}]\r\n" +
                $"StrInitial\t\t\t[{StrInitial}]\r\n" +
                $"IntRegType\t\t[{IntRegType}]\r\n" +
                $"IntPrintType\t\t[{IntPrintType}]\r\n" +
                $"", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SubstrateRegistrationPrintPreviewDialog_Load(object sender, EventArgs e) {
            ToolStrip tool = (ToolStrip)SubstrateRegistrationPrintPreviewDialog.Controls[1];
            tool.Items[0].Visible = false;
        }
    }
}
