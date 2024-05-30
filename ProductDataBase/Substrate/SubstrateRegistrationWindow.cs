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

        readonly List<string> checkBoxNames = new() {
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "DefectNumberCheckBox",
                    "RevisionCheckBox", "ExtraCheckBox1", "ExtraCheckBox2", "ExtraCheckBox3", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" };

        public SubstrateRegistrationWindow() => InitializeComponent();
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
                DateTime _dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = _dtNow.ToShortDateString();

                // DB1へ接続し担当者取得
                using (SQLiteConnection _con = new(MainWindow.GetConnectionString1())) {
                    _con.Open();
                    using SQLiteCommand _cmd = _con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    _cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using SQLiteDataReader _dr = _cmd.ExecuteReader();
                    while (_dr.Read()) {
                        PersonComboBox.Items.Add($"{_dr["col_Person_Name"]}");
                    }
                }

                // 在庫管理する基盤はDB2へ接続し対象製品の在庫取得
                if (IntRegType != 0) {
                    using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
                    _con.Open();
                    using SQLiteCommand _cmd = _con.CreateCommand();
                    // テーブル検索SQL - [Product_Name]_stockテーブルの[col_Substrate_Model]列の[col_Stock]の合計を取得
                    _cmd.CommandText = $@"SELECT total(col_stock) FROM Stock_{StrStockName} WHERE col_Substrate_Model = @col_Substrate_Model";
                    _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = StrSubstrateModel;
                    StockLabel2.Text = _cmd.ExecuteScalar().ToString();

                    // テーブル検索SQL - [Substrate_Reg_[Product_Name]]テーブルの最新の[col_Revison]を取得
                    _cmd.CommandText = $@"SELECT col_Revision FROM Substrate_Reg_{StrStockName} WHERE col_Substrate_Model = @col_Substrate_Model AND col_Revision IS NOT NULL ORDER BY _rowid_ DESC";
                    _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = StrSubstrateModel;
                    object _result = _cmd.ExecuteScalar();
                    RevisionTextBox.Text = _result?.ToString() ?? "";
                }
                else { StockLabel2.Text = "---"; }

                // 変数[check_bin]の値に応じてCheckboxにチェックを入れる
                foreach (string _checkBoxName in checkBoxNames) {
                    if (Controls[_checkBoxName] is CheckBox checkBox) {
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
        private void LoadSettings(string StrSettingFilePath) {
            try {
                StreamReader _sr = new(StrSettingFilePath, new System.Text.UTF8Encoding(false));
                System.Xml.Serialization.XmlSerializer _serializer = new(typeof(CSettingsLabelSub));
                if (_serializer.Deserialize(_sr) is CSettingsLabelSub _result) { SettingsLabelSub = _result; }
                _sr.Close();
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
                bool _anyTextBoxEnabled = false;
                bool _allTextBoxesFilled = true;

                foreach (Control _control in Controls) {
                    if (_control is TextBoxBase textBox && textBox.Enabled) {
                        _anyTextBoxEnabled = true;
                        if (string.IsNullOrWhiteSpace(textBox.Text)) {
                            _allTextBoxesFilled = false;
                            break;
                        }
                    }
                }
                if (!_anyTextBoxEnabled) { throw new Exception("何も入力されていません"); }
                if (!_allTextBoxesFilled) { throw new Exception("空欄があります。"); }

                if (ManufacturingNumberCheckBox.Checked && ManufacturingNumberMaskedTextBox.Text.Length != 15) { throw new Exception("製番を10桁+4桁で入力して下さい。"); }

                if (QuantityCheckBox.Checked && int.Parse(QuantityTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

                if (DefectNumberCheckBox.Checked && int.Parse(DefectNumberTextBox.Text) <= 0) { throw new Exception("1台以上入力して下さい。"); }

                if (!QuantityCheckBox.Checked && !DefectNumberCheckBox.Checked) { throw new Exception("数量か不良数を入力してください。"); }

                if (!DefectNumberCheckBox.Checked && string.IsNullOrEmpty(PrintPostionNumericUpDown.Text)) { PrintPostionNumericUpDown.Text = "1"; }

                DialogResult _result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (_result == DialogResult.Cancel) return;

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
                using SQLiteConnection _con = new(MainWindow.GetConnectionString2());
                _con.Open();

                // 製番が新規かチェック
                if (IntRegType != 0) {
                    string _substrateName = string.Empty;
                    using (SQLiteCommand _cmd = _con.CreateCommand()) {
                        _cmd.CommandText = $@"SELECT * FROM Stock_{StrStockName} WHERE col_Substrate_Num = @col_Substrate_Num ORDER BY _rowid_ DESC LIMIT 1";
                        _cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                        using SQLiteDataReader _dr = _cmd.ExecuteReader();
                        while (_dr.Read()) {
                            _substrateName = $"{_dr["col_Substrate_Name"]}";
                        }
                    }

                    if (_substrateName != string.Empty) {
                        if (StrSubstrateName == _substrateName) {
                            DialogResult _result = MessageBox.Show($"[{ManufacturingNumberMaskedTextBox.Text}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                            if (_result == DialogResult.No) { return false; }
                        }
                        else { throw new Exception($"[{ManufacturingNumberMaskedTextBox.Text}]は[{_substrateName}]として在庫があります。確認してください。"); }
                    }
                }

                // 在庫管理する基盤のみ
                if (IntRegType != 0) {
                    //在庫追加
                    if (QuantityCheckBox.Checked && !DefectNumberCheckBox.Checked) {
                        // 基板在庫テーブルへ追加_製番が一致した行を更新する、製番がない場合は新規追加
                        using SQLiteCommand _cmd = _con.CreateCommand();
                        _cmd.CommandText =
                            $@"
                            INSERT INTO Stock_{StrStockName} (
                                col_Flg, col_Substrate_Name, col_Substrate_Model, col_Substrate_Num, col_Order_Num, col_Stock)
                            VALUES (
                                @col_Flg, @col_Substrate_Name, @col_Substrate_Model, @col_Substrate_Num, @col_Order_Num, @col_Stock)
                                on conflict(col_Substrate_Num)
                            DO UPDATE
                                set col_Flg = 1, col_Stock = col_Stock + excluded.col_Stock
                            ";

                        _cmd.Parameters.Add("@col_Flg", DbType.String).Value = 1;
                        _cmd.Parameters.Add("@col_Substrate_Name", DbType.String).Value = StrSubstrateName;
                        _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = StrSubstrateModel;
                        _cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                        _cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = OrderNumberTextBox.Text;
                        _cmd.Parameters.Add("@col_Stock", DbType.String).Value = QuantityTextBox.Text;

                        _cmd.ExecuteNonQuery();
                    }
                    //不良処理
                    else if (!QuantityCheckBox.Checked && DefectNumberCheckBox.Checked) {
                        using SQLiteCommand _cmd = _con.CreateCommand();
                        _cmd.CommandText = $@"SELECT col_Stock FROM Stock_{StrStockName} WHERE col_Substrate_Num = @col_Substrate_Num";
                        _cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                        int _intStock = Convert.ToInt32(_cmd.ExecuteScalar());

                        if (_intStock == 0) { throw new Exception("該当する製番の在庫がありません。"); }

                        int _intStockFlg;
                        if ((_intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text)) > 0) { _intStockFlg = 1; }
                        else if ((_intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text)) == 0) { _intStockFlg = 0; }
                        else { throw new Exception("不良数が在庫より多く入力されています。"); }

                        _cmd.CommandText =
                            $@"UPDATE Stock_{StrStockName} SET
                                col_Flg = @col_Flg,
                                col_Stock = @col_Stock,
                                col_History = ifnull(col_History, '') || @col_History
                            WHERE
                                col_Substrate_Num = @col_Substrate_Num
                            ";
                        _cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;

                        _cmd.Parameters.Add("@col_Flg", DbType.String).Value = _intStockFlg;
                        _cmd.Parameters.Add("@col_Stock", DbType.String).Value = _intStock - Convert.ToInt32(ManufacturingNumberMaskedTextBox.Text);
                        _cmd.Parameters.Add("@col_History", DbType.String).Value = $"[不良]{ManufacturingNumberMaskedTextBox.Text}";

                        _cmd.ExecuteNonQuery();
                    }
                }

                // 基板登録テーブルへ追加
                using (SQLiteCommand _cmd = _con.CreateCommand()) {
                    _cmd.CommandText =
                        $@"
                        INSERT INTO Substrate_Reg_{StrProductName}
                            (col_Substrate_Name,
                            col_Substrate_Model,
                            col_Substrate_Num,
                            col_Order_Num,
                            col_Increase,
                            col_Defect,
                            col_Person,
                            col_RegDate,
                            col_Revision,
                            col_Comment)
                        VALUES
                            (@col_Substrate_Name,
                            @col_Substrate_Model,
                            @col_Substrate_Num,
                            @col_Order_Num,
                            @col_Increase,
                            @col_Defect,
                            @col_Person,
                            @col_RegDate,
                            @col_Revision,
                            @col_Comment)
                        ";

                    // チェックボックスにチェックがない場合はNullを
                    _cmd.Parameters.Add("@col_Substrate_Name", DbType.String).Value = StrSubstrateName;
                    _cmd.Parameters.Add("@col_Substrate_Model", DbType.String).Value = StrSubstrateModel;
                    _cmd.Parameters.Add("@col_Substrate_Num", DbType.String).Value = ManufacturingNumberMaskedTextBox.Text;
                    _cmd.Parameters.Add("@col_Order_Num", DbType.String).Value = OrderNumberTextBox.Text;
                    _cmd.Parameters.Add("@col_Increase", DbType.String).Value = QuantityCheckBox.Checked ? QuantityTextBox.Text : DBNull.Value;
                    _cmd.Parameters.Add("@col_Defect", DbType.String).Value = DefectNumberCheckBox.Checked ? DefectNumberTextBox.Text : DBNull.Value;
                    _cmd.Parameters.Add("@col_RegDate", DbType.String).Value = RegistrationDateMaskedTextBox.Text;
                    _cmd.Parameters.Add("@col_Person", DbType.String).Value = PersonComboBox.Text;
                    _cmd.Parameters.Add("@col_Revision", DbType.String).Value = RevisionTextBox.Text;
                    _cmd.Parameters.Add("@col_Comment", DbType.String).Value = CommentTextBox.Text;

                    _cmd.ExecuteNonQuery();
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
            using System.Drawing.Printing.PrintDocument _pd = new();

            //// PrintPageイベントハンドラの追加
            _pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocumentPrintPage);

            LabelSubNumLabelsToPrint = int.Parse(QuantityTextBox.Text);
            LabelSubPageNum = 0;

            switch (PrintFlg) {
                case 1:
                    SubstrateRegistrationPrintDialog.Document = _pd;
                    DialogResult _r = SubstrateRegistrationPrintDialog.ShowDialog();

                    if (_r == DialogResult.OK) {
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
                    SubstrateRegistrationPrintPreviewDialog.Document = _pd;
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
                int _startLine = (int)PrintPostionNumericUpDown.Value - 1;
                float _sizeY = (float)SettingsLabelSub.LabelSubPageSettings.SizeY;
                double _offsetX = SettingsLabelSub.LabelSubPageSettings.OffsetX;
                double _offsetY = SettingsLabelSub.LabelSubPageSettings.OffsetY;
                double _intervalY = SettingsLabelSub.LabelSubPageSettings.IntervalY;
                Point _offset;
                if (!SubstrateRegistrationPrintDocument.PrintController.IsPreview) {
                    _offsetX -= e.PageSettings.HardMarginX * 0.254;
                    _offsetY -= e.PageSettings.HardMarginY * 0.254;
                    if (LabelSubPageNum == 0) {
                        _offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (_startLine * (_intervalY + _sizeY))));
                    }
                    else {
                        _offset = new Point((int)(e.PageSettings.HardMarginX * -0.254), (int)((e.PageSettings.HardMarginY * -0.254) + (0 * (_intervalY + _sizeY))));
                    }
                }
                else {
                    _offset = new Point(0, 0);
                }

                e.PageSettings.Margins.Left = 0;
                e.PageSettings.Margins.Top = 0;
                Point _headerPos = SettingsLabelSub.LabelSubPageSettings.HeaderPos;
                _headerPos.Offset(_offset);

                string _headerString = ConvertHeaderFooterString(SettingsLabelSub.LabelSubPageSettings.HeaderString);
                e.Graphics.DrawString(_headerString, SettingsLabelSub.LabelSubPageSettings.HeaderFooterFont, Brushes.Black, _headerPos);
                LabelSubNSerial = ManufacturingNumberMaskedTextBox.Text;

                if (LabelSubPageNum >= 1) {
                    _startLine = 0;
                }

                int _maxY = SettingsLabelSub.LabelSubPageSettings.NumLabelsY;
                int _y;
                int _intNumLabels = SettingsLabelSub.LabelSubLabelSettings.NumLabels;
                for (_y = _startLine; _y < _maxY; _y++) {

                    int _maxX = SettingsLabelSub.LabelSubPageSettings.NumLabelsX;

                    int _x;
                    for (_x = 0; _x < _maxX; _x++) {
                        string _s = GenerateCode(LabelSubNSerial);

                        float _sizeX = (float)SettingsLabelSub.LabelSubPageSettings.SizeX;

                        double _intervalX = SettingsLabelSub.LabelSubPageSettings.IntervalX;
                        float _posX = (float)(_offsetX + (_x * (_intervalX + _sizeX)));
                        float _posY = (float)(_offsetY + (_y * (_intervalY + _sizeY)));
                        e.Graphics.DrawImage(MakeLabelImage(_s, (int)e.Graphics.DpiX, 1), _posX, _posY, _sizeX, _sizeY);

                        LabelSubNLabel = 0;
                        LabelSubNumLabelsToPrint--;

                        if (LabelSubNumLabelsToPrint <= 0) {
                            _intNumLabels--;
                            if (_intNumLabels <= 0) {
                                e.HasMorePages = false;
                                LabelSubPageNum = 0;
                                int _txtNumPublish = 0;
                                LabelSubNumLabelsToPrint = _txtNumPublish;
                                return;
                            }
                            else {
                                LabelSubNumLabelsToPrint += _x + 1;
                                break;
                            }
                        }

                        if (_x >= _maxX - 1) {
                            _intNumLabels--;
                            if (_intNumLabels <= 0) {
                                _intNumLabels = SettingsLabelSub.LabelSubLabelSettings.NumLabels;
                            }
                            else if (_intNumLabels > 0) {
                                LabelSubNumLabelsToPrint += _x + 1;
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
        private string GenerateCode(string Serial) {
            string _monthCode = DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("MM");

            _monthCode = _monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => _monthCode
            };

            string _outputCode = SettingsLabelSub.LabelSubLabelSettings.Format;
            string _serialCode = Serial.Substring(5, 5);
            _outputCode = _outputCode.Replace("%T", StrInitial)
                                    .Replace("%R", RevisionTextBox.Text)
                                    .Replace("%Y", DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(RegistrationDateMaskedTextBox.Text).ToString("MM"))
                                    .Replace("%M", string.IsNullOrEmpty(_monthCode) ? string.Empty : _monthCode[^1..])
                                    .Replace("%S", _serialCode);

            return _outputCode;
        }
        private Bitmap MakeLabelImage(string Text, int Resolution, int Magnitude) {
            if (SettingsLabelSub is null) { throw new Exception(); }
            decimal _sizeX = (decimal)SettingsLabelSub.LabelSubPageSettings.SizeX / 25.4M * Resolution * Magnitude;
            decimal _sizeY = (decimal)SettingsLabelSub.LabelSubPageSettings.SizeY / 25.4M * Resolution * Magnitude;

            Bitmap _labelImage = new((int)_sizeX, (int)_sizeY);
            using (Graphics _g = Graphics.FromImage(_labelImage)) {
                decimal _fontSize = (decimal)SettingsLabelSub.LabelSubLabelSettings.Font.SizeInPoints / 72.0M * Resolution * Magnitude;
                using (Font _fnt = new(SettingsLabelSub.LabelSubLabelSettings.Font.Name, (float)_fontSize)) {
                    SizeF _stringSize = _g.MeasureString(Text, _fnt);

                    int _stringPosX = (int)((_labelImage.Width / 2) - (_stringSize.Width / 2));
                    int _stringPosY = (int)((decimal)SettingsLabelSub.LabelSubLabelSettings.StringPosY / 25.4M * Resolution * Magnitude);

                    _g.DrawString(Text, _fnt, Brushes.Black, _stringPosX, _stringPosY);
                }

                _g.Dispose();
            }

            return _labelImage;
        }
        // コメント用テンプレート
        private void TemplateComment() {
            string _templateWord = CommentComboBox.SelectedIndex switch {
                0 => "[Rev.UP]変更点番号:",
                _ => string.Empty
            };
            CommentTextBox.Text = $"{CommentTextBox.Text}{_templateWord}";
        }
        // チェックボックスイベント
        private void CheckBoxChecked(object sender, EventArgs e) {
            CheckBox _checkBox = (CheckBox)sender;

            switch (_checkBox.Name) {
                case "OrderNumberCheckBox":
                    OrderNumberTextBox.Enabled = _checkBox.Checked;
                    break;
                case "ManufacturingNumberCheckBox":
                    ManufacturingNumberMaskedTextBox.Enabled = _checkBox.Checked;
                    break;
                case "QuantityCheckBox":
                    QuantityTextBox.Enabled = _checkBox.Checked;
                    if (_checkBox.Checked) DefectNumberCheckBox.Checked = false;
                    break;
                case "DefectNumberCheckBox":
                    DefectNumberTextBox.Enabled = _checkBox.Checked;
                    if (_checkBox.Checked) QuantityCheckBox.Checked = false;
                    break;
                case "RevisionCheckBox":
                    RevisionTextBox.Enabled = _checkBox.Checked;
                    if (_checkBox.Checked) MessageBox.Show("変更する場合は理由を記載して下さい。");
                    break;
                case "ExtraCheckBox1":
                    ExtraTextBox1.Enabled = _checkBox.Checked;
                    break;
                case "ExtraCheckBox2":
                    ExtraTextBox2.Enabled = _checkBox.Checked;
                    break;
                case "ExtraCheckBox3":
                    ExtraTextBox3.Enabled = _checkBox.Checked;
                    break;
                case "RegistrationDateCheckBox":
                    RegistrationDateMaskedTextBox.Enabled = _checkBox.Checked;
                    break;
                case "PersonCheckBox":
                    PersonComboBox.Enabled = _checkBox.Checked;
                    break;
                case "ExtraCheckBox4":
                    ExtraTextBox4.Enabled = _checkBox.Checked;
                    break;
                case "ExtraCheckBox5":
                    ExtraTextBox5.Enabled = _checkBox.Checked;
                    break;
                case "ExtraCheckBox6":
                    ExtraTextBox6.Enabled = _checkBox.Checked;
                    break;
                case "CommentCheckBox":
                    CommentTextBox.Enabled = _checkBox.Checked;
                    CommentComboBox.Enabled = _checkBox.Checked;
                    TemplateButton.Enabled = _checkBox.Checked;
                    break;
                case "PrintOnlyCheckBox":
                    PrintButton.Enabled = _checkBox.Checked;
                    if (PrintOnlyCheckBox.Checked) { MessageBox.Show($"登録せずに印刷だけ行います。", "注意", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                    break;
                default:
                    break;
            }
        }
        // 入力数値のみ
        private void NumericOnly(object sender, KeyPressEventArgs e) {
            // 0～9と、バックスペース以外の時は、イベントをキャンセルする
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b') {
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
            SubstratePrintSetting _ls = new();
            _ls.ShowDialog(this);

            try {
                StreamWriter? _sw = null;

                _sw = new StreamWriter(StrSettingFilePath, false, new System.Text.UTF8Encoding(false));
                System.Xml.Serialization.XmlSerializer _serializer = new(typeof(CSettingsLabelSub));
                _serializer.Serialize(_sw, SettingsLabelSub);
                _sw?.Close();

                LoadSettings(StrSettingFilePath);
            } catch (Exception ex) {
                MessageBox.Show($"設定ファイルの保存に失敗しました。{Environment.NewLine}{ex.Message}");
            } finally {
            }
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
            ToolStrip _tool = (ToolStrip)SubstrateRegistrationPrintPreviewDialog.Controls[1];
            _tool.Items[0].Visible = false;
        }
    }
}
