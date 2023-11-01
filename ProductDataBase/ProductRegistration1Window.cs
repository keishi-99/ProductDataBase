using System.Data.SQLite;

namespace ProductDataBase {
    public partial class ProductRegistration1Window : Form {

        public string? StrFontName { get; set; }
        public int IntFontSize { get; set; }

        public string? StrProductName { get; set; }
        public string? StrStockName { get; set; }
        public string? StrProductType { get; set; }
        public string? StrProductModel { get; set; }
        public string? StrUseSubstrate { get; set; }
        public string? StrInitial { get; set; }
        public int IntRegType { get; set; }
        public int IntPrintType { get; set; }
        public int IntCheckBin { get; set; }
        public int IntSerialDigit { get; set; }
        public int IntSerialLastNum;

        public string? StrProness1 { get; }
        public string? StrProness2 { get; }
        public string? StrProness3 { get; }
        public int StrProness4 { get; }
        public string? StrProness5 { get; }

        readonly List<string> checkBoxNames = new() {
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "ExtraCheckBox1",
                    "RevisionCheckBox", "ExtraCheckBox2", "ExtraCheckBox3", "FirstSerialNumberCheckBox", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" };

        public ProductRegistration1Window(MainWindow mainWindow) {
            StrFontName = mainWindow.StrFontName;
            IntFontSize = mainWindow.IntFontSize;

            StrProductName = mainWindow.StrProductName;
            StrStockName = mainWindow.StrStockName;
            StrProductType = mainWindow.StrSubstrateName;
            StrProductModel = mainWindow.StrSubstrateModel;
            StrUseSubstrate = mainWindow.StrUseSubstrate;
            StrInitial = mainWindow.StrInitial;
            IntRegType = mainWindow.IntRegType;
            IntPrintType = mainWindow.IntPrintType;
            IntCheckBin = mainWindow.IntCheckBin;
            IntSerialDigit = mainWindow.IntSerialDigit;

            StrProness1 = mainWindow.StrProness1;
            StrProness2 = mainWindow.StrProness2;
            StrProness3 = mainWindow.StrProness3;
            StrProness4 = mainWindow.StrProness4;
            StrProness5 = mainWindow.StrProness5;
        }

        public ProductRegistration1Window() {
            InitializeComponent();
        }
        // ロードイベント
        private void LoadEvents() {
            try {
                if (StrFontName != null) { Font = new Font(StrFontName, IntFontSize); }

                ProductNameLabel2.Text = StrProductName;
                SubstrateModelLabel2.Text = $"{StrProductName} - {StrProductModel}";

                if (StrProness5 != null) OrderNumberTextBox.Text = StrProness5;
                if (StrProness1 != null) ManufacturingNumberMaskedTextBox.Text = StrProness1;
                if (StrProness4 != 0) QuantityTextBox.Text = StrProness4.ToString();

                FirstSerialNumberTextBox.MaxLength = IntSerialDigit;

                RegisterButton.Enabled = true;

                // TextBoxへ今日の年月日を入力
                DateTime dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = dtNow.ToShortDateString();

                // DB1へ接続し担当者取得
                using (SQLiteConnection con = new(MainWindow.GetConnectionString1())) {
                    con.Open();
                    using SQLiteCommand cmd = con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using SQLiteDataReader dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        PersonComboBox.Items.Add($"{dr["col_Person_Name"]}");
                    }
                }

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using (SQLiteConnection con = new(MainWindow.GetConnectionString2())) {
                    con.Open();
                    using SQLiteCommand cmd = con.CreateCommand();
                    // テーブル検索SQL - [Product_Name]_stockテーブルの[col_Substrate_Model]列の[col_Stock]の合計を取得
                    cmd.CommandText = $"SELECT col_Revision FROM 'Product_Reg_{StrProductName}' ORDER BY _rowid_ DESC";
                    object result = cmd.ExecuteScalar();
                    RevisionTextBox.Text = result?.ToString() ?? "";

                    // テーブル検索SQL - [Product_Reg_[Product_Name]]テーブルの最新の[col_Serial_LastNum]を取得
                    cmd.CommandText = $"SELECT col_Serial_LastNum FROM 'Product_Reg_{StrProductName}' ORDER BY _rowid_ DESC";
                    IntSerialLastNum = Convert.ToInt32(cmd.ExecuteScalar());
                    FirstSerialNumberTextBox.Text = (IntSerialLastNum + 1).ToString("000");
                }

                // 変数[check_bin]の値に応じてCheckboxにチェックを入れる
                foreach (string checkBoxName in checkBoxNames) {
                    if (Controls[checkBoxName] is CheckBox checkBox) {
                        checkBox.Checked = (IntCheckBin & 0x1) == 1;
                        IntCheckBin >>= 1;
                    }
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // クロージングイベント
        private void ClosingEvents() { Dispose(); }
        // 登録チェック
        private void RegisterCheck() {
            try {
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

                DialogResult result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) return;

                int quantity = Convert.ToInt32(QuantityTextBox.Text);
                int firstSerial = Convert.ToInt32(FirstSerialNumberTextBox.Text);

                switch (IntSerialDigit) {
                    case 3:
                        CheckAndAdjustSerial(999, 1);
                        break;
                    case 4:
                        CheckAndAdjustSerial(9999, 1);
                        break;
                }

                RegisterButton.Enabled = false;

                ProductRegistration2Window productRegistration2Window = new() {
                    StrProductName = StrProductName,
                    StrStockName = StrStockName,
                    StrProductType = StrProductType,
                    IntRegType = IntRegType,
                    IntPrintType = IntPrintType,
                    IntSerialDigit = IntSerialDigit,
                    StrProductModel = StrProductModel,
                    IntCheckBin = IntCheckBin,
                    StrUseSubstrate = StrUseSubstrate,
                    StrInitial = StrInitial,
                    StrOrderNumber = OrderNumberTextBox.Text,
                    StrProductNumber = ManufacturingNumberMaskedTextBox.Text,
                    StrRegDate = RegistrationDateMaskedTextBox.Text,
                    StrPerson = PersonComboBox.Text,
                    StrRevision = RevisionTextBox.Text,
                    StrComment = CommentTextBox.Text,
                    IntQuantity = Convert.ToInt32(QuantityTextBox.Text),
                    IntSerialFirstNumber = Convert.ToInt32(FirstSerialNumberTextBox.Text)
                };
                productRegistration2Window.ShowDialog(this);


                void CheckAndAdjustSerial(int threshold, int resetValue) {
                    if (quantity + firstSerial >= threshold) {
                        MessageBox.Show($"シリアルが{threshold}を超えるので{resetValue.ToString().PadLeft(IntSerialDigit, '0')}から開始します。");
                        FirstSerialNumberTextBox.Text = resetValue.ToString();
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                RegisterButton.Enabled = true;
            }
        }
        // コメント用テンプレート
        private void TemplateComment() {
            string templateWord = CommentComboBox.SelectedIndex switch {
                1 => "[Rev.UP]変更点番号:",
                _ => string.Empty
            };
            CommentTextBox.Text = $"{CommentTextBox.Text}{templateWord}";
        }
        // チェックボックスイベント
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
                    if (checkBox.Checked) ExtraCheckBox1.Checked = false;
                    break;
                case "DefectNumberCheckBox":
                    ExtraTextBox1.Enabled = checkBox.Checked;
                    if (checkBox.Checked) QuantityCheckBox.Checked = false;
                    break;
                case "RevisionCheckBox":
                    RevisionTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) MessageBox.Show("変更する場合は理由を記載して下さい。");
                    break;
                case "ExtraCheckBox1":
                    ExtraTextBox2.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox2":
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

        private void ProductRegistration1Window_Load(object sender, EventArgs e) { LoadEvents(); }
        private void ProductRegistration1Window_FormClosing(object sender, FormClosingEventArgs e) { ClosingEvents(); }
        private void RegisterButton_Click(object sender, EventArgs e) { RegisterCheck(); }
        private void TemplateButton_Click(object sender, EventArgs e) { TemplateComment(); }
        private void NumberCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBoxChecked(sender, e); }
        private void QuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void RegistrationDateMaskedTextBox_TypeValidationCompleted(object sender, TypeValidationEventArgs e) { RegistrationDateCheck(sender, e); }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            MessageBox.Show($"" +
                            $"StrProness1\t\t[{StrProness1}]\r\n" +
                            $"StrProness2\t\t[{StrProness2}]\r\n" +
                            $"StrProness3\t\t[{StrProness3}]\r\n" +
                            $"StrProness4\t\t[{StrProness4}]\r\n" +
                            $"StrProness5\t\t[{StrProness5}]\r\n" +
                            $"StrProductName\t\t[{StrProductName}]\r\n" +
                            $"StrStockName\t\t[{StrStockName}]\r\n" +
                            $"StrProductName\t\t[{StrProductName}]\r\n" +
                            $"StrProductModel\t\t[{StrProductModel}]\r\n" +
                            $"StrInitial\t\t\t[{StrInitial}]\r\n" +
                            $"IntRegType\t\t[{IntRegType}]\r\n" +
                            $"IntPrintType\t\t[{IntPrintType}]\r\n" +
                            $"IntSerialDigit\t\t[{IntSerialDigit}]" +
                            $"", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

    }
}
