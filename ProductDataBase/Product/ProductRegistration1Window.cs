using System.Data.SQLite;

namespace ProductDatabase {
    public partial class ProductRegistration1Window : Form {

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public string StrProductName { get; set; } = string.Empty;
        public string StrStockName { get; set; } = string.Empty;
        public string StrProductType { get; set; } = string.Empty;
        public string StrProductModel { get; set; } = string.Empty;
        public string StrUseSubstrate { get; set; } = string.Empty;
        public string StrInitial { get; set; } = string.Empty;

        public int IntRegType { get; set; }
        public int IntPrintType { get; set; }
        public int IntCheckBin { get; set; }
        public int IntSerialDigit { get; set; }
        public int IntSerialLastNum;

        public string StrProness1 { get; } = string.Empty;
        public string StrProness2 { get; } = string.Empty;
        public string StrProness3 { get; } = string.Empty;
        public int StrProness4 { get; }
        public string StrProness5 { get; } = string.Empty;

        readonly List<string> checkBoxNames = new() {
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "ExtraCheckBox1",
                    "RevisionCheckBox", "ExtraCheckBox2", "ExtraCheckBox3", "FirstSerialNumberCheckBox", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" };

        public ProductRegistration1Window() => InitializeComponent();
        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(StrFontName, IntFontSize);

                ProductNameLabel2.Text = StrProductName;
                SubstrateModelLabel2.Text = $"{StrProductName} - {StrProductModel}";

                OrderNumberTextBox.Text = StrProness5;
                ManufacturingNumberMaskedTextBox.Text = !string.IsNullOrEmpty(StrProness1) ? StrProness1 : ManufacturingNumberMaskedTextBox.Text;
                QuantityTextBox.Text = StrProness4.ToString();

                FirstSerialNumberTextBox.MaxLength = IntSerialDigit;

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

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using (SQLiteConnection _con = new(MainWindow.GetConnectionString2())) {
                    _con.Open();
                    using SQLiteCommand _cmd = _con.CreateCommand();
                    // テーブル検索SQL - [Product_Name]_stockテーブルの[col_Substrate_Model]列の[col_Stock]の合計を取得
                    _cmd.CommandText = $"SELECT col_Revision FROM 'Product_Reg_{StrProductName}' ORDER BY _rowid_ DESC";
                    object _result = _cmd.ExecuteScalar();
                    RevisionTextBox.Text = _result?.ToString() ?? "";

                    // テーブル検索SQL - [Product_Reg_[Product_Name]]テーブルの最新の[col_Serial_LastNum]を取得
                    _cmd.CommandText = $"SELECT col_Serial_LastNum FROM 'Product_Reg_{StrProductName}' ORDER BY _rowid_ DESC";
                    IntSerialLastNum = Convert.ToInt32(_cmd.ExecuteScalar());
                    FirstSerialNumberTextBox.Text = (IntSerialLastNum + 1).ToString("000");
                }

                // 変数[check_bin]の値に応じてCheckboxにチェックを入れる
                foreach (string _checkBoxName in checkBoxNames) {
                    if (Controls[_checkBoxName] is CheckBox checkBox) {
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
                bool _anyTextBoxEnabled = false;
                bool _allTextBoxesFilled = true;

                foreach (Control control in Controls) {
                    if (control is TextBoxBase textBox && textBox.Enabled) {
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

                DialogResult _result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (_result == DialogResult.Cancel) return;

                switch (IntSerialDigit) {
                    case 3:
                        CheckAndAdjustSerial(999, 1);
                        break;
                    case 4:
                        CheckAndAdjustSerial(9999, 1);
                        break;
                    default:
                        break;
                }

                RegisterButton.Enabled = false;

                using ProductRegistration2Window _window = new();
                _window.StrFontName = StrFontName;
                _window.IntFontSize = IntFontSize;
                _window.StrProductName = StrProductName;
                _window.StrStockName = StrStockName;
                _window.StrProductType = StrProductType;
                _window.IntRegType = IntRegType;
                _window.IntPrintType = IntPrintType;
                _window.IntSerialDigit = IntSerialDigit;
                _window.StrProductModel = StrProductModel;
                _window.IntCheckBin = IntCheckBin;
                _window.StrUseSubstrate = StrUseSubstrate;
                _window.StrInitial = StrInitial;
                _window.StrOrderNumber = OrderNumberTextBox.Text;
                _window.StrProductNumber = ManufacturingNumberMaskedTextBox.Text;
                _window.StrRegDate = RegistrationDateMaskedTextBox.Text;
                _window.StrPerson = PersonComboBox.Text;
                _window.StrRevision = RevisionTextBox.Text;
                _window.StrComment = CommentTextBox.Text;
                _window.IntQuantity = Convert.ToInt32(QuantityTextBox.Text);
                _window.IntSerialFirstNumber = Convert.ToInt32(FirstSerialNumberTextBox.Text);
                _window.ShowDialog(this);
                Close();

                void CheckAndAdjustSerial(int threshold, int resetValue) {

                    int _quantity = Convert.ToInt32(QuantityTextBox.Text);
                    int _firstSerial = Convert.ToInt32(FirstSerialNumberTextBox.Text);
                    if (_quantity + _firstSerial >= threshold) {
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
            string _templateWord = CommentComboBox.SelectedIndex switch {
                1 => "[Rev.UP]変更点番号:",
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
                    if (_checkBox.Checked) ExtraCheckBox1.Checked = false;
                    break;
                case "DefectNumberCheckBox":
                    ExtraTextBox1.Enabled = _checkBox.Checked;
                    if (_checkBox.Checked) QuantityCheckBox.Checked = false;
                    break;
                case "RevisionCheckBox":
                    RevisionTextBox.Enabled = _checkBox.Checked;
                    if (_checkBox.Checked) MessageBox.Show("変更する場合は理由を記載して下さい。");
                    break;
                case "ExtraCheckBox1":
                    ExtraTextBox2.Enabled = _checkBox.Checked;
                    break;
                case "ExtraCheckBox2":
                    ExtraTextBox3.Enabled = _checkBox.Checked;
                    break;
                case "FirstSerialNumberCheckBox":
                    FirstSerialNumberTextBox.Enabled = _checkBox.Checked;
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
                            $"StrProductModel\t\t[{StrProductModel}]\r\n" +
                            $"StrInitial\t\t\t[{StrInitial}]\r\n" +
                            $"IntRegType\t\t[{IntRegType}]\r\n" +
                            $"IntPrintType\t\t[{IntPrintType}]\r\n" +
                            $"IntSerialDigit\t\t[{IntSerialDigit}]" +
                            $"", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

    }
}
