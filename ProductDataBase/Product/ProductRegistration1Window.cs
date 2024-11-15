using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class ProductRegistration1Window : Form {

        public ProductInfomation ProductInfo { get; }

        private int _serialLastNum;

        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "ExtraCheckBox1",
                    "RevisionCheckBox", "ExtraCheckBox2", "ExtraCheckBox3", "FirstSerialNumberCheckBox", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" ];

        public ProductRegistration1Window(ProductInfomation productInfo) {
            InitializeComponent();
            ProductInfo = productInfo;
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);

                ProductNameLabel2.Text = ProductInfo.ProductName;
                SubstrateModelLabel2.Text = $"{ProductInfo.ProductName} - {ProductInfo.ProductModel}";
                ProductTypeLabel2.Text = ProductInfo.ProductType;

                OrderNumberTextBox.Text = ProductInfo.Proness5;
                ManufacturingNumberMaskedTextBox.Text = !string.IsNullOrEmpty(ProductInfo.Proness1) ? ProductInfo.Proness1 : ManufacturingNumberMaskedTextBox.Text;
                QuantityTextBox.Text = (ProductInfo.Proness4 != 0) ? ProductInfo.Proness4.ToString() : string.Empty;

                FirstSerialNumberTextBox.MaxLength = ProductInfo.SerialDigit;

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
                using (SQLiteConnection con = new(GetConnectionInfomation())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        PersonComboBox.Items.Add($"{dr["PersonName"]}");
                    }
                }

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using (SQLiteConnection con = new(GetConnectionRegistration())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - [Product_[ProductName]]テーブルの最新の[Revision]を取得
                    cmd.CommandText = $"""SELECT Revision FROM "Product_{ProductInfo.ProductName}" WHERE RevisionGroup = "{ProductInfo.RevisionGroup}" ORDER BY _rowid_ DESC""";
                    var result = cmd.ExecuteScalar();
                    RevisionTextBox.Text = result?.ToString() ?? "";

                    // テーブル検索SQL - [Product_[ProductName]]テーブルの最新の[SerialLastNumber]を取得
                    cmd.CommandText = $"""SELECT SerialLastNumber FROM "Product_{ProductInfo.ProductName}" ORDER BY _rowid_ DESC""";
                    FirstSerialNumberTextBox.Text = int.TryParse(cmd.ExecuteScalar()?.ToString(), out _serialLastNum)
                        ? (_serialLastNum + 1).ToString("000")
                        : throw new Exception("シリアル番号の取得に失敗しました。");
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // 登録チェック
        private void RegisterCheck() {
            try {
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

                var result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) {
                    return;
                }

                switch (ProductInfo.SerialDigit) {
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

                ProductInfo.OrderNumber = OrderNumberTextBox.Text;
                ProductInfo.ProductNumber = ManufacturingNumberMaskedTextBox.Text;
                ProductInfo.RegDate = RegistrationDateMaskedTextBox.Text;
                ProductInfo.Person = PersonComboBox.Text;
                ProductInfo.Revision = RevisionTextBox.Text;
                ProductInfo.Comment = CommentTextBox.Text;
                ProductInfo.Quantity = Convert.ToInt32(QuantityTextBox.Text ?? throw new Exception("QuantityTextBox.Text is null"));
                ProductInfo.SerialFirstNumber = Convert.ToInt32(FirstSerialNumberTextBox.Text ?? throw new Exception("FirstSerialNumberTextBox.Text is null"));
                using ProductRegistration2Window window = new();
                window.ProductInfo = ProductInfo;
                window.Closed += (s, e) => this.Close();
                window.ShowDialog(this);

                void CheckAndAdjustSerial(int threshold, int resetValue) {

                    var quantity = Convert.ToInt32(QuantityTextBox.Text ?? throw new Exception());
                    var firstSerial = Convert.ToInt32(FirstSerialNumberTextBox.Text ?? throw new Exception());
                    if (quantity + firstSerial >= threshold) {
                        MessageBox.Show($"シリアルが{threshold}を超えるので{resetValue.ToString().PadLeft(ProductInfo.SerialDigit, '0')}から開始します。");
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
                        ExtraCheckBox1.Checked = false;
                    }

                    break;
                case "DefectNumberCheckBox":
                    ExtraTextBox1.Enabled = checkBox.Checked;
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

        private void ProductRegistration1Window_Load(object sender, EventArgs e) { LoadEvents(); }
        private void RegisterButton_Click(object sender, EventArgs e) { RegisterCheck(); }
        private void TemplateButton_Click(object sender, EventArgs e) { TemplateComment(); }
        private void NumberCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBoxChecked(sender, e); }
        private void QuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void RegistrationDateMaskedTextBox_TypeValidationCompleted(object sender, TypeValidationEventArgs e) { RegistrationDateCheck(sender, e); }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            var message = string.Join(Environment.NewLine,
                $"StrProness1\t\t[{ProductInfo.Proness1}]",
                $"StrProness2\t\t[{ProductInfo.Proness2}]",
                $"StrProness3\t\t[{ProductInfo.Proness3}]",
                $"StrProness4\t\t[{ProductInfo.Proness4}]",
                $"StrProness5\t\t[{ProductInfo.Proness5}]",
                $"StrProductName\t\t[{ProductInfo.ProductName}]",
                $"StrStockName\t\t[{ProductInfo.StockName}]",
                $"StrProductModel\t\t[{ProductInfo.ProductModel}]",
                $"StrInitial\t\t\t[{ProductInfo.Initial}]",
                $"IntRegType\t\t[{ProductInfo.RegType}]",
                $"IntPrintType\t\t[{ProductInfo.PrintType}]",
                $"IntSerialDigit\t\t[{ProductInfo.SerialDigit}]"
            );
            MessageBox.Show(message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
