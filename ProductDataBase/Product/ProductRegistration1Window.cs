using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class ProductRegistration1Window : Form {

        public ProductInformation ProductInfo { get; }

        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "ExtraCheckBox1",
                    "RevisionCheckBox", "ExtraCheckBox2", "ExtraCheckBox3", "FirstSerialNumberCheckBox", "RegistrationDateCheckBox",
                    "PersonCheckBox", "ExtraCheckBox4", "ExtraCheckBox5", "ExtraCheckBox6", "CommentCheckBox" ];
        // プロパティ設定
        private bool IsSerialGeneration => ProductInfo.RegType is 1 or 2 or 3;

        public ProductRegistration1Window(ProductInformation productInfo) {
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

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. ProductInfo.PersonList]);

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using (SQLiteConnection con = new(GetConnectionRegistration())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - [[ProductName]_Product]テーブルの最新の[Revision]を取得
                    cmd.CommandText = $"""SELECT Revision FROM "{ProductInfo.CategoryName}_Product" WHERE ProductName = @ProductName AND RevisionGroup = @RevisionGroup ORDER BY "ID" DESC""";
                    cmd.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;
                    cmd.Parameters.Add("@RevisionGroup", DbType.String).Value = ProductInfo.RevisionGroup;
                    var result = cmd.ExecuteScalar();
                    RevisionTextBox.Text = result?.ToString() ?? "";

                    // シリアル番号の最後を取得する共通メソッド
                    int GetLastSerialNumber(SQLiteCommand cmd) {
                        cmd.CommandText = $"""SELECT SerialLastNumber FROM "{ProductInfo.CategoryName}_Product" WHERE ProductName = @ProductName AND SerialLastNumber NOT NULL ORDER BY "ID" DESC""";
                        cmd.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;
                        return int.TryParse(cmd.ExecuteScalar()?.ToString(), out var serialLastNum)
                            ? serialLastNum
                            : throw new Exception("シリアル番号の取得に失敗しました。");
                    }

                    // 登録日を取得して年月を返す共通メソッド
                    string GetLastRegistrationYearMonth(SQLiteCommand cmd) {
                        cmd.CommandText = $"""SELECT RegDate FROM "{ProductInfo.CategoryName}_Product" WHERE ProductName = @ProductName ORDER BY "ID" DESC""";
                        var lastRegDate = cmd.ExecuteScalar()?.ToString() ?? string.Empty;
                        var splitRegDate = lastRegDate.Split("/");
                        return splitRegDate.Length >= 2 ? $"{splitRegDate[0]}/{splitRegDate[1]}" : throw new Exception("登録日の取得に失敗しました。");
                    }
                    // メイン処理
                    switch (ProductInfo.SerialDigit) {
                        case 93:
                        case 94: {
                                // 登録日の年月と現在の年月を比較
                                var lastYearMonth = GetLastRegistrationYearMonth(cmd);
                                var currentYearMonth = string.Join("/", ProductInfo.RegDate.Split("/").Take(2)); // 現在の年月

                                if (lastYearMonth != currentYearMonth) {
                                    // 月が異なる場合はシリアル番号をリセット
                                    FirstSerialNumberTextBox.Text = "1";
                                }
                                else {
                                    // 月が同じ場合は最後のシリアル番号を取得して +1
                                    FirstSerialNumberTextBox.Text = (GetLastSerialNumber(cmd) + 1).ToString();
                                }
                                break;
                            }
                        default: {
                                // 他の場合は単純に最後のシリアル番号を取得して +1
                                if (ProductInfo.RegType == 0) { return; }
                                FirstSerialNumberTextBox.Text = (GetLastSerialNumber(cmd) + 1).ToString();
                                break;
                            }
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                ProductInfo.OrderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
                ProductInfo.ProductNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
                ProductInfo.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateMaskedTextBox.Text : string.Empty;
                ProductInfo.Person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
                ProductInfo.Revision = RevisionCheckBox.Checked ? RevisionTextBox.Text : string.Empty;
                ProductInfo.Comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;
                ProductInfo.Quantity = Convert.ToInt32(QuantityTextBox.Text ?? throw new Exception("QuantityTextBox.Text is null"));
                ProductInfo.SerialFirstNumber = IsSerialGeneration
                    ? Convert.ToInt32(FirstSerialNumberTextBox.Text ?? throw new Exception("FirstSerialNumberTextBox.Text is null"))
                    : -1;
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
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void ProductRegistration1Window_Load(object sender, EventArgs e) { LoadEvents(); }
        private void QrCodeButton_Click(object sender, EventArgs e) { QrInput(); }
        private void RegisterButton_Click(object sender, EventArgs e) { RegisterCheck(); }
        private void TemplateButton_Click(object sender, EventArgs e) { TemplateComment(); }
        private void NumberCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBoxChecked(sender, e); }
        private void QuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void RegistrationDateMaskedTextBox_TypeValidationCompleted(object sender, TypeValidationEventArgs e) { RegistrationDateCheck(sender, e); }
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
                    ("StrProductModel", $"{ProductInfo.ProductModel}"),
                    ("RevisionGroup", $"{ProductInfo.RevisionGroup}"),
                    ("StrInitial", $"{ProductInfo.Initial}"),
                    ("IntRegType", $"{ProductInfo.RegType}"),
                    ("IntPrintType", $"{ProductInfo.PrintType}"),
                    ("IntSerialDigit", $"{ProductInfo.SerialDigit}")
                };

            const int ColumnWidth = 15;
            var message = string.Join(Environment.NewLine,
                entries.Select(entry => $"{entry.Item1,-ColumnWidth}[{entry.Item2}]"));

            MessageBox.Show(message, "取得情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
