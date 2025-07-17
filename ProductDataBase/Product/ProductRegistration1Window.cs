using ProductDatabase.Other;
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

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. ProductInfo.PersonList]);

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using SQLiteConnection con = new(GetConnectionRegistration());
                con.Open();
                using var cmd = con.CreateCommand();

                // テーブル検索SQL - [[ProductName]_Product]テーブルの最新の[Revision]を取得
                cmd.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;
                cmd.Parameters.Add("@RevisionGroup", DbType.String).Value = ProductInfo.RevisionGroup;
                cmd.CommandText = $"""SELECT Revision FROM "{ProductInfo.CategoryName}_Product" WHERE ProductName = @ProductName AND RevisionGroup = @RevisionGroup ORDER BY "ID" DESC""";
                var revisionResult = cmd.ExecuteScalar();
                RevisionTextBox.Text = revisionResult?.ToString() ?? "";

                if (ProductInfo.IsSerialGeneration) {
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@ProductName", DbType.String).Value = ProductInfo.ProductName;
                    cmd.CommandText = $"""SELECT SerialLastNumber FROM "{ProductInfo.CategoryName}_Product" WHERE ProductName = @ProductName AND SerialLastNumber NOT NULL ORDER BY "ID" DESC""";
                    var serialResult = cmd.ExecuteScalar();
                    if (!int.TryParse(serialResult?.ToString(), out var serialLastNum)) { throw new Exception("シリアル番号の取得に失敗しました。"); }

                    // シリアル番号の初期値を設定
                    if (ProductInfo.RegType == 0) { return; }
                    var formatString = ProductInfo.SerialDigit switch {
                        3 => "000",
                        4 => "0000",
                        _ => ""
                    };
                    FirstSerialNumberTextBox.Text = (serialLastNum + 1).ToString(formatString);
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

                var revision = RevisionTextBox.Text.Trim();
                if (RevisionCheckBox.Checked) {
                    // revision.Any(...) は、revision 内のいずれかの文字が条件を満たす場合に true を返します。
                    // char.ToUpperInvariant(c) は、文字を大文字に変換し、比較を行います。
                    if (revision.Any(c => "IO".Contains(char.ToUpperInvariant(c)))) {
                        MessageBox.Show("Revisionに I, O は使用できません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        RevisionTextBox.Focus();
                        return;
                    }
                }

                if (ManufacturingNumberCheckBox.Checked && ManufacturingNumberMaskedTextBox.Text.Length != 15) { throw new Exception("製番を10桁+4桁で入力して下さい。"); }

                if (string.IsNullOrWhiteSpace(QuantityTextBox.Text)) {
                    MessageBox.Show("数量を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    QuantityTextBox.Focus();
                    return;
                }
                if (!int.TryParse(QuantityTextBox.Text, out var quantity)) {
                    MessageBox.Show("数量は有効な数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    QuantityTextBox.Focus();
                    return;
                }
                if (quantity <= 0) {
                    MessageBox.Show("1台以上入力して下さい。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    QuantityTextBox.Focus();
                    return;
                }

                var result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) {
                    return;
                }

                var firstSerial = -1;
                if (ProductInfo.IsSerialGeneration) {

                    if (string.IsNullOrWhiteSpace(FirstSerialNumberTextBox.Text)) {
                        MessageBox.Show("シリアル開始番号を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        FirstSerialNumberTextBox.Focus();
                        return;
                    }
                    if (!int.TryParse(FirstSerialNumberTextBox.Text, out firstSerial)) {
                        MessageBox.Show("シリアル開始番号が不正な形式です。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        FirstSerialNumberTextBox.Focus();
                        return;
                    }
                    if (firstSerial <= 0) {
                        MessageBox.Show("1以上入力して下さい。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        FirstSerialNumberTextBox.Focus();
                        return;
                    }

                    // 最終シリアル番号計算
                    var calculatedLastSerial = quantity + firstSerial - 1; // 数量と開始番号から最終シリアルを算出

                    // シリアル番号の桁数に応じて、閾値とリセット値を設定
                    (var minNumber, var maxNumber, var digit) = ProductInfo.SerialType switch {
                        3 => (1, 999, 3),
                        4 => (1, 9999, 4),
                        101 => (1, 899, 3),
                        102 => (901, 999, 3),
                        _ => throw new InvalidOperationException("不明なシリアル桁数です。") // より具体的な例外
                    };

                    if (calculatedLastSerial > maxNumber || firstSerial < minNumber) {// あるいは firstSerialがminNumber未満の場合も対象に
                        MessageBox.Show($"シリアルが範囲外になるため、{minNumber.ToString().PadLeft(digit, '0')}から開始します。", "シリアル番号リセット", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FirstSerialNumberTextBox.Text = minNumber.ToString();
                        firstSerial = minNumber;
                    }
                }

                RegisterButton.Enabled = false;

                ProductInfo.OrderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
                ProductInfo.ProductNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
                ProductInfo.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateTimePicker.Value.ToShortDateString() : string.Empty;
                ProductInfo.Person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
                ProductInfo.Revision = RevisionCheckBox.Checked ? RevisionTextBox.Text : string.Empty;
                ProductInfo.Comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;
                ProductInfo.Quantity = quantity;
                ProductInfo.SerialFirstNumber = ProductInfo.IsSerialGeneration ? firstSerial : -1;
                using ProductRegistration2Window window = new();
                window.ProductInfo = ProductInfo;
                window.Closed += (s, e) => this.Close();
                window.ShowDialog(this);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                RegisterButton.Enabled = true;
            }
        }
        // revisionの変更
        private void RevisionChange() {
            var result = MessageBox.Show("レビジョンを変更しますか？",
                "確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Exclamation,
                MessageBoxDefaultButton.Button2);

            if (result == DialogResult.No) {
                return;
            }

            var revision = RevisionTextBox.Text.Trim();
            // revision.Any(...) は、revision 内のいずれかの文字が条件を満たす場合に true を返します。
            // char.ToUpperInvariant(c) は、文字を大文字に変換し、比較を行います。
            if (revision.Any(c => "IO".Contains(char.ToUpperInvariant(c)))) {
                MessageBox.Show("Revisionに I, O は使用できません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                RevisionTextBox.Focus();
                return;
            }

            ProductInfo.Revision = revision;
            ProductInfo.RegDate = RegistrationDateTimePicker.Value.ToShortDateString();
            ProductInfo.Comment = CommentTextBox.Text.Trim();

            using var connection = new SQLiteConnection(GetConnectionRegistration());
            connection.Open();
            using var transaction = connection.BeginTransaction();
            try {
                var commandText =
                $"""
                INSERT INTO {ProductInfo.CategoryName}_Product (
                    ProductName, ProductType, ProductModel, RegDate, Revision, RevisionGroup, SerialLastNumber, Comment
                )
                VALUES (
                    @ProductName, @ProductType, @ProductModel, @RegDate, @Revision, @RevisionGroup, @SerialLastNumber, @Comment
                );
                """;

                var serialLastNum = ExecuteScalar(connection,
                    $"""SELECT SerialLastNumber FROM {ProductInfo.CategoryName}_Product WHERE ProductName = @ProductName AND SerialLastNumber NOT NULL ORDER BY ID DESC;""",
                    ("@ProductName", ProductInfo.ProductName));

                ExecuteNonQuery(connection, commandText,
                    ("@ProductName", ProductInfo.ProductName),
                    ("@ProductType", ProductInfo.ProductType),
                    ("@ProductModel", ProductInfo.ProductModel),
                    ("@Revision", ProductInfo.Revision),
                    ("@RegDate", ProductInfo.RegDate),
                    ("@RevisionGroup", ProductInfo.RevisionGroup),
                    ("@SerialLastNumber", Convert.ToInt32(serialLastNum)),
                    ("@Comment", ProductInfo.Comment));

                ProductInfo.ProductID = Convert.ToInt32(ExecuteScalar(connection, $"SELECT MAX(ID) FROM {ProductInfo.CategoryName}_Product"));

                transaction.Commit();

                LogRegistration(ProductInfo);

                MessageBox.Show("Revision変更完了", "Revision変更", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } catch (Exception ex) {
                if (transaction.Connection != null) { //接続が開いているか確認する。
                    transaction.Rollback();
                }
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private static void ExecuteNonQuery(SQLiteConnection connection, string commandText, params (string, object?)[] parameters) {
            using var command = connection.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                // 空の文字列の場合にNULLを設定
                var sqlValue = value is string strValue && string.IsNullOrEmpty(strValue) ? DBNull.Value : value ?? DBNull.Value;
                command.Parameters.Add(name, DbType.String).Value = sqlValue;
            }

            command.ExecuteNonQuery();
        }
        private static object ExecuteScalar(SQLiteConnection connection, string commandText, params (string, object)[] parameters) {
            using var command = connection.CreateCommand();
            command.CommandText = commandText;
            foreach (var (name, value) in parameters) {
                command.Parameters.Add(name, DbType.String).Value = value ?? DBNull.Value;
            }

            return command.ExecuteScalar() ?? 0;
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
                    RegistrationDateTimePicker.Enabled = checkBox.Checked;
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
        // ログ出力
        private static void LogRegistration(ProductInformation productInfo) {
            string[] logMessageArray = [
                $"[Rev変更]",
                $"[{productInfo.CategoryName}]",
                $"ID{productInfo.ProductID}]",
                $"[]",
                $"[]",
                $"製品名[{productInfo.ProductName}]",
                $"タイプ[{productInfo.ProductType}]",
                $"型式[{productInfo.ProductModel}]",
                $"[]",
                $"[]",
                $"[]",
                $"Revision[{productInfo.Revision}]",
                $"登録日[{productInfo.RegDate}]",
                $"[]",
                $"コメント[{productInfo.Comment}]"
            ];
            CommonUtils.Logger.AppendLog(logMessageArray);
        }

        private void ProductRegistration1Window_Load(object sender, EventArgs e) { LoadEvents(); }
        private void QrCodeButton_Click(object sender, EventArgs e) { QrInput(); }
        private void RevisionChangeButton_Click(object sender, EventArgs e) { RevisionChange(); }
        private void RegisterButton_Click(object sender, EventArgs e) { RegisterCheck(); }
        private void TemplateButton_Click(object sender, EventArgs e) { TemplateComment(); }
        private void NumberCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBoxChecked(sender, e); }
        private void QuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) {
            var items = new Dictionary<string, string>
                {
                    {"Proness1", $"{ProductInfo.Proness1}" },
                    {"Proness2", $"{ProductInfo.Proness2}"},
                    {"Proness3", $"{ProductInfo.Proness3}"},
                    {"Proness4", $"{ProductInfo.Proness4}"},
                    {"Proness5", $"{ProductInfo.Proness5}"},
                    {"ProductName", $"{ProductInfo.ProductName}"},
                    {"StockName", $"{ProductInfo.StockName}"},
                    {"ProductModel", $"{ProductInfo.ProductModel}"},
                    {"RevisionGroup", $"{ProductInfo.RevisionGroup}"},
                    {"Initial", $"{ProductInfo.Initial}"},
                    {"RegType", $"{ProductInfo.RegType}"},
                    {"PrintType", $"{ProductInfo.PrintType}"},
                    {"SerialDigit", $"{ProductInfo.SerialDigit}"}
                };

            var maxLabelLength = items.Keys.Max(k => k.Length);
            var message = string.Join("\n", items.Select(kvp =>
                kvp.Key.PadRight(maxLabelLength + 4) + $"[ {kvp.Value} ]"
            ));

            var form = new Form {
                Text = "取得情報",
                Width = 300,
                Height = 300,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var label = new Label {
                Text = message,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("PlemolJP", 10)
                //Font = font ?? new Font("MS Gothic", 10)
            };

            form.Controls.Add(label);

            form.ShowDialog();
        }
        private void QrCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }

    }
}
