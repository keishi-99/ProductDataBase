using Dapper;
using Microsoft.Data.Sqlite;
using Newtonsoft.Json.Linq;
using ProductDatabase.Other;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class ProductRegistration1Window : Form {

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;
        private readonly AppSettings _appSettings;

        readonly string _messageFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "Product", "ProductMessages.json");
        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "FirstSerialNumberCheckBox",
                    "RevisionCheckBox", "ExtraCheckBox1", "OLesNumberCheckBox", "ExtraCheckBox3", "RegistrationDateCheckBox",
                    "PersonCheckBox", "CommentCheckBox" ];

        public ProductRegistration1Window(ProductMaster productMaster, ProductRegisterWork productRegisterWork, AppSettings appSettings) {
            InitializeComponent();

            _productMaster = productMaster;
            _productRegisterWork = productRegisterWork;
            _appSettings = appSettings;
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(_appSettings.FontName, _appSettings.FontSize);

                ProductNameLabel2.Text = _productMaster.ProductName;
                SubstrateModelLabel2.Text = $"{_productMaster.ProductName} - {_productMaster.ProductModel}";
                ProductTypeLabel2.Text = _productMaster.ProductType;

                OrderNumberTextBox.Text = _productRegisterWork.OrderNumber;
                ManufacturingNumberMaskedTextBox.Text = !string.IsNullOrEmpty(_productRegisterWork.ProductNumber) ? _productRegisterWork.ProductNumber : ManufacturingNumberMaskedTextBox.Text;
                QuantityTextBox.Text = (_productRegisterWork.Quantity != 0) ? _productRegisterWork.Quantity.ToString() : string.Empty;

                FirstSerialNumberTextBox.MaxLength = _productMaster.SerialDigit;

                RegisterButton.Enabled = true;

                CommentComboBox.Items.Add("[Rev.UP]変更点番号:");

                // 変数[check_bin]のビットに応じてCheckboxにチェックを入れる
                for (var i = 0; i < _checkBoxNames.Count; i++) {
                    if (Controls[_checkBoxNames[i]] is CheckBox checkBox) {
                        // i番目のビットが1かどうかをチェック
                        checkBox.Checked = (_productMaster.CheckBin & (1 << i)) != 0;
                    }
                }

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. _appSettings.PersonList]);

                // DB2へ接続し対象製品テーブルの最新のシリアル,レビジョン取得
                using SqliteConnection con = new(GetConnectionRegistration());

                // テーブル検索SQL - [[ProductName]_Product]テーブルの最新の[Revision]を取得
                var revisionSql =
                    $"""
                    SELECT Revision 
                    FROM {Constants.VProductTableName}
                    WHERE ProductName = @ProductName 
                      AND RevisionGroup = @RevisionGroup 
                      AND IsDeleted = 0 
                    ORDER BY ID DESC
                    LIMIT 1
                    """;

                var revisionResult = con.ExecuteScalar<string>(
                    revisionSql,
                    new {
                        _productMaster.ProductName,
                        _productMaster.RevisionGroup
                    });

                RevisionTextBox.Text = revisionResult ?? "";

                if (_productMaster.IsSerialGeneration) {
                    var serialSql =
                        $"""
                        SELECT SerialLastNumber 
                        FROM {Constants.VProductTableName} 
                        WHERE ProductName = @ProductName 
                          AND IsDeleted = 0 
                          AND SerialLastNumber IS NOT NULL 
                        ORDER BY ID DESC
                        LIMIT 1
                        """;

                    var serialResult = con.ExecuteScalar<int?>(serialSql, new { _productMaster.ProductName })
                        ?? throw new Exception("シリアル番号の取得に失敗しました。");
                    var serialLastNum = serialResult;
                    var nextSerialNumber = serialLastNum + 1;

                    // シリアル番号の桁数に応じて、閾値とリセット値を設定
                    (var minNumber, var maxNumber, var digit) = _productMaster.SerialDigitType switch {
                        3 => (1, 999, 3),
                        4 => (1, 9999, 4),
                        101 => (1, 899, 3),
                        102 => (901, 999, 3),
                        _ => throw new InvalidOperationException("不明なシリアル桁数です。")
                    };

                    if (nextSerialNumber > maxNumber || nextSerialNumber < minNumber) {// あるいは firstSerialが minNumber未満の場合も対象に
                        MessageBox.Show($"シリアルが範囲外になるため、{minNumber.ToString().PadLeft(digit, '0')}から開始します。", "シリアル番号リセット", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        nextSerialNumber = minNumber;
                    }

                    FirstSerialNumberTextBox.Text = nextSerialNumber.ToString();
                }

                // 機種別注意メッセージ表示
                if (!File.Exists(_messageFilePath)) {
                    File.WriteAllText(_messageFilePath, "{}");
                }
                var productMessage = GetProductMessage(_messageFilePath, _productMaster.ProductName);
                if (!string.IsNullOrEmpty(productMessage)) {
                    MessageBox.Show(productMessage, "注意", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    MessageTextBox.Text = productMessage;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 登録チェック
        private void RegisterCheck() {
            try {
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

                if (PersonComboBox.SelectedIndex == -1 && PersonComboBox.Enabled) {
                    MessageBox.Show("担当者が選択されていません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    PersonComboBox.Focus();
                    return;
                }

                if (!anyTextBoxEnabled) { throw new Exception("何も入力されていません"); }
                if (!allTextBoxesFilled || (PersonCheckBox.Checked && string.IsNullOrWhiteSpace(PersonComboBox.Text))) { throw new Exception("空欄があります。"); }

                var revision = RevisionTextBox.Text.Trim();
                if (RevisionCheckBox.Checked) {
                    if (revision.Any(c => "IO".Contains(char.ToUpperInvariant(c)))) {
                        MessageBox.Show("Revisionに I, O は使用できません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        RevisionTextBox.Focus();
                        return;
                    }
                }

                if (ManufacturingNumberCheckBox.Checked) {
                    string text = ManufacturingNumberMaskedTextBox.Text.Trim();
                    bool startsWithR = text.StartsWith("R", StringComparison.OrdinalIgnoreCase);
                    bool hasCorrectLength = text.Length == 15;

                    if (!startsWithR && !hasCorrectLength) {
                        MessageBox.Show("製番を10桁+4桁で入力して下さい。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        ManufacturingNumberMaskedTextBox.Focus();
                        return;
                    }
                }

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
                if (_productMaster.IsSerialGeneration) {

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

                    var calculatedLastSerial = quantity + firstSerial - 1;

                    (var minNumber, var maxNumber, var digit) = _productMaster.SerialDigitType switch {
                        3 => (1, 999, 3),
                        4 => (1, 9999, 4),
                        101 => (1, 899, 3),
                        102 => (901, 999, 3),
                        _ => throw new InvalidOperationException("不明なシリアル桁数です。")
                    };

                    if (calculatedLastSerial > maxNumber || firstSerial < minNumber) {
                        MessageBox.Show($"シリアルが範囲外になるため、{minNumber.ToString().PadLeft(digit, '0')}から開始します。", "シリアル番号リセット", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FirstSerialNumberTextBox.Text = minNumber.ToString();
                        firstSerial = minNumber;
                    }
                }

                RegisterButton.Enabled = false;

                _productRegisterWork.OrderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
                _productRegisterWork.ProductNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
                _productRegisterWork.OLesNumber = OLesNumberCheckBox.Checked ? OLesNumberTextBox.Text : string.Empty;
                _productRegisterWork.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateTimePicker.Value.ToShortDateString() : string.Empty;
                _productRegisterWork.Person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
                _productRegisterWork.Revision = RevisionCheckBox.Checked ? RevisionTextBox.Text : string.Empty;
                _productRegisterWork.Comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;
                _productRegisterWork.Quantity = quantity;
                _productRegisterWork.SerialFirstNumber = _productMaster.IsSerialGeneration ? firstSerial : -1;

                using ProductRegistration2Window window = new(_productMaster, _productRegisterWork, _appSettings);
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
            if (revision.Any(c => "IO".Contains(char.ToUpperInvariant(c)))) {
                MessageBox.Show("Revisionに I, O は使用できません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                RevisionTextBox.Focus();
                return;
            }

            _productRegisterWork.Revision = revision;
            _productRegisterWork.RegDate = RegistrationDateTimePicker.Value.ToShortDateString();
            _productRegisterWork.Comment = CommentTextBox.Text.Trim();

            using var connection = new SqliteConnection(GetConnectionRegistration());
            connection.Open();
            using var transaction = connection.BeginTransaction();
            try {
                var commandText =
                $"""
                INSERT INTO {Constants.TProductTableName} 
                (
                    ProductID,
                    RegDate,
                    Revision,
                    RevisionGroup,
                    SerialLastNumber,
                    Comment
                )
                VALUES 
                (
                    @ProductID, 
                    @RegDate,
                    @Revision,
                    @RevisionGroup,
                    @SerialLastNumber,
                    @Comment
                );
                """;

                var serialLastNum = connection.ExecuteScalar<int>(
                    $"""
                    SELECT SerialLastNumber
                    FROM {Constants.VProductTableName} 
                    WHERE ProductName = @ProductName 
                    AND IsDeleted = 0
                    AND SerialLastNumber IS NOT NULL 
                    ORDER BY ID DESC
                    LIMIT 1
                    """,
                    new { _productMaster.ProductName },
                    transaction: transaction);

                connection.Execute(commandText, new {
                    _productMaster.ProductID,
                    _productRegisterWork.Revision,
                    _productRegisterWork.RegDate,
                    _productMaster.RevisionGroup,
                    SerialLastNumber = serialLastNum,
                    _productRegisterWork.Comment
                }, transaction: transaction);

                var id = connection.ExecuteScalar<int>("SELECT last_insert_rowid();", transaction: transaction);

                transaction.Commit();

                LogRegistration(_productMaster, _productRegisterWork, id);

                MessageBox.Show("Revision変更完了", "Revision変更", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } catch (Exception ex) {
                if (transaction.Connection is not null) {
                    transaction.Rollback();
                }
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    RNumberCheckBox.Enabled = checkBox.Checked;
                    break;
                case "RNumberCheckBox":
                    ManufacturingNumberMaskedTextBox.Mask = checkBox.Checked ? "R00000000000000" : ">LA00A00000-0000";
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
                    ExtraTextBox1.Enabled = checkBox.Checked;
                    break;
                case "OLesNumberCheckBox":
                    OLesNumberTextBox.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox3":
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
            if (string.IsNullOrWhiteSpace(QrCodeTextBox.Text)) { return; }
            string[] separator = ["//"];
            var code = textToUpperCheckBox.Checked ? QrCodeTextBox.Text.ToUpper() : QrCodeTextBox.Text;
            var arr = code.Split(separator, StringSplitOptions.None);
            if (arr.Length != 4) {
                MessageBox.Show("QRコードが正しくありません。");
                return;
            }
            if (arr is not null) {
                ManufacturingNumberMaskedTextBox.Text = arr[0];
                QuantityTextBox.Text = arr[2];
                OrderNumberTextBox.Text = arr[3];
            }
        }
        // ログ出力
        private static void LogRegistration(ProductMaster productMaster, ProductRegisterWork productRegisterWork, long id) {
            string[] logMessageArray = [
                $"[Rev変更]",
                $"[{productMaster.CategoryName}]",
                $"ID{id}]",
                $"[]",
                $"[]",
                $"[]",
                $"製品名[{productMaster.ProductName}]",
                $"タイプ[{productMaster.ProductType}]",
                $"型式[{productMaster.ProductModel}]",
                $"[]",
                $"[]",
                $"[]",
                $"Revision[{productRegisterWork.Revision}]",
                $"登録日[{productRegisterWork.RegDate}]",
                $"[]",
                $"コメント[{productRegisterWork.Comment}]"
            ];
            CommonUtils.Logger.AppendLog(logMessageArray);
        }
        // 取得情報表示
        private void ShowInfo() {
            var items = new Dictionary<string, string>
                {
                    {"ProductID", $"{_productMaster.ProductID}"},
                    {"ProductName", $"{_productMaster.ProductName}"},
                    {"ProductModel", $"{_productMaster.ProductModel}"},
                    {"RevisionGroup", $"{_productMaster.RevisionGroup}"},
                    {"Initial", $"{_productMaster.Initial}"},
                    {"RegType", $"{_productMaster.RegType}"},
                    {"SerialPrintType", $"{_productMaster.SerialPrintType}"},
                    {"SheetPrintType", $"{_productMaster.SheetPrintType}"},
                    {"SerialDigit", $"{_productMaster.SerialDigit}"}
                };

            var form = new Form {
                Text = "取得情報",
                Width = 300,
                Height = 400,
                AutoSize = true,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                ShowInTaskbar = false,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var listView = new ListView {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Font = new Font("PlemolJP", _appSettings.FontSize),
            };

            listView.Columns.Add("", 0);
            listView.Columns.Add("項目", 200, HorizontalAlignment.Right);
            listView.Columns.Add("値", 360);

            foreach (var kvp in items) {
                var item = new ListViewItem("");
                item.SubItems.Add(kvp.Key);
                item.SubItems.Add(kvp.Value);
                listView.Items.Add(item);
            }
            form.Controls.Add(listView);

            form.Shown += (_, _) => {
                listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            };
            form.ShowDialog();
        }

        // JSON から機種別注意メッセージ取得
        public static string? GetProductMessage(string filePath, string productName) {
            var jsonText = File.ReadAllText(filePath);
            var jsonObj = JObject.Parse(jsonText);

            return jsonObj[productName]?.ToString();
        }
        // 注意メッセージ更新
        private static readonly object s_fileLock = new();
        private void ProductMessageChange() {
            lock (s_fileLock) {
                if (!File.Exists(_messageFilePath)) {
                    throw new FileNotFoundException($"{_messageFilePath}\nが見つかりません。");
                }

                var json = JObject.Parse(File.ReadAllText(_messageFilePath));

                string currentMessage = json.ContainsKey(_productMaster.ProductName)
                    ? json[_productMaster.ProductName]!.ToString()
                    : "";

                using var dialog = new InputDialog("メッセージ編集", "メッセージを入力してください", currentMessage);

                if (dialog.ShowDialog() == DialogResult.OK) {
                    string newMessage = dialog.InputText;

                    json[_productMaster.ProductName] = newMessage;

                    File.WriteAllText(_messageFilePath, json.ToString());
                }
            }
        }
        public class InputDialog : Form {
            private readonly TextBox _textBox;
            public string InputText => _textBox.Text;

            public InputDialog(string title, string message, string defaultText = "") {
                Text = title;
                Width = 400;
                Height = 180;
                StartPosition = FormStartPosition.CenterParent;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;

                Label label = new Label() {
                    Text = message,
                    Left = 20,
                    Top = 20,
                    Width = 340
                };

                _textBox = new TextBox() {
                    Left = 20,
                    Top = 50,
                    Width = 340,
                    MaxLength = 80,
                    Text = defaultText
                };

                Button ok = new Button() {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Left = 180,
                    Top = 85,
                    Width = 80
                };

                Button cancel = new Button() {
                    Text = "キャンセル",
                    DialogResult = DialogResult.Cancel,
                    Left = 280,
                    Top = 85,
                    Width = 80
                };

                Controls.Add(label);
                Controls.Add(_textBox);
                Controls.Add(ok);
                Controls.Add(cancel);

                AcceptButton = ok;
                CancelButton = cancel;
            }
        }

        private void ProductRegistration1Window_Load(object sender, EventArgs e) { LoadEvents(); }
        private void QrCodeButton_Click(object sender, EventArgs e) { QrInput(); }
        private void RevisionChangeButton_Click(object sender, EventArgs e) { RevisionChange(); }
        private void RegisterButton_Click(object sender, EventArgs e) { RegisterCheck(); }
        private void TemplateButton_Click(object sender, EventArgs e) { TemplateComment(); }
        private void NumberCheckBox_CheckedChanged(object sender, EventArgs e) { CheckBoxChecked(sender, e); }
        private void QuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void メッセージ設定ToolStripMenuItem_Click(object sender, EventArgs e) { ProductMessageChange(); }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) { ShowInfo(); }
        private void QrCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }

    }
}
