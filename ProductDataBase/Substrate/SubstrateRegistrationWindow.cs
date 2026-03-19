using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.ExcelService;
using ProductDatabase.Models;
using ProductDatabase.Other;
using ProductDatabase.Print;
using static ProductDatabase.Data.ProductRepository;
using static ProductDatabase.Print.PrintManager;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase {
    public partial class SubstrateRegistrationWindow : Form {

        private readonly SubstrateMaster _substrateMaster;
        private readonly SubstrateRegisterWork _substrateRegisterWork;
        private readonly AppSettings _appSettings;

        public DocumentPrintSettings SubstratePrintSettings { get; set; } = new DocumentPrintSettings();
        public LabelPrintSettings LabelPrintSettings => SubstratePrintSettings.LabelPrintSettings ?? new LabelPrintSettings();
        public string PrintSettingPath { get; } = Path.Combine(Environment.CurrentDirectory, "config", "Substrate", "SubstrateConfig.json");

        private class SubstrateStockDto {
            public long SubstrateID { get; set; }
            public string SubstrateName { get; set; } = string.Empty;
            public string SubstrateModel { get; set; } = string.Empty;
            public string SubstrateNumber { get; set; } = string.Empty;
            public string OrderNumber { get; set; } = string.Empty;
            public int Stock { get; set; }
        }

        private readonly List<string> _serialList = [];
        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox", "DefectQuantityCheckBox",
                    "ExtraCheckBox7", "ExtraCheckBox1", "ExtraCheckBox2", "ExtraCheckBox3", "RegistrationDateCheckBox",
                    "PersonCheckBox", "CommentCheckBox"
        ];

        // プロパティ設定
        private bool IsRegistration => _substrateMaster.RegType is 0 or 1;

        public SubstrateRegistrationWindow(SubstrateMaster substrateMaster, SubstrateRegisterWork substrateRegisterWork, AppSettings appSettings) {
            InitializeComponent();

            RegistrationDateCheckBox.AutoCheck = false;

            _substrateMaster = substrateMaster;
            _substrateRegisterWork = substrateRegisterWork;
            _appSettings = appSettings;
        }

        // フォームロード時にUIを初期化しDBから在庫数取得・チェックボックス状態設定・印刷設定読み込みを行う
        private void LoadEvents() {
            try {
                Font = new Font(_appSettings.FontName, _appSettings.FontSize);

                var inputControls = new Control[] {
                    OrderNumberTextBox, ManufacturingNumberMaskedTextBox, QuantityTextBox, DefectQuantityTextBox, CommentTextBox
                };
                foreach (var ctrl in inputControls) {
                    ctrl.TextChanged += InputControls_TextChanged;
                }
                PersonComboBox.SelectedIndexChanged += InputControls_TextChanged;

                ValidateAllInputs();

                ProductNameLabel2.Text = _substrateMaster.ProductName;

                SubstrateModelLabel2.Text = $"{_substrateMaster.SubstrateName} - {_substrateMaster.SubstrateModel}";

                var stockQuantity = GetStockQuantity();
                StockLabel2.Text = stockQuantity;

                OrderNumberTextBox.Text = _substrateRegisterWork.OrderNumber;
                ManufacturingNumberMaskedTextBox.Text = !string.IsNullOrEmpty(_substrateRegisterWork.ProductNumber) ? _substrateRegisterWork.ProductNumber : ManufacturingNumberMaskedTextBox.Text;
                QuantityTextBox.Text = (_substrateRegisterWork.AddQuantity != 0) ? _substrateRegisterWork.AddQuantity.ToString() : string.Empty;

                RegisterButton.Enabled = true;

                CommentComboBox.Items.Add("[Rev.UP]変更点番号:");

                // 変数[check_bin]のビットに応じてCheckboxにチェックを入れる
                for (var i = 0; i < _checkBoxNames.Count; i++) {
                    if (Controls[_checkBoxNames[i]] is CheckBox checkBox) {
                        // i番目のビットが1かどうかをチェック
                        checkBox.Checked = (_substrateMaster.CheckBin & (1 << i)) != 0;
                    }
                }

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. _appSettings.PersonList]);

                // 印刷しない場合は関連コントロール非表示に
                if (!_substrateMaster.IsLabelPrint) {
                    PrintRowLabel.Visible = false;
                    PrintPositionNumericUpDown.Visible = false;
                    PrintOnlyCheckBox.Visible = false;
                    PrintButton.Visible = false;
                }

                if (!File.Exists(PrintSettingPath)) { throw new Exception("印刷設定ファイルが見つかりませんでした"); }
                LoadSettings(PrintSettingPath);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }
        // 指定パスのJSONファイルから印刷設定を読み込みSubstratePrintSettingsに反映する
        private void LoadSettings(string settingFilePath) {
            try {
                var jsonString = File.ReadAllText(settingFilePath);
                SubstratePrintSettings = System.Text.Json.JsonSerializer.Deserialize<DocumentPrintSettings>(jsonString) ?? new DocumentPrintSettings();
            } catch (Exception ex) {
                MessageBox.Show($"設定ファイルの読み込みに失敗しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 入力検証後にDB登録を行いラベル印刷フラグに応じて印刷またはプレビューを実行してウィンドウを閉じる
        private async Task ProcessRegistration(bool isPrint) {
            try {

                if (!ValidateData()) { return; }

                RegisterButton.Enabled = false;

                if (!PrintOnlyCheckBox.Checked && isPrint) {
                    if (!Registration()) {
                        return;
                    }
                }

                if (!_substrateMaster.IsLabelPrint) {
                    MessageBox.Show("登録完了");
                    Close();
                    return;
                }

                if (QuantityCheckBox.Checked) {
                    if (isPrint) {
                        MessageBox.Show("登録完了 続けて印刷します。");
                    }
                    await PrintStart(isPrint);
                }
                else if (DefectQuantityCheckBox.Checked) {
                    MessageBox.Show("登録完了");
                }

                Close();

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                RegisterButton.Enabled = true;
            }
        }
        // トランザクションで基板登録テーブルに入荷/不良レコードを挿入しログ記録とバックアップを行う
        private bool Registration() {
            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();

            using var transaction = con.BeginTransaction();
            try {
                var orderNumber = _substrateRegisterWork.OrderNumber;
                var substrateNumber = _substrateRegisterWork.ProductNumber;
                var quantity = _substrateRegisterWork.AddQuantity;
                var defectQuantity = string.IsNullOrWhiteSpace(DefectQuantityTextBox.Text) ? 0 : Convert.ToInt32(DefectQuantityTextBox.Text);
                var registrationDate = _substrateRegisterWork.RegDate;
                var person = _substrateRegisterWork.Person;
                var comment = _substrateRegisterWork.Comment;
                var rowId = string.Empty;
                var commandText = string.Empty;

                // 製番が新規かチェック
                if (IsRegistration) {
                    commandText =
                        $"""
                        SELECT
                            SubstrateID,
                            SubstrateName,
                            SubstrateModel,
                            SubstrateNumber,
                            OrderNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                        FROM
                            {Constants.VSubstrateTableName}
                        WHERE
                            SubstrateID = @SubstrateID
                            AND SubstrateNumber = @SubstrateNumber
                        GROUP BY
                            SubstrateID, 
                            SubstrateName,
                            SubstrateModel,
                            SubstrateNumber,
                            OrderNumber
                        ORDER BY
                            MIN(ID)
                        LIMIT 1
                        ;
                        """;
                    var result = con.QueryFirstOrDefault<SubstrateStockDto>(
                        commandText,
                        new {
                            _substrateMaster.SubstrateID,
                            SubstrateNumber = substrateNumber
                        },
                        transaction: transaction);

                    var substrateName = result?.SubstrateName ?? string.Empty;

                    if (substrateName != string.Empty) {
                        if (_substrateMaster.SubstrateName == substrateName) {
                            if (QuantityCheckBox.Checked && !DefectQuantityCheckBox.Checked) {
                                var dialogResult = MessageBox.Show($"[{substrateNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                                if (dialogResult == DialogResult.No) { return false; }
                            }
                        }
                    }
                }

                // 不良処理時在庫チェック
                if (DefectQuantityCheckBox.Checked) {
                    commandText =
                        $"""
                        SELECT
                            SubstrateID,
                            SubstrateName,
                            SubstrateModel,
                            SubstrateNumber,
                            OrderNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                        FROM
                            {Constants.VSubstrateTableName}
                        WHERE
                            SubstrateID = @SubstrateID 
                            AND SubstrateModel = @SubstrateModel 
                            AND SubstrateNumber = @SubstrateNumber
                        GROUP BY
                            SubstrateID, 
                            SubstrateName, 
                            SubstrateModel,
                            SubstrateNumber,
                            OrderNumber
                        ORDER BY
                            MIN(ID)
                        LIMIT 1
                        ;
                        """;
                    var stockResult = con.QueryFirstOrDefault<SubstrateStockDto>(
                        commandText,
                        new {
                            _substrateMaster.SubstrateID,
                            _substrateMaster.SubstrateModel,
                            SubstrateNumber = substrateNumber
                        },
                        transaction: transaction)
                        ?? throw new Exception($"[{substrateNumber}]は登録がありません。");

                    if (stockResult.Stock < defectQuantity) {
                        throw new Exception($"[{substrateNumber}]は在庫が[{stockResult.Stock}]です。");
                    }
                }

                // 基板登録テーブルへ追加
                commandText =
                    $"""
                    INSERT INTO {Constants.TSubstrateTableName}
                    (
                        SubstrateID,
                        SubstrateNumber,
                        OrderNumber,
                        Increase,
                        Defect,
                        Person,
                        RegDate,
                        Comment
                    )
                    VALUES 
                    (
                        @SubstrateID,
                        @SubstrateNumber,
                        @OrderNumber,
                        @Increase,
                        @Defect,
                        @Person,
                        @RegDate,
                        @Comment
                    )
                    ;
                    """;

                con.Execute(commandText, new {
                    _substrateMaster.SubstrateID,
                    SubstrateNumber = string.IsNullOrWhiteSpace(substrateNumber) ? (object)DBNull.Value : substrateNumber,
                    OrderNumber = string.IsNullOrWhiteSpace(orderNumber) ? (object)DBNull.Value : orderNumber,
                    Increase = QuantityCheckBox.Checked ? (object)quantity : DBNull.Value,
                    Defect = DefectQuantityCheckBox.Checked ? (object)$"-{defectQuantity}" : DBNull.Value,
                    RegDate = string.IsNullOrWhiteSpace(registrationDate) ? (object)DBNull.Value : registrationDate,
                    Person = string.IsNullOrWhiteSpace(person) ? (object)DBNull.Value : person,
                    Comment = string.IsNullOrWhiteSpace(comment) ? (object)DBNull.Value : comment
                }, transaction: transaction);

                rowId = con.ExecuteScalar<string>("SELECT last_insert_rowid();", transaction: transaction);

                // ログ出力
                var logQuantity = QuantityCheckBox.Checked ? quantity.ToString() : string.Empty;
                var logDefectQuantity = DefectQuantityCheckBox.Checked ? (0 - defectQuantity).ToString() : string.Empty;

                string[] logMessageArray = [
                    $"[基板登録]",
                    $"[{_substrateMaster.CategoryName}]",
                    $"ID[{rowId}]",
                    $"注文番号[{orderNumber}]",
                    $"製造番号[{substrateNumber}]",
                    $"[]",
                    $"製品名[{_substrateMaster.ProductName}]",
                    $"基板名[{_substrateMaster.SubstrateName}]",
                    $"型式[{_substrateMaster.SubstrateModel}]",
                    $"追加数[{logQuantity}]",
                    $"使用数[]",
                    $"減少数[{logDefectQuantity}]",
                    $"[]",
                    $"登録日[{registrationDate}]",
                    $"担当者[{person}]",
                    $"コメント[{comment}]"
                ];
                CommonUtils.Logger.AppendLog(logMessageArray);

                transaction.Commit();

                // バックアップ作成
                CommonUtils.BackupManager.CreateBackup();

                return true;
            } catch (Exception ex) {
                if (transaction.Connection is not null) { // 接続が開いているか確認する。
                    transaction.Rollback();
                }
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

        }
        // 入力値の数値チェックと範囲検証を行いWorkオブジェクトへ値をセットしてシリアルリストを生成する
        private bool ValidateData() {

            var quantity = 0;
            var defectQuantity = 0;

            if (QuantityCheckBox.Checked) {
                if (!int.TryParse(QuantityTextBox.Text, out quantity) || quantity <= 0) {
                    MessageBox.Show("数量は1以上の有効な数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    QuantityTextBox.Focus();
                    return false;
                }
            }
            if (DefectQuantityCheckBox.Checked) {
                if (!int.TryParse(DefectQuantityTextBox.Text, out defectQuantity) || defectQuantity <= 0) {
                    MessageBox.Show("数量は1以上の有効な数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    DefectQuantityTextBox.Focus();
                    return false;
                }
            }

            _substrateRegisterWork.OrderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
            _substrateRegisterWork.ProductNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
            _substrateRegisterWork.AddQuantity = quantity;
            _substrateRegisterWork.DefectQuantity = defectQuantity;
            _substrateRegisterWork.Person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
            _substrateRegisterWork.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateTimePicker.Value.ToShortDateString() : string.Empty;
            _substrateRegisterWork.Comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;

            _serialList.Clear();

            string text = ManufacturingNumberMaskedTextBox.Text.Trim();

            var substrateNumber = otherNumberCheckBox.Checked
                ? text[^6..]
                : text.Substring(5, 5);

            var regDate = DateTime.TryParse(_substrateRegisterWork.RegDate, out var parsedDate)
                ? parsedDate
                : DateTime.Today;
            var monthCode = CommonUtils.ToMonthCode(regDate);
            var map = new Dictionary<string, string> {
                ["{Y}"] = regDate.ToString("yy"),
                ["{MM}"] = regDate.ToString("MM"),
                ["{M}"] = monthCode[^1..],
                ["{S}"] = substrateNumber
            };

            var outputCodes = LabelPrintSettings.TextFormat ?? string.Empty;
            foreach (var kv in map) {
                outputCodes = outputCodes.Replace(kv.Key, kv.Value);
            }

            for (var i = 0; i < quantity; i++) {
                _serialList.Add(outputCodes);
            }

            return true;
        }

        // PrintDocumentを使って基板ラベルを印刷またはプレビュー表示する
        private async Task PrintStart(bool isPrint) {
            try {
                using System.Drawing.Printing.PrintDocument pd = new();

                var isPreview = !isPrint;
                var startLine = (int)PrintPositionNumericUpDown.Value - 1;
                CurrentSerialType = SerialType.Substrate;

                pd.BeginPrint += (sender, e) => {
                    PrintManager.SubstrateInitialize(_substrateMaster, _substrateRegisterWork, SubstratePrintSettings, _serialList);
                };
                pd.PrintPage += (sender, e) => {
                    var hasMore = PrintManager.PrintSerialCommon(e, isPreview, startLine, CurrentSerialType);
                    e.HasMorePages = hasMore;
                };

                _substrateRegisterWork.AddQuantity = int.TryParse(QuantityTextBox.Text, out var quantity) ? quantity : 0;
                if (_substrateRegisterWork.AddQuantity == 0) {
                    throw new Exception("数量が入力されていません。");
                }

                switch (isPrint) {
                    case true:
                        var pdlg = new PrintDialog {
                            Document = pd
                        };
                        if (pdlg.ShowDialog() == DialogResult.OK) {
                            using (var overlay = new LoadingOverlay(this)) {
                                await Task.Run(() => pd.Print());
                            }
                        }
                        break;
                    case false:
                        _substrateRegisterWork.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateTimePicker.Value.ToShortDateString() : string.Empty;
                        var ppd = new PrintPreviewDialog();
                        ppd.Shown += (sender, e) => {
                            var tool = (ToolStrip)ppd.Controls[1];
                            tool.Items[0].Visible = false;
                            if (sender is Form form) {
                                form.WindowState = FormWindowState.Maximized;
                            }
                        };
                        ppd.PrintPreviewControl.Zoom = 3;
                        ppd.Document = pd;
                        ppd.ShowDialog();
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // DBから対象基板の現在庫数を集計して文字列で返す
        private string GetStockQuantity() {
            using var con = new SqliteConnection(GetConnectionRegistration());

            var commandText =
                $"""
                SELECT
                    SubstrateID,
                    SubstrateName,
                    SubstrateModel,
                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                FROM {Constants.VSubstrateTableName}
                WHERE SubstrateID = @SubstrateID
                GROUP BY SubstrateID, SubstrateName, SubstrateModel
                """;

            var stock = con.ExecuteScalar<int?>(
                commandText,
                new { _substrateMaster.SubstrateID });

            return stock?.ToString() ?? "0";
        }

        // ComboBoxで選択したテンプレート文字列をコメントテキストボックスに追記する
        private void TemplateComment() {
            var templateWord = CommentComboBox.SelectedIndex switch {
                0 => "[Rev.UP]変更点番号:",
                _ => string.Empty
            };
            CommentTextBox.Text = $"{CommentTextBox.Text}{templateWord}";
        }
        // チェックボックスのON/OFFに応じて対応する入力コントロールの有効状態を切り替え入力検証を再実行する
        private void CheckBoxChecked(object sender, EventArgs e) {
            var checkBox = (CheckBox)sender;

            switch (checkBox.Name) {
                case "OrderNumberCheckBox":
                    OrderNumberTextBox.Enabled = checkBox.Checked;
                    break;
                case "ManufacturingNumberCheckBox":
                    ManufacturingNumberMaskedTextBox.Enabled = checkBox.Checked;
                    otherNumberCheckBox.Enabled = checkBox.Checked;
                    break;
                case "otherNumberCheckBox":
                    ManufacturingNumberMaskedTextBox.Mask = checkBox.Checked ? "AAAAAAAAAAAAAA" : ">LA00A00000-0000";
                    ManufacturingNumberMaskedTextBox.Text = string.Empty;
                    break;
                case "QuantityCheckBox":
                    QuantityTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) { DefectQuantityCheckBox.Checked = false; }
                    break;
                case "DefectQuantityCheckBox":
                    DefectQuantityTextBox.Enabled = checkBox.Checked;
                    if (checkBox.Checked) { QuantityCheckBox.Checked = false; }
                    break;
                case "ExtraCheckBox7":
                    ExtraTextBox7.Enabled = checkBox.Checked;
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
                case "PrintOnlyCheckBox":
                    PrintButton.Enabled = checkBox.Checked;
                    RegisterButton.Enabled = !checkBox.Checked;
                    if (PrintOnlyCheckBox.Checked) { MessageBox.Show($"登録せずに印刷だけ行います。", "注意", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                    break;
                default:
                    break;
            }
            ValidateAllInputs();
        }
        // KeyPressイベントで数字とバックスペース以外の入力をキャンセルする
        private void NumericOnly(object sender, KeyPressEventArgs e) {
            // 0～9と、バックスペース以外の時は、イベントをキャンセルする
            if (e.KeyChar is (< '0' or > '9') and not '\b') {
                e.Handled = true;
            }
        }
        // QRコードテキストを"//"区切りでパースし製番・数量・注文番号の各フィールドに反映する
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
                ManufacturingNumberMaskedTextBox.Text = arr[0];
                QuantityTextBox.Text = arr[2];
                OrderNumberTextBox.Text = arr[3];
            } catch (Exception ex) {
                throw new Exception($"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", ex);
            }
        }
        // 現在の基板・作業データのフィールド値をListViewで確認表示するデバッグ用ダイアログを開く
        private void ShowInfo() {
            var items = new Dictionary<string, string>
                {
                    {"SubstrateID", $"{_substrateMaster.SubstrateID}"},
                    {"ProductName", $"{_substrateMaster.ProductName}"},
                    {"SubstrateName", $"{_substrateMaster.SubstrateName}"},
                    {"SubstrateModel", $"{_substrateMaster.SubstrateModel}"},
                    {"ProductNumber", $"{_substrateRegisterWork.ProductNumber}"},
                    {"RegType", $"{_substrateMaster.RegType}"},
                    {"SerialPrintType", $"{_substrateMaster.SerialPrintType}"},
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
                Font = new Font("PlemolJP", _appSettings.FontSize), // 等幅フォント
            };

            listView.Columns.Add("", 0);   // 値列の幅（調整可）
            listView.Columns.Add("項目", 200, HorizontalAlignment.Right);  // 項目列の幅
            listView.Columns.Add("値", 360);   // 値列の幅（調整可）

            foreach (var kvp in items) {
                var item = new ListViewItem("");  // ダミー1列目
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
        // 基板マスター情報をもとに基板仕様Excelファイルを開く（例外はcaller側でハンドル）
        private void OpenSubstrateInformation() {
            ExcelServiceClosedXml.SubstrateInformationClosedXml.OpenSubstrateInformationClosedXml(_substrateMaster);
        }

        // 有効なテキストボックスの入力状態・製番桁数・数量・担当者選択を検証しエラー時は登録ボタンを無効化する
        private void ValidateAllInputs() {
            ErrorMessageLabel.Text = "";
            RegisterButton.Enabled = true;
            印刷プレビューToolStripMenuItem.Enabled = true;

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

            // 未入力チェック
            if (!anyTextBoxEnabled) {
                ShowError("何も入力されていません");
                return;
            }

            // 空欄チェック
            if (!allTextBoxesFilled) {
                ShowError("空欄があります。");
                return;
            }

            string manufacturingNumber = ManufacturingNumberMaskedTextBox.Text.Trim();
            if (ManufacturingNumberCheckBox.Checked) {
                bool isValid = manufacturingNumber.Length == 15;
                if (!otherNumberCheckBox.Checked && !isValid) {
                    ShowError("製番を10桁+4桁で入力して下さい。");
                    return;
                }
            }

            if (QuantityCheckBox.Checked) {
                if (string.IsNullOrWhiteSpace(QuantityTextBox.Text)) {
                    ShowError("数量を入力してください。");
                    return;
                }
                if (!int.TryParse(QuantityTextBox.Text, out var quantity) || quantity <= 0) {
                    ShowError("数量は1以上の有効な数値を入力してください。");
                    return;
                }
            }

            if (PersonComboBox.SelectedIndex == -1 && PersonComboBox.Enabled) {
                ShowError("担当者が選択されていません。");
                return;
            }

        }
        // エラーメッセージラベルに赤字でメッセージを表示し登録ボタンと印刷プレビューを無効化する
        private void ShowError(string message) {
            ErrorMessageLabel.Text = message;
            ErrorMessageLabel.ForeColor = Color.Red;
            RegisterButton.Enabled = false;
            印刷プレビューToolStripMenuItem.Enabled = false;
        }

        // 入力コントロール変更時にすべての入力検証を再実行する
        private void InputControls_TextChanged(object? sender, EventArgs e) {
            ValidateAllInputs();
        }

        private void SubstrateRegistrationWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void QrCodeButton_Click(object sender, EventArgs e) { QrInput(); }
        private async void RegisterButton_Click(object sender, EventArgs e) { await ProcessRegistration(true); }
        private async void PrintButton_Click(object sender, EventArgs e) { await ProcessRegistration(true); }
        private async void OpenSubstrateInformationButton_Click(object sender, EventArgs e)
        {
            OpenSubstrateInformationButton.Enabled = false;
            try
            {
                await CommonUtils.RunOnStaThreadAsync(OpenSubstrateInformation);
            }
            catch (Exception ex)
            {
                // UIスレッドで MessageBox を表示（スレッド安全）
                MessageBox.Show(ex.Message, $"[{nameof(OpenSubstrateInformationButton_Click)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                OpenSubstrateInformationButton.Enabled = true;
            }
        }
        private void TemplateButton_Click(object sender, EventArgs e) { TemplateComment(); }
        private void QuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private void DefectQuantityTextBox_KeyPress(object sender, KeyPressEventArgs e) { NumericOnly(sender, e); }
        private async void 印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) { await ProcessRegistration(false); }
        private void 印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            CurrentSerialType = SerialType.Substrate;
            using (var ls = new PrintSettingsWindow {
                SubstrateMaster = _substrateMaster
            }) {
                ls.ShowDialog(this);
            }
            LoadSettings(PrintSettingPath);
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) { ShowInfo(); }
        private void QrCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }

    }
}
