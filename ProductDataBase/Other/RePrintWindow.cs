using Microsoft.Data.Sqlite;
using ProductDatabase.Common;
using ProductDatabase.Data;
using ProductDatabase.Models;
using ProductDatabase.Other;
using ProductDatabase.Print;
using ProductDatabase.Services;
using static ProductDatabase.Print.PrintManager;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase {
    public partial class RePrintWindow : Form {

        public DocumentPrintSettings ProductPrintSettings { get; set; } = new DocumentPrintSettings();
        public LabelPrintSettings LabelPrintSettings => this.ProductPrintSettings.LabelPrintSettings ?? new LabelPrintSettings();
        public BarcodePrintSettings BarcodePrintSettings => ProductPrintSettings.BarcodePrintSettings ?? new BarcodePrintSettings();
        public NameplatePrintSettings NameplatePrintSettings => ProductPrintSettings.NameplatePrintSettings ?? new NameplatePrintSettings();

        public string PrintSettingPath = string.Empty;

        private readonly PrintManager _printManager = new();
        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;
        private readonly AppSettings _appSettings;

        private readonly List<string> _checkBoxNames = [
                    "OrderNumberCheckBox", "ManufacturingNumberCheckBox", "QuantityCheckBox",  "FirstSerialNumberCheckBox",
                    "RevisionCheckBox", "ExtraCheckBox1", "ExtraCheckBox2", "ExtraCheckBox3", "RegistrationDateCheckBox",
                    "PersonCheckBox", "CommentCheckBox" ];

        public RePrintWindow(ProductMaster productMaster, ProductRegisterWork productRegisterWork, AppSettings appSettings) {
            InitializeComponent();

            RegistrationDateCheckBox.AutoCheck = false;
            FirstSerialNumberCheckBox.AutoCheck = false;
            QuantityCheckBox.AutoCheck = false;

            _productMaster = productMaster;
            _productRegisterWork = productRegisterWork;
            _appSettings = appSettings;
        }

        // フォームロード時に初期UI設定・DBからのリビジョン取得・印刷設定読み込みを行う
        private void LoadEvents() {
            try {
                Font = new System.Drawing.Font(_appSettings.FontName, _appSettings.FontSize);

                var inputControls = new Control[] {
                    OrderNumberTextBox, ManufacturingNumberMaskedTextBox, QuantityTextBox, FirstSerialNumberTextBox, RevisionTextBox, CommentTextBox
                };
                foreach (var ctrl in inputControls) {
                    ctrl.TextChanged += InputControls_TextChanged;
                }
                PersonComboBox.SelectedIndexChanged += InputControls_TextChanged;

                ValidateAllInputs();

                ProductNameLabel2.Text = _productMaster.ProductName;
                SubstrateModelLabel2.Text = $"{_productMaster.ProductName} - {_productMaster.ProductModel}";

                FirstSerialNumberTextBox.MaxLength = _productMaster.SerialDigit;
                FirstSerialNumberTextBox.Text = _productMaster.SerialDigit switch {
                    3 => "000",
                    4 => "0000",
                    _ => string.Empty
                };

                // 変数[check_bin]のビットに応じてCheckboxにチェックを入れる
                for (var i = 0; i < _checkBoxNames.Count; i++) {
                    if (Controls[_checkBoxNames[i]] is CheckBox checkBox) {
                        // i番目のビットが1かどうかをチェック
                        checkBox.Checked = (_productMaster.CheckBin & (1 << i)) != 0;
                    }
                }

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. _appSettings.PersonList]);

                // DBから最新リビジョンを取得
                using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
                var revisionResult = ProductRegistrationRepository.GetLatestRevision(
                    con, _productMaster.ProductName, _productMaster.RevisionGroup.ToString());
                RevisionTextBox.Text = revisionResult ?? "";

                ConfigurePrintSettings();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }
        // 製品マスターの印刷フラグに応じてメニューの有効無効を切り替え印刷設定を読み込む
        private void ConfigurePrintSettings() {
            FirstSerialNumberCheckBox.Checked = true;

            シリアルラベル印刷設定ToolStripMenuItem.Enabled = _productMaster.IsLabelPrint;
            バーコード印刷設定ToolStripMenuItem.Enabled = _productMaster.IsBarcodePrint;
            銘版印刷設定ToolStripMenuItem.Enabled = _productMaster.IsNameplatePrint;

            LoadSettings();
        }
        // 製品・型式に対応するJSON印刷設定ファイルを読み込んでProductPrintSettingsに反映する
        private void LoadSettings() {
            try {
                ProductPrintSettings = new DocumentPrintSettings();
                PrintSettingPath = Path.Combine(Environment.CurrentDirectory, "config", "Product", _productMaster.CategoryName, _productMaster.ProductName, $"PrintConfig_{_productMaster.ProductName}_{_productMaster.ProductID}_{_productMaster.ProductModel}.json");
                if (!File.Exists(PrintSettingPath)) { throw new DirectoryNotFoundException($"ラベル印刷用設定ファイルがありません。"); }
                var jsonString = File.ReadAllText(PrintSettingPath);
                ProductPrintSettings = System.Text.Json.JsonSerializer.Deserialize<DocumentPrintSettings>(jsonString) ?? new DocumentPrintSettings();
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 入力チェック後に確認ダイアログを表示してDB登録と印刷処理を実行する
        private async Task RegisterCheck(bool isPrint) {
            if (!DataCheck()) { return; }

            if (isPrint) {
                DialogResult result;
                result = MessageBox.Show("入力に不備がないか確認して下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) { return; }

                result = MessageBox.Show("同一のシリアルラベルが複数存在しないようにして下さい。", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Cancel) { return; }

                _productRegisterWork.Person = PersonComboBox.Text;
                if (!Registration()) { throw new Exception("登録できませんでした。"); }
            }

            switch (_printManager.CurrentSerialType) {
                case SerialType.Nameplate:
                    _printManager.PrintUsingBPac(NameplatePrintSettings, GenerateSerialListForType(_printManager.CurrentSerialType));
                    break;
                case SerialType.Label: {
                        if (!await Print(isPrint, GenerateLabelList(), GenerateOlesListOrNull())) {
                            MessageBox.Show("キャンセルしました。");
                            return;
                        }
                        break;
                    }
                default:
                    if (!await Print(isPrint, GenerateSerialListForType(_printManager.CurrentSerialType))) {
                        MessageBox.Show("キャンセルしました。");
                        return;
                    }
                    break;
            }
        }
        // 再印刷テーブルにレコードを挿入しバックアップとログ記録を行う
        private bool Registration() {
            try {
                _productRegisterWork.RowID = RePrintRepository.InsertRePrintRecord(_productMaster, _productRegisterWork);
                BackupManager.CreateBackup();
                HistoryAuditLogger.LogRePrint(_productMaster, _productRegisterWork);
                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // 入力値のバリデーションを行いシリアルリストと作業データをセットする
        private bool DataCheck() {
            var revision = RevisionTextBox.Text.Trim();
            if (RevisionCheckBox.Checked) {
                if (revision.Any(c => "IO".Contains(char.ToUpperInvariant(c)))) {
                    MessageBox.Show("Revisionに I, O は使用できません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    RevisionTextBox.Focus();
                    return false;
                }
            }

            if (!int.TryParse(QuantityTextBox.Text, out var quantity)) {
                MessageBox.Show("数量は有効な数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                QuantityTextBox.Focus();
                return false;
            }
            if (quantity <= 0) {
                MessageBox.Show("1台以上入力して下さい。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                QuantityTextBox.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(FirstSerialNumberTextBox.Text)) {
                MessageBox.Show("シリアル開始番号を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                FirstSerialNumberTextBox.Focus();
                return false;
            }
            if (!int.TryParse(FirstSerialNumberTextBox.Text, out var firstSerial)) {
                MessageBox.Show("シリアル開始番号が不正な形式です。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                FirstSerialNumberTextBox.Focus();
                return false;
            }
            if (firstSerial <= 0) {
                MessageBox.Show("1以上入力して下さい。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                FirstSerialNumberTextBox.Focus();
                return false;
            }

            var calculatedLastSerial = quantity + firstSerial - 1;

            (var minNumber, var maxNumber, var digit) = _productMaster.GetSerialRange();

            if (calculatedLastSerial > maxNumber || firstSerial < minNumber) {
                MessageBox.Show($"シリアルが範囲外になるため、{minNumber.ToString().PadLeft(digit, '0')}から開始します。", "シリアル番号リセット", MessageBoxButtons.OK, MessageBoxIcon.Information);
                FirstSerialNumberTextBox.Text = minNumber.ToString();
                firstSerial = minNumber;
            }

            _productRegisterWork.OrderNumber = OrderNumberCheckBox.Checked ? OrderNumberTextBox.Text : string.Empty;
            _productRegisterWork.ProductNumber = ManufacturingNumberCheckBox.Checked ? ManufacturingNumberMaskedTextBox.Text : string.Empty;
            _productRegisterWork.Quantity = quantity;
            _productRegisterWork.Person = PersonCheckBox.Checked ? PersonComboBox.Text : string.Empty;
            _productRegisterWork.RegDate = RegistrationDateCheckBox.Checked ? RegistrationDateTimePicker.Value.ToShortDateString() : string.Empty;
            _productRegisterWork.Revision = RevisionCheckBox.Checked ? RevisionTextBox.Text : string.Empty;
            _productRegisterWork.Comment = CommentCheckBox.Checked ? CommentTextBox.Text : string.Empty;

            _productRegisterWork.SerialFirstNumber = firstSerial;
            _productRegisterWork.SerialLastNumber = _productRegisterWork.SerialFirstNumber + _productRegisterWork.Quantity - 1;

            _productRegisterWork.SerialFirst = GenerateCode(_productRegisterWork.SerialFirstNumber);
            _productRegisterWork.SerialLast = GenerateCode(_productRegisterWork.SerialLastNumber);
            return true;
        }

        // isPrintがtrueなら印刷ダイアログ経由で印刷しfalseならプレビューを表示する
        // olesLabelList を渡した場合は OLes 交互印刷として PrintManager に委譲する
        private async Task<bool> Print(bool isPrint, List<string> serialList, List<string>? olesLabelList = null) {
            try {
                using System.Drawing.Printing.PrintDocument pd = new();
                var isPreview = !isPrint;

                var startLine = (int)PrintPositionNumericUpDown.Value - 1;

                pd.BeginPrint += (sender, e) => {
                    _printManager.ProductInitialize(_productMaster, _productRegisterWork, ProductPrintSettings, serialList, olesLabelList);
                };
                pd.PrintPage += (sender, e) => {
                    bool hasMore = _printManager.PrintSerialCommon(e, isPreview, startLine, _printManager.CurrentSerialType);
                    e.HasMorePages = hasMore;
                };

                switch (isPrint) {
                    case true:
                        var pdlg = new PrintDialog {
                            Document = pd
                        };
                        if (pdlg.ShowDialog() == DialogResult.OK) {
                            using var overlay = new LoadingOverlay(this);
                            await CommonUtils.RunOnStaThreadAsync(() => pd.Print());
                        }
                        else {
                            return false;
                        }
                        return true;
                    case false:
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

                        return true;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        // シリアルコードと日付情報からテキストフォーマットに従ってラベル印字コードを生成する
        // type を指定した場合はそのタイプのフォーマットを使用する（null の場合は _printManager.CurrentSerialType を使用）
        private string GenerateCode(int serialCode, SerialType? type = null) {
            var regDate = DateTime.TryParse(_productRegisterWork.RegDate, out var parsedDate)
                ? parsedDate
                : DateTime.Today;
            var monthCode = CommonUtils.ToMonthCode(regDate);

            var resolvedType = type ?? _printManager.CurrentSerialType;
            var outputCode = resolvedType switch {
                SerialType.Label => LabelPrintSettings.LabelTextFormat ?? string.Empty,
                SerialType.OLesLabel => LabelPrintSettings.OLesLabelTextFormat ?? string.Empty,
                SerialType.Barcode => BarcodePrintSettings.LabelTextFormat ?? string.Empty,
                SerialType.Nameplate => NameplatePrintSettings.NameplateTextFormat ?? string.Empty,
                _ => string.Empty
            };

            var map = new Dictionary<string, string> {
                ["{T}"] = _productMaster.Initial ?? string.Empty,
                ["{OT}"] = _productMaster.OLesInitial ?? string.Empty,
                ["{Y}"] = regDate.ToString("yy"),
                ["{MM}"] = regDate.ToString("MM"),
                ["{R}"] = _productRegisterWork.Revision,
                ["{M}"] = monthCode[^1..],
                ["{S}"] = serialCode.ToString($"D{_productMaster.SerialDigit}"),
                ["{SA}"] = _productMaster.OLesSerialSuffix ?? string.Empty
            };

            foreach (var kv in map) {
                outputCode = outputCode.Replace(kv.Key, kv.Value);
            }

            return outputCode;
        }
        // 指定タイプのフォーマットでシリアルリストを生成して返す（_printManager.CurrentSerialType を変更しない）
        private List<string> GenerateSerialListForType(SerialType type) {
            return [.. Enumerable.Range(0, _productRegisterWork.Quantity).Select(i => GenerateCode(_productRegisterWork.SerialFirstNumber + i, type))];
        }
        private List<string> GenerateLabelList() => GenerateSerialListForType(SerialType.Label);
        private List<string>? GenerateOlesListOrNull() =>
            _productMaster.IsOLesLabelPrint ? GenerateSerialListForType(SerialType.OLesLabel) : null;
        // チェックボックスのOn/Offに連動して関連する入力コントロールの有効無効を切り替える
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
                    break;
                case "ExtraCheckBox1":
                    ExtraTextBox1.Enabled = checkBox.Checked;
                    break;
                case "RevisionCheckBox":
                    RevisionTextBox.Enabled = checkBox.Checked;
                    break;
                case "ExtraCheckBox2":
                    ExtraTextBox2.Enabled = checkBox.Checked;
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
            }
            ValidateAllInputs();
        }
        // 数字とバックスペース以外のキー入力を無効化して数値入力専用にする
        private void NumericOnly(object sender, KeyPressEventArgs e) {
            // 0～9と、バックスペース以外の時は、イベントをキャンセルする
            if (e.KeyChar is (< '0' or > '9') and not '\b') {
                e.Handled = true;
            }
        }
        // QRコードテキストを「//」区切りで分割して製番・数量・注番を各入力欄にセットする
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
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 現在の製品・作業データのフィールド値をリスト形式のサブウィンドウで確認表示する
        private void ShowInfo() {
            var items = new Dictionary<string, string>
                {
                    {"ProductName", $"{_productMaster.ProductName}"},
                    {"ProductModel", $"{_productMaster.ProductModel}"},
                    {"ProductType", $"{_productMaster.ProductType}"},
                    {"OrderNumber", $"{_productRegisterWork.OrderNumber}"},
                    {"ProductNumber", $"{_productRegisterWork.ProductNumber}"},
                    {"Revision", $"{_productRegisterWork.Revision}"},
                    {"RegType", $"{_productMaster.RegType}"},
                    {"RegDate", $"{_productRegisterWork.RegDate}"},
                    {"Person", $"{_productRegisterWork.Person}"},
                    {"Quantity", $"{_productRegisterWork.Quantity}"},
                    {"SerialFirstNumber", $"{_productRegisterWork.SerialFirstNumber}"},
                    {"SerialLastNumber", $"{_productRegisterWork.SerialLastNumber}"},
                    {"Initial", $"{_productMaster.Initial}"},
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

        // 全入力コントロールを検証し問題があればエラー表示して印刷ボタンを無効化する
        private void ValidateAllInputs() {
            ErrorMessageLabel.Text = "";

            LabelPrintButton.Enabled = _productMaster.IsLabelPrint;
            BarcodePrintButton.Enabled = _productMaster.IsBarcodePrint;
            NameplatePrintButton.Enabled = _productMaster.IsNameplatePrint;

            シリアルラベル印刷プレビューToolStripMenuItem.Enabled = _productMaster.IsLabelPrint;
            バーコード印刷プレビューToolStripMenuItem.Enabled = _productMaster.IsBarcodePrint;

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
                if (!RNumberCheckBox.Checked && !isValid) {
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

            if (FirstSerialNumberCheckBox.Checked) {
                if (string.IsNullOrWhiteSpace(FirstSerialNumberTextBox.Text)) {
                    ShowError("シリアル開始番号を入力してください。");
                    return;
                }
                if (!int.TryParse(FirstSerialNumberTextBox.Text, out var quantity) || quantity <= 0) {
                    ShowError("シリアル開始番号は1以上の有効な数値を入力してください。");
                    return;
                }
            }

            var revision = RevisionTextBox.Text.Trim();
            if (RevisionCheckBox.Checked) {
                if (revision.Any(c => "IO".Contains(char.ToUpperInvariant(c)))) {
                    ShowError("Revisionに I, O は使用できません。");
                    return;
                }
            }

            if (PersonComboBox.SelectedIndex == -1 && PersonComboBox.Enabled) {
                ShowError("担当者が選択されていません。");
                return;
            }

        }
        // エラーメッセージを赤字で表示して全印刷ボタンを無効化する
        private void ShowError(string message) {
            ErrorMessageLabel.Text = message;
            ErrorMessageLabel.ForeColor = Color.Red;
            LabelPrintButton.Enabled = false;
            BarcodePrintButton.Enabled = false;
            NameplatePrintButton.Enabled = false;
            シリアルラベル印刷プレビューToolStripMenuItem.Enabled = false;
            バーコード印刷プレビューToolStripMenuItem.Enabled = false;
        }

        // 入力コントロール変更時にすべての入力検証を再実行する
        private void InputControls_TextChanged(object? sender, EventArgs e) {
            ValidateAllInputs();
        }

        private void RePrintWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void QrCodeButton_Click(object sender, EventArgs e) { QrInput(); }
        private async void LabelPrintButton_Click(object sender, EventArgs e) {
            _printManager.CurrentSerialType = SerialType.Label;
            await RegisterCheck(true);
        }
        private async void BarcodePrintButton_Click(object sender, EventArgs e) {
            _printManager.CurrentSerialType = SerialType.Barcode;
            await RegisterCheck(true);
        }
        private async void NamePlatePrintButton_Click(object sender, EventArgs e) {
            _printManager.CurrentSerialType = SerialType.Nameplate;
            await RegisterCheck(true);
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) { ShowInfo(); }
        private async void シリアルラベル印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _printManager.CurrentSerialType = SerialType.Label;
            await RegisterCheck(false);
        }
        private async void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            _printManager.CurrentSerialType = SerialType.Barcode;
            await RegisterCheck(false);
        }
        private void シリアルラベル印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            _printManager.CurrentSerialType = SerialType.Label;
            PrintSettingsWindow ls = new() {
                AppSettings = _appSettings,
                CurrentSerialType = _printManager.CurrentSerialType
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void バーコード印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            _printManager.CurrentSerialType = SerialType.Barcode;
            PrintSettingsWindow ls = new() {
                AppSettings = _appSettings,
                CurrentSerialType = _printManager.CurrentSerialType
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void 銘版印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            _printManager.CurrentSerialType = SerialType.Nameplate;
            PrintSettingsWindow ls = new() {
                AppSettings = _appSettings,
                CurrentSerialType = _printManager.CurrentSerialType
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void QrCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }

    }
}
