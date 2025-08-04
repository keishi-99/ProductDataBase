using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using static ProductDatabase.MainWindow;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase {
    public partial class PrintSettingsWindow : Form {

        public DocumentPrintSettings DocumentPrintSettings { get; set; } = new DocumentPrintSettings();
        private string _documentPrintSettingFilePath = string.Empty;
        private JsonSerializerOptions? _jsonSerializerOptions;

        public ProductInformation ProductInfo { get; set; } = new ProductInformation();

        public string serialType = string.Empty;

        public PrintSettingsWindow() {
            InitializeComponent();
        }

        private void LoadSettings() {
            switch (Owner) {
                case ProductRegistration2Window productWindow:
                    ProductInfo = productWindow.ProductInfo;
                    LoadSettingsFromWindow(productWindow.ProductPrintSettings, productWindow.printSettingPath);
                    break;
                case SubstrateRegistrationWindow substrateWindow:
                    ProductInfo = substrateWindow.ProductInfo;
                    LoadSettingsFromWindow(substrateWindow.SubstratePrintSettings, substrateWindow.printSettingPath);
                    break;
                case RePrintWindow rePrintWindow:
                    LoadSettingsFromWindow(rePrintWindow.ProductPrintSettings, rePrintWindow.printSettingPath);
                    break;
                default:
                    MessageBox.Show("この画面を開くには正しいウィンドウから開いてください。");
                    Close();
                    break;
            }
        }
        private void LoadSettingsFromWindow(DocumentPrintSettings settings, string filePath) {
            DocumentPrintSettings = settings;
            _documentPrintSettingFilePath = filePath;

            switch (serialType) {
                case "Label":
                    BarcodePanel.Visible = false;
                    SetPageSettings(DocumentPrintSettings.LabelPageSettings);
                    SetLabelSettings(DocumentPrintSettings.LabelLayoutSettings);
                    break;
                case "Barcode":
                    BarcodePanel.Visible = true;

                    SetPageSettings(DocumentPrintSettings.BarcodePageSettings);
                    SetLabelSettings(DocumentPrintSettings.BarcodeLayoutSettings);
                    break;
                case "Substrate":
                    BarcodePanel.Visible = false;
                    SetPageSettings(DocumentPrintSettings.LabelPageSettings);
                    SetLabelSettings(DocumentPrintSettings.LabelLayoutSettings);
                    break;
            }
        }

        private void SetLabelSettings(PrintLayoutSettings? printLayoutSettings) {
            if (printLayoutSettings != null) {
                BarcodeHeightTextBox.Text = printLayoutSettings.BarcodeHeight.ToString();
                BarcodeWidthTextBox.Text = printLayoutSettings.BarcodeWidth.ToString();
                CopiesPerLabelTextBox.Text = printLayoutSettings.CopiesPerLabel.ToString();
                FormatTextBox.Text = printLayoutSettings.TextFormat;

                if (printLayoutSettings.TextFont != null) {
                    FontDialog.Font = printLayoutSettings.TextFont;
                    FontTextBox.Text = $"{FontDialog.Font.Name} {FontDialog.Font.SizeInPoints}pt";
                }

                BarcodePositionXTextBox.Text = printLayoutSettings.BarcodePositionX.ToString();
                BarcodePositionYTextBox.Text = printLayoutSettings.BarcodePositionY.ToString();
                AlignBarcodeCenterXCheckBox.Checked = printLayoutSettings.AlignBarcodeCenterX;

                TextPositionXTextBox.Text = printLayoutSettings.TextPositionX.ToString();
                TextPositionYTextBox.Text = printLayoutSettings.TextPositionY.ToString();
                AlignTextCenterXCheckBox.Checked = printLayoutSettings.AlignTextCenterX;
                AlignTextCenterYCheckBox.Checked = printLayoutSettings.AlignTextCenterY;
            }
            BarcodePositionXTextBox.Enabled = !AlignBarcodeCenterXCheckBox.Checked;
            TextPositionXTextBox.Enabled = !AlignTextCenterXCheckBox.Checked;
            TextPositionYTextBox.Enabled = !AlignTextCenterYCheckBox.Checked;
        }
        private void SetPageSettings(PrintPageSettings? printPageSettings) {
            if (printPageSettings != null) {
                LabelWidthTextBox.Text = printPageSettings.LabelWidth.ToString();
                LabelHeightTextBox.Text = printPageSettings.LabelHeight.ToString();
                LabelsPerRowTextBox.Text = printPageSettings.LabelsPerRow.ToString();
                LabelsPerColumnTextBox.Text = printPageSettings.LabelsPerColumn.ToString();
                LabelMarginXTextBox.Text = printPageSettings.MarginX.ToString();
                LabelMarginYTextBox.Text = printPageSettings.MarginY.ToString();
                LabelIntervalXTextBox.Text = printPageSettings.IntervalX.ToString();
                LabelIntervalYTextBox.Text = printPageSettings.IntervalY.ToString();
                HeaderFormatTextBox.Text = printPageSettings.HeaderTextFormat;
                HeaderPositionXTextBox.Text = printPageSettings.HeaderPositionX.ToString();
                HeaderPositionYTextBox.Text = printPageSettings.HeaderPositionY.ToString();

                if (printPageSettings.HeaderFont != null) {
                    HeaderFontDialog.Font = printPageSettings.HeaderFont;
                    HeaderFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
                }
            }
        }

        private bool ValidateInput() {
            // double 型の検証
            var doubleTextBoxes = new Dictionary<TextBox, string>
            {
                { LabelHeightTextBox, "ラベル縦" },
                { LabelWidthTextBox, "ラベル幅" },
                { BarcodeHeightTextBox, "バーコード縦" },
                { BarcodeWidthTextBox, "バーコード横" },
                { LabelMarginXTextBox, "マージンX" },
                { LabelMarginYTextBox, "マージンY" },
                { LabelIntervalXTextBox, "間隔X" },
                { LabelIntervalYTextBox, "間隔Y" },
                { BarcodePositionXTextBox, "バーコード位置X" },
                { BarcodePositionYTextBox, "バーコード位置Y" },
                { TextPositionXTextBox, "フォント位置X" },
                { TextPositionYTextBox, "フォント位置Y" },
                { HeaderPositionXTextBox, "ヘッダー位置X" },
                { HeaderPositionYTextBox, "ヘッダー位置Y" }
            };

            foreach (var textBox in doubleTextBoxes) {
                if (!double.TryParse(textBox.Key.Text, out var value)) {
                    MessageBox.Show($"{textBox.Value} に不正な数値が入力されました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // 範囲チェック
                if (value < 0) {
                    MessageBox.Show($"{textBox.Value} は 0 以上である必要があります。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            // int 型の検証
            var intTextBoxes = new Dictionary<TextBox, string>
            {
                { LabelsPerRowTextBox, "ラベル行数" },
                { LabelsPerColumnTextBox, "ラベル列数" },
                { CopiesPerLabelTextBox, "ラベルごとのコピー数" },
            };

            foreach (var textBox in intTextBoxes) {
                if (!int.TryParse(textBox.Key.Text, out var value)) {
                    MessageBox.Show($"{textBox.Value} に不正な整数が入力されました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                // 範囲チェック
                if (value <= 0) {
                    MessageBox.Show($"{textBox.Value} は 1 以上である必要があります。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            return true; // すべての検証が成功した場合に true を返す
        }
        private void UpdateProductPrintSettings() {
            switch (serialType) {
                case "Label":
                    if (DocumentPrintSettings.LabelPageSettings != null) {
                        DocumentPrintSettings.LabelPageSettings.LabelWidth = ParseDouble(LabelWidthTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.LabelHeight = ParseDouble(LabelHeightTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.LabelsPerRow = ParseInt(LabelsPerRowTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.LabelsPerColumn = ParseInt(LabelsPerColumnTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.MarginX = ParseDouble(LabelMarginXTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.MarginY = ParseDouble(LabelMarginYTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.IntervalX = ParseDouble(LabelIntervalXTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.IntervalY = ParseDouble(LabelIntervalYTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.HeaderPositionX = ParseDouble(HeaderPositionXTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.HeaderPositionY = ParseDouble(HeaderPositionYTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.HeaderTextFormat = HeaderFormatTextBox.Text;
                        DocumentPrintSettings.LabelPageSettings.HeaderFont = HeaderFontDialog.Font;
                    }
                    if (DocumentPrintSettings.LabelLayoutSettings != null) {
                        DocumentPrintSettings.LabelLayoutSettings.TextFormat = FormatTextBox.Text;
                        DocumentPrintSettings.LabelLayoutSettings.TextFont = FontDialog.Font;
                        DocumentPrintSettings.LabelLayoutSettings.TextPositionX = ParseDouble(TextPositionXTextBox.Text);
                        DocumentPrintSettings.LabelLayoutSettings.TextPositionY = ParseDouble(TextPositionYTextBox.Text);
                        DocumentPrintSettings.LabelLayoutSettings.AlignTextCenterX = AlignTextCenterXCheckBox.Checked;
                        DocumentPrintSettings.LabelLayoutSettings.AlignTextCenterY = AlignTextCenterYCheckBox.Checked;
                        DocumentPrintSettings.LabelLayoutSettings.CopiesPerLabel = ParseInt(CopiesPerLabelTextBox.Text);
                    }
                    break;

                case "Barcode":
                    if (DocumentPrintSettings.BarcodePageSettings != null) {
                        DocumentPrintSettings.BarcodePageSettings.LabelWidth = ParseDouble(LabelWidthTextBox.Text);
                        DocumentPrintSettings.BarcodePageSettings.LabelHeight = ParseDouble(LabelHeightTextBox.Text);
                        DocumentPrintSettings.BarcodePageSettings.LabelsPerRow = ParseInt(LabelsPerRowTextBox.Text);
                        DocumentPrintSettings.BarcodePageSettings.LabelsPerColumn = ParseInt(LabelsPerColumnTextBox.Text);
                        DocumentPrintSettings.BarcodePageSettings.MarginX = ParseDouble(LabelMarginXTextBox.Text);
                        DocumentPrintSettings.BarcodePageSettings.MarginY = ParseDouble(LabelMarginYTextBox.Text);
                        DocumentPrintSettings.BarcodePageSettings.IntervalX = ParseDouble(LabelIntervalXTextBox.Text);
                        DocumentPrintSettings.BarcodePageSettings.IntervalY = ParseDouble(LabelIntervalYTextBox.Text);
                        DocumentPrintSettings.BarcodePageSettings.HeaderPositionX = ParseDouble(HeaderPositionXTextBox.Text);
                        DocumentPrintSettings.BarcodePageSettings.HeaderPositionY = ParseDouble(HeaderPositionYTextBox.Text);
                        DocumentPrintSettings.BarcodePageSettings.HeaderTextFormat = HeaderFormatTextBox.Text;
                        DocumentPrintSettings.BarcodePageSettings.HeaderFont = HeaderFontDialog.Font;
                    }
                    if (DocumentPrintSettings.BarcodeLayoutSettings != null) {
                        DocumentPrintSettings.BarcodeLayoutSettings.BarcodeHeight = ParseDouble(BarcodeHeightTextBox.Text);
                        DocumentPrintSettings.BarcodeLayoutSettings.BarcodeWidth = ParseDouble(BarcodeWidthTextBox.Text);
                        DocumentPrintSettings.BarcodeLayoutSettings.BarcodePositionX = ParseDouble(BarcodePositionXTextBox.Text);
                        DocumentPrintSettings.BarcodeLayoutSettings.BarcodePositionY = ParseDouble(BarcodePositionYTextBox.Text);
                        DocumentPrintSettings.BarcodeLayoutSettings.AlignBarcodeCenterX = AlignBarcodeCenterXCheckBox.Checked;
                        DocumentPrintSettings.BarcodeLayoutSettings.TextFormat = FormatTextBox.Text;
                        DocumentPrintSettings.BarcodeLayoutSettings.TextFont = FontDialog.Font;
                        DocumentPrintSettings.BarcodeLayoutSettings.TextPositionX = ParseDouble(TextPositionXTextBox.Text);
                        DocumentPrintSettings.BarcodeLayoutSettings.TextPositionY = ParseDouble(TextPositionYTextBox.Text);
                        DocumentPrintSettings.BarcodeLayoutSettings.AlignTextCenterX = AlignTextCenterXCheckBox.Checked;
                        DocumentPrintSettings.BarcodeLayoutSettings.AlignTextCenterY = AlignTextCenterYCheckBox.Checked;
                        DocumentPrintSettings.BarcodeLayoutSettings.CopiesPerLabel = ParseInt(CopiesPerLabelTextBox.Text);
                    }
                    break;

                case "Substrate":
                    if (DocumentPrintSettings.LabelPageSettings != null) {
                        DocumentPrintSettings.LabelPageSettings.LabelWidth = ParseDouble(LabelWidthTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.LabelHeight = ParseDouble(LabelHeightTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.LabelsPerRow = ParseInt(LabelsPerRowTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.LabelsPerColumn = ParseInt(LabelsPerColumnTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.MarginX = ParseDouble(LabelMarginXTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.MarginY = ParseDouble(LabelMarginYTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.IntervalX = ParseDouble(LabelIntervalXTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.IntervalY = ParseDouble(LabelIntervalYTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.HeaderPositionX = ParseDouble(HeaderPositionXTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.HeaderPositionY = ParseDouble(HeaderPositionYTextBox.Text);
                        DocumentPrintSettings.LabelPageSettings.HeaderTextFormat = HeaderFormatTextBox.Text;
                        DocumentPrintSettings.LabelPageSettings.HeaderFont = HeaderFontDialog.Font;
                    }
                    if (DocumentPrintSettings.LabelLayoutSettings != null) {
                        DocumentPrintSettings.LabelLayoutSettings.TextFormat = FormatTextBox.Text;
                        DocumentPrintSettings.LabelLayoutSettings.TextFont = FontDialog.Font;
                        DocumentPrintSettings.LabelLayoutSettings.TextPositionX = ParseDouble(TextPositionXTextBox.Text);
                        DocumentPrintSettings.LabelLayoutSettings.TextPositionY = ParseDouble(TextPositionYTextBox.Text);
                        DocumentPrintSettings.LabelLayoutSettings.AlignTextCenterX = AlignTextCenterXCheckBox.Checked;
                        DocumentPrintSettings.LabelLayoutSettings.AlignTextCenterY = AlignTextCenterYCheckBox.Checked;
                        DocumentPrintSettings.LabelLayoutSettings.CopiesPerLabel = ParseInt(CopiesPerLabelTextBox.Text);
                    }
                    break;

                default:
                    MessageBox.Show("不正なシリアルタイプです。");
                    return;
            }
        }
        private static double ParseDouble(string text) {
            if (double.TryParse(text, out var value)) {
                return value;
            }
            else {
                MessageBox.Show($"不正な数値が入力されました: {text}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }
        private static int ParseInt(string text) {
            if (int.TryParse(text, out var value)) {
                return value;
            }
            else {
                MessageBox.Show($"不正な数値が入力されました: {text}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }
        private void SaveProductPrintSettings() {
            try {
                DocumentPrintSettings.SetSettingsType(ProductInfo.IsLabelPrint, ProductInfo.IsBarcodePrint);
                // JsonSerializerOptions のインスタンスをキャッシュ
                _jsonSerializerOptions ??= new JsonSerializerOptions {
                    WriteIndented = true,
                    PropertyNamingPolicy = null,
                    Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                };
                var jsonString = JsonSerializer.Serialize(DocumentPrintSettings, _jsonSerializerOptions);
                File.WriteAllText(_documentPrintSettingFilePath, jsonString);
                DialogResult = DialogResult.OK;
                Close();
            } catch (Exception ex) {
                MessageBox.Show($"設定の保存中にエラーが発生しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
            }
        }

        private void ProductBarcodePrintSetting_Load(object sender, EventArgs e) { LoadSettings(); }
        private void BarcodeCenterCheckBox_CheckedChanged(object sender, EventArgs e) { BarcodePositionXTextBox.Enabled = !AlignBarcodeCenterXCheckBox.Checked; }
        private void AlignTextXCenterCheckBox_CheckedChanged(object sender, EventArgs e) { TextPositionXTextBox.Enabled = !AlignTextCenterXCheckBox.Checked; }
        private void AlignTextYCenterCheckBox_CheckedChanged(object sender, EventArgs e) { TextPositionYTextBox.Enabled = !AlignTextCenterYCheckBox.Checked; }
        private void BarcodeHeaderFontButton_Click(object sender, EventArgs e) {
            var r = HeaderFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }
            HeaderFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
        }
        private void BarcodeFontButton_Click(object sender, EventArgs e) {
            var r = FontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }
            FontTextBox.Text = $"{FontDialog.Font.Name} {FontDialog.Font.SizeInPoints}pt";
        }
        private void OKButton_Click(object sender, EventArgs e) {
            if (!ValidateInput()) {
                return;
            }
            UpdateProductPrintSettings();
            SaveProductPrintSettings();
        }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }

    }
}
