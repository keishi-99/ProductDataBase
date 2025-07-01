using ProductDatabase.Product;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class ProductBarcodeSettingsWindow : Form {

        public ProductPrintSettings ProductPrintSettings { get; set; } = new ProductPrintSettings();
        private string _productPrintSettingFilePath = string.Empty;
        private JsonSerializerOptions? _jsonSerializerOptions;

        public ProductInformation ProductInfo { get; set; } = new ProductInformation();

        public ProductBarcodeSettingsWindow() {
            InitializeComponent();
        }

        private void LoadSettings() {
            if (Owner is ProductRegistration2Window productWindow) {
                LoadSettingsFromWindow(productWindow.ProductPrintSettings, productWindow.printSettingPath);
            }
            else if (Owner is RePrintWindow rePrintWindow) {
                LoadSettingsFromWindow(rePrintWindow.ProductPrintSettings, rePrintWindow.printSettingPath);
            }
            else {
                MessageBox.Show("この画面を開くには正しいウィンドウから開いてください。");
                Close();
            }
        }
        private void LoadSettingsFromWindow(ProductPrintSettings settings, string filePath) {
            ProductPrintSettings = settings;
            _productPrintSettingFilePath = filePath;

            SetPageSettings(ProductPrintSettings.BarcodePageSettings);
            SetLabelSettings(ProductPrintSettings.BarcodeLayoutSettings);
        }

        private void SetLabelSettings(BarcodeLayoutSettings? labelSettings) {
            if (labelSettings != null) {
                BarcodeHeightTextBox.Text = labelSettings.BarcodeHeight.ToString();
                BarcodeWidthTextBox.Text = labelSettings.BarcodeWidth.ToString();
                BarcodeCopiesPerLabelTextBox.Text = labelSettings.CopiesPerLabel.ToString();
                BarcodeFormatTextBox.Text = labelSettings.TextFormat;

                if (labelSettings.TextFont != null) {
                    BarcodeFontDialog.Font = labelSettings.TextFont;
                    BarcodeFontTextBox.Text = $"{BarcodeFontDialog.Font.Name} {BarcodeFontDialog.Font.SizeInPoints}pt";
                }

                BarcodePostionXTextBox.Text = labelSettings.BarcodePositionX.ToString();
                BarcodePostionYTextBox.Text = labelSettings.BarcodePositionY.ToString();
                BarcodeCenterCheckBox.Checked = labelSettings.AlignBarcodeCenterX;

                FontPostionXTextBox.Text = labelSettings.TextPositionX.ToString();
                FontPostionYTextBox.Text = labelSettings.TextPositionY.ToString();
                FontCenterCheckBox.Checked = labelSettings.AlignTextCenterX;
            }
            BarcodePostionXTextBox.Enabled = !BarcodeCenterCheckBox.Checked;
            FontPostionXTextBox.Enabled = !FontCenterCheckBox.Checked;
        }
        private void SetPageSettings(BarcodePageSettings? pageSettings) {
            if (pageSettings != null) {
                BarcodeLabelWidthTextBox.Text = pageSettings.LabelWidth.ToString();
                BarcodeLabelHeightTextBox.Text = pageSettings.LabelHeight.ToString();
                BarcodeLabelsPerRowTextBox.Text = pageSettings.LabelsPerRow.ToString();
                BarcodeLabelsPerColumnTextBox.Text = pageSettings.LabelsPerColumn.ToString();
                BarcodeLabelMarginXTextBox.Text = pageSettings.MarginX.ToString();
                BarcodeLabelMarginYTextBox.Text = pageSettings.MarginY.ToString();
                BarcodeLabelIntervalXTextBox.Text = pageSettings.IntervalX.ToString();
                BarcodeLabelIntervalYTextBox.Text = pageSettings.IntervalY.ToString();
                BarcodeLabelHeaderTextBox.Text = pageSettings.HeaderTextFormat;
                BarcodeLabelHeaderPostionXTextBox.Text = pageSettings.HeaderPositionX.ToString();
                BarcodeLabelHeaderPostionYTextBox.Text = pageSettings.HeaderPositionY.ToString();

                if (pageSettings.HeaderFont != null) {
                    BarcodeLabelHeaderFontDialog.Font = pageSettings.HeaderFont;
                    BarcodeLabelHeaderFontTextBox.Text = $"{BarcodeLabelHeaderFontDialog.Font.Name} {BarcodeLabelHeaderFontDialog.Font.SizeInPoints}pt";
                }
            }
        }

        private bool ValidateInput() {
            // double 型の検証
            var doubleTextBoxes = new Dictionary<TextBox, string>
            {
                { BarcodeLabelHeightTextBox, "ラベル縦" },
                { BarcodeLabelWidthTextBox, "ラベル幅" },
                { BarcodeHeightTextBox, "バーコード縦" },
                { BarcodeWidthTextBox, "バーコード横" },
                { BarcodeLabelMarginXTextBox, "マージンX" },
                { BarcodeLabelMarginYTextBox, "マージンY" },
                { BarcodeLabelIntervalXTextBox, "間隔X" },
                { BarcodeLabelIntervalYTextBox, "間隔Y" },
                { BarcodePostionXTextBox, "バーコード位置X" },
                { BarcodePostionYTextBox, "バーコード位置Y" },
                { FontPostionXTextBox, "フォント位置X" },
                { FontPostionYTextBox, "フォント位置Y" },
                { BarcodeLabelHeaderPostionXTextBox, "ヘッダー位置X" },
                { BarcodeLabelHeaderPostionYTextBox, "ヘッダー位置Y" }
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
                { BarcodeLabelsPerRowTextBox, "ラベル行数" },
                { BarcodeLabelsPerColumnTextBox, "ラベル列数" },
                { BarcodeCopiesPerLabelTextBox, "ラベルごとのコピー数" },
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
            if (ProductPrintSettings.BarcodePageSettings != null) {
                ProductPrintSettings.BarcodePageSettings.LabelWidth = ParseDouble(BarcodeLabelWidthTextBox.Text);
                ProductPrintSettings.BarcodePageSettings.LabelHeight = ParseDouble(BarcodeLabelHeightTextBox.Text);
                ProductPrintSettings.BarcodePageSettings.LabelsPerRow = ParseInt(BarcodeLabelsPerRowTextBox.Text);
                ProductPrintSettings.BarcodePageSettings.LabelsPerColumn = ParseInt(BarcodeLabelsPerColumnTextBox.Text);
                ProductPrintSettings.BarcodePageSettings.MarginX = ParseDouble(BarcodeLabelMarginXTextBox.Text);
                ProductPrintSettings.BarcodePageSettings.MarginY = ParseDouble(BarcodeLabelMarginYTextBox.Text);
                ProductPrintSettings.BarcodePageSettings.IntervalX = ParseDouble(BarcodeLabelIntervalXTextBox.Text);
                ProductPrintSettings.BarcodePageSettings.IntervalY = ParseDouble(BarcodeLabelIntervalYTextBox.Text);
                ProductPrintSettings.BarcodePageSettings.HeaderPositionX = ParseDouble(BarcodeLabelHeaderPostionXTextBox.Text);
                ProductPrintSettings.BarcodePageSettings.HeaderPositionY = ParseDouble(BarcodeLabelHeaderPostionYTextBox.Text);
                ProductPrintSettings.BarcodePageSettings.HeaderTextFormat = BarcodeLabelHeaderTextBox.Text;
                ProductPrintSettings.BarcodePageSettings.HeaderFont = BarcodeLabelHeaderFontDialog.Font;
            }
            if (ProductPrintSettings.BarcodeLayoutSettings != null) {
                ProductPrintSettings.BarcodeLayoutSettings.BarcodeHeight = ParseDouble(BarcodeHeightTextBox.Text);
                ProductPrintSettings.BarcodeLayoutSettings.BarcodeWidth = ParseDouble(BarcodeWidthTextBox.Text);
                ProductPrintSettings.BarcodeLayoutSettings.BarcodePositionX = ParseDouble(BarcodePostionXTextBox.Text);
                ProductPrintSettings.BarcodeLayoutSettings.BarcodePositionY = ParseDouble(BarcodePostionYTextBox.Text);
                ProductPrintSettings.BarcodeLayoutSettings.AlignBarcodeCenterX = BarcodeCenterCheckBox.Checked;
                ProductPrintSettings.BarcodeLayoutSettings.TextFormat = BarcodeFormatTextBox.Text;
                ProductPrintSettings.BarcodeLayoutSettings.TextFont = BarcodeFontDialog.Font;
                ProductPrintSettings.BarcodeLayoutSettings.TextPositionX = ParseDouble(FontPostionXTextBox.Text);
                ProductPrintSettings.BarcodeLayoutSettings.TextPositionY = ParseDouble(FontPostionYTextBox.Text);
                ProductPrintSettings.BarcodeLayoutSettings.AlignTextCenterX = FontCenterCheckBox.Checked;
                ProductPrintSettings.BarcodeLayoutSettings.CopiesPerLabel = ParseInt(BarcodeCopiesPerLabelTextBox.Text);
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
                ProductPrintSettings.SetSettingsType(ProductInfo.IsLabelPrint, ProductInfo.IsBarcodePrint);
                // JsonSerializerOptions のインスタンスをキャッシュ
                _jsonSerializerOptions ??= new JsonSerializerOptions {
                    WriteIndented = true,
                    PropertyNamingPolicy = null,
                    Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                };
                var jsonString = JsonSerializer.Serialize(ProductPrintSettings, _jsonSerializerOptions);
                File.WriteAllText(_productPrintSettingFilePath, jsonString);
                DialogResult = DialogResult.OK;
                Close();
            } catch (Exception ex) {
                MessageBox.Show($"設定の保存中にエラーが発生しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
            }
        }

        private void ProductBarcodePrintSetting_Load(object sender, EventArgs e) { LoadSettings(); }
        private void BarcodeCenterCheckBox_CheckedChanged(object sender, EventArgs e) { BarcodePostionXTextBox.Enabled = !BarcodeCenterCheckBox.Checked; }
        private void FontCenterCheckBox_CheckedChanged(object sender, EventArgs e) { FontPostionXTextBox.Enabled = !FontCenterCheckBox.Checked; }
        private void BarcodeHeaderFontButton_Click(object sender, EventArgs e) {
            var r = BarcodeLabelHeaderFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }
            BarcodeLabelHeaderFontTextBox.Text = $"{BarcodeLabelHeaderFontDialog.Font.Name} {BarcodeLabelHeaderFontDialog.Font.SizeInPoints}pt";
        }
        private void BarcodeFontButton_Click(object sender, EventArgs e) {
            var r = BarcodeFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }
            BarcodeFontTextBox.Text = $"{BarcodeFontDialog.Font.Name} {BarcodeFontDialog.Font.SizeInPoints}pt";
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
