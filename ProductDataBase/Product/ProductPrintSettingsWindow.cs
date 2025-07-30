using ProductDatabase.Product;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class ProductPrintSettingsWindow : Form {

        public ProductPrintSettings ProductPrintSettings { get; set; } = new ProductPrintSettings();
        private string _productPrintSettingFilePath = string.Empty;
        private JsonSerializerOptions? _jsonSerializerOptions;

        public ProductInformation ProductInfo { get; set; } = new ProductInformation();

        public ProductPrintSettingsWindow() {
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

            SetPageSettings(ProductPrintSettings.LabelPageSettings);
            SetLabelSettings(ProductPrintSettings.LabelLayoutSettings);
        }

        private void SetLabelSettings(Product.LabelLayoutSettings? labelSettings) {
            if (labelSettings != null) {
                CopiesPerLabelTextBox.Text = labelSettings.CopiesPerLabel.ToString();
                LabelFormatTextBox.Text = labelSettings.TextFormat;

                if (labelSettings.TextFont != null) {
                    LabelFontDialog.Font = labelSettings.TextFont;
                    PrintTextFontTextBox.Text = $"{LabelFontDialog.Font.Name} {LabelFontDialog.Font.SizeInPoints}pt";
                }

                LabelTextPositionXTextBox.Text = labelSettings.TextPositionX.ToString();
                LabelTextPositionYTextBox.Text = labelSettings.TextPositionY.ToString();
                AlignTextXCenterCheckBox.Checked = labelSettings.AlignTextCenterX;
                AlignTextYCenterCheckBox.Checked = labelSettings.AlignTextCenterY;
            }
            LabelTextPositionXTextBox.Enabled = !AlignTextXCenterCheckBox.Checked;
        }
        private void SetPageSettings(Product.LabelPageSettings? pageSettings) {
            if (pageSettings != null) {
                LabelWidthTextBox.Text = pageSettings.LabelWidth.ToString();
                LabelHeightTextBox.Text = pageSettings.LabelHeight.ToString();
                LabelsPerColumnTextBox.Text = pageSettings.LabelsPerColumn.ToString();
                LabelsPerRowTextBox.Text = pageSettings.LabelsPerRow.ToString();
                LabelMarginXTextBox.Text = pageSettings.MarginX.ToString();
                LabelMarginYTextBox.Text = pageSettings.MarginY.ToString();
                LabelIntervalXTextBox.Text = pageSettings.IntervalX.ToString();
                LabelIntervalYTextBox.Text = pageSettings.IntervalY.ToString();
                HeaderTextTextBox.Text = pageSettings.HeaderTextFormat;
                HeaderPositionXTextBox.Text = pageSettings.HeaderPositionX.ToString();
                HeaderPositionYTextBox.Text = pageSettings.HeaderPositionY.ToString();

                if (pageSettings.HeaderFont != null) {
                    HeaderFontDialog.Font = pageSettings.HeaderFont;
                    HeaderFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
                }
            }
        }

        private bool ValidateInput() {
            // double 型の検証
            var doubleTextBoxes = new Dictionary<TextBox, string>
            {
                { LabelWidthTextBox, "ラベル幅" },
                { LabelHeightTextBox, "ラベル高さ" },
                { LabelMarginXTextBox, "マージンX" },
                { LabelMarginYTextBox, "マージンY" },
                { LabelIntervalXTextBox, "間隔X" },
                { LabelIntervalYTextBox, "間隔Y" },
                { LabelTextPositionXTextBox, "フォント位置X" },
                { LabelTextPositionYTextBox, "フォント位置Y" },
                { HeaderPositionXTextBox, "ヘッダー位置X" },
                { HeaderPositionYTextBox, "ヘッダー位置Y" },
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
            return true;
        }
        private void UpdateProductPrintSettings() {
            if (ProductPrintSettings.LabelPageSettings != null) {
                ProductPrintSettings.LabelPageSettings.LabelWidth = ParseDouble(LabelWidthTextBox.Text);
                ProductPrintSettings.LabelPageSettings.LabelHeight = ParseDouble(LabelHeightTextBox.Text);
                ProductPrintSettings.LabelPageSettings.LabelsPerRow = ParseInt(LabelsPerRowTextBox.Text);
                ProductPrintSettings.LabelPageSettings.LabelsPerColumn = ParseInt(LabelsPerColumnTextBox.Text);
                ProductPrintSettings.LabelPageSettings.MarginX = ParseDouble(LabelMarginXTextBox.Text);
                ProductPrintSettings.LabelPageSettings.MarginY = ParseDouble(LabelMarginYTextBox.Text);
                ProductPrintSettings.LabelPageSettings.IntervalX = ParseDouble(LabelIntervalXTextBox.Text);
                ProductPrintSettings.LabelPageSettings.IntervalY = ParseDouble(LabelIntervalYTextBox.Text);
                ProductPrintSettings.LabelPageSettings.HeaderPositionX = ParseDouble(HeaderPositionXTextBox.Text);
                ProductPrintSettings.LabelPageSettings.HeaderPositionY = ParseDouble(HeaderPositionYTextBox.Text);
                ProductPrintSettings.LabelPageSettings.HeaderTextFormat = HeaderTextTextBox.Text;
                ProductPrintSettings.LabelPageSettings.HeaderFont = HeaderFontDialog.Font;
            }
            if (ProductPrintSettings.LabelLayoutSettings != null) {
                ProductPrintSettings.LabelLayoutSettings.TextFormat = LabelFormatTextBox.Text;
                ProductPrintSettings.LabelLayoutSettings.TextFont = LabelFontDialog.Font;
                ProductPrintSettings.LabelLayoutSettings.TextPositionX = ParseDouble(LabelTextPositionXTextBox.Text);
                ProductPrintSettings.LabelLayoutSettings.TextPositionY = ParseDouble(LabelTextPositionYTextBox.Text);
                ProductPrintSettings.LabelLayoutSettings.AlignTextCenterX = AlignTextXCenterCheckBox.Checked;
                ProductPrintSettings.LabelLayoutSettings.AlignTextCenterY = AlignTextYCenterCheckBox.Checked;
                ProductPrintSettings.LabelLayoutSettings.CopiesPerLabel = ParseInt(CopiesPerLabelTextBox.Text);
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

        private void ProductPrintSetting_Load(object sender, EventArgs e) { LoadSettings(); }
        private void TextXCenterCheckBox_CheckedChanged(object sender, EventArgs e) { LabelTextPositionXTextBox.Enabled = !AlignTextXCenterCheckBox.Checked; }
        private void TextYCenterCheckBox_CheckedChanged(object sender, EventArgs e) { LabelTextPositionYTextBox.Enabled = !AlignTextYCenterCheckBox.Checked; }
        private void HeaderFontButton_Click(object sender, EventArgs e) {
            var r = HeaderFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }

            HeaderFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
        }
        private void PrintTextFontButton_Click(object sender, EventArgs e) {
            var r = LabelFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }

            PrintTextFontTextBox.Text = $"{LabelFontDialog.Font.Name} {LabelFontDialog.Font.SizeInPoints}pt";
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
