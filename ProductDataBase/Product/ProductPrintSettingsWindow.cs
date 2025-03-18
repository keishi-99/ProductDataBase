using ProductDatabase.Product;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace ProductDatabase {
    public partial class ProductPrintSettingsWindow : Form {

        public ProductPrintSettings ProductPrintSettings { get; set; } = new ProductPrintSettings();
        private string _productPrintSettingFilePath = string.Empty;

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

        private void SetLabelSettings(Product.LabelLayoutSettings labelSettings) {
            CopiesPerLabelTextBox.Text = labelSettings.CopiesPerLabel.ToString();
            LabelFormatTextBox.Text = labelSettings.Format;

            if (labelSettings.TextFont != null) {
                LabelFontDialog.Font = labelSettings.TextFont;
                PrintTextFontTextBox.Text = $"{LabelFontDialog.Font.Name} {LabelFontDialog.Font.SizeInPoints}pt";
            }

            LabelTextPostionXTextBox.Text = labelSettings.TextPositionX.ToString();
            LabelTextPostionYTextBox.Text = labelSettings.TextPositionY.ToString();
            AlignTextXCenterCheckBox.Checked = labelSettings.AlignTextXCenter;
            AlignTextYCenterCheckBox.Checked = labelSettings.AlignTextYCenter;

            LabelTextPostionXTextBox.Enabled = !AlignTextXCenterCheckBox.Checked;
        }
        private void SetPageSettings(Product.LabelPageSettings pageSettings) {
            LabelWidthTextBox.Text = pageSettings.LabelWidth.ToString();
            LabelHeightTextBox.Text = pageSettings.LabelHeight.ToString();
            LabelsPerColumnTextBox.Text = pageSettings.LabelsPerColumn.ToString();
            LabelsPerRowTextBox.Text = pageSettings.LabelsPerRow.ToString();
            LabelMarginXTextBox.Text = pageSettings.MarginX.ToString();
            LabelMarginYTextBox.Text = pageSettings.MarginY.ToString();
            LabelIntervalXTextBox.Text = pageSettings.IntervalX.ToString();
            LabelIntervalYTextBox.Text = pageSettings.IntervalY.ToString();
            HeaderTextTextBox.Text = pageSettings.HeaderText;
            HeaderPostionXTextBox.Text = pageSettings.HeaderPositionX.ToString();
            HeaderPostionYTextBox.Text = pageSettings.HeaderPositionY.ToString();

            if (pageSettings.HeaderFont != null) {
                HeaderFontDialog.Font = pageSettings.HeaderFont;
                HeaderFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
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
                { LabelTextPostionXTextBox, "フォント位置X" },
                { LabelTextPostionYTextBox, "フォント位置Y" },
                { HeaderPostionXTextBox, "ヘッダー位置X" },
                { HeaderPostionYTextBox, "ヘッダー位置Y" },
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
            ProductPrintSettings.LabelPageSettings.LabelWidth = ParseDouble(LabelWidthTextBox.Text);
            ProductPrintSettings.LabelPageSettings.LabelHeight = ParseDouble(LabelHeightTextBox.Text);
            ProductPrintSettings.LabelPageSettings.LabelsPerRow = ParseInt(LabelsPerRowTextBox.Text);
            ProductPrintSettings.LabelPageSettings.LabelsPerColumn = ParseInt(LabelsPerColumnTextBox.Text);
            ProductPrintSettings.LabelPageSettings.MarginX = ParseDouble(LabelMarginXTextBox.Text);
            ProductPrintSettings.LabelPageSettings.MarginY = ParseDouble(LabelMarginYTextBox.Text);
            ProductPrintSettings.LabelPageSettings.IntervalX = ParseDouble(LabelIntervalXTextBox.Text);
            ProductPrintSettings.LabelPageSettings.IntervalY = ParseDouble(LabelIntervalYTextBox.Text);
            ProductPrintSettings.LabelPageSettings.HeaderPositionX = ParseDouble(HeaderPostionXTextBox.Text);
            ProductPrintSettings.LabelPageSettings.HeaderPositionY = ParseDouble(HeaderPostionYTextBox.Text);
            ProductPrintSettings.LabelPageSettings.HeaderText = HeaderFontTextBox.Text;
            ProductPrintSettings.LabelPageSettings.HeaderFont = HeaderFontDialog.Font;

            ProductPrintSettings.LabelLayoutSettings.Format = LabelFormatTextBox.Text;
            ProductPrintSettings.LabelLayoutSettings.TextFont = LabelFontDialog.Font;
            ProductPrintSettings.LabelLayoutSettings.TextPositionX = ParseDouble(LabelTextPostionXTextBox.Text);
            ProductPrintSettings.LabelLayoutSettings.TextPositionY = ParseDouble(LabelTextPostionYTextBox.Text);
            ProductPrintSettings.LabelLayoutSettings.AlignTextXCenter = AlignTextXCenterCheckBox.Checked;
            ProductPrintSettings.LabelLayoutSettings.AlignTextYCenter = AlignTextYCenterCheckBox.Checked;
            ProductPrintSettings.LabelLayoutSettings.CopiesPerLabel = ParseInt(CopiesPerLabelTextBox.Text);
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
                var options = new JsonSerializerOptions {
                    WriteIndented = true,
                    PropertyNamingPolicy = null,
                    Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
                };
                var jsonString = JsonSerializer.Serialize(ProductPrintSettings, options);
                File.WriteAllText(_productPrintSettingFilePath, jsonString);
                DialogResult = DialogResult.OK;
                Close();
            } catch (Exception ex) {
                MessageBox.Show($"設定の保存中にエラーが発生しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
            }
        }

        private void ProductPrintSetting_Load(object sender, EventArgs e) { LoadSettings(); }
        private void TextXCenterCheckBox_CheckedChanged(object sender, EventArgs e) { LabelTextPostionXTextBox.Enabled = !AlignTextXCenterCheckBox.Checked; }
        private void TextYCenterCheckBox_CheckedChanged(object sender, EventArgs e) { LabelTextPostionYTextBox.Enabled = !AlignTextYCenterCheckBox.Checked; }
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
