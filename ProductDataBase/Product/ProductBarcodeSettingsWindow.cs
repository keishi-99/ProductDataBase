using ProductDatabase.Product;

namespace ProductDatabase {
    public partial class ProductBarcodeSettingsWindow : Form {

        public ProductPrintSettings ProductPrintSettings { get; set; } = new ProductPrintSettings();
        private string _productPrintSettingFilePath = string.Empty;

        public ProductBarcodeSettingsWindow() {
            InitializeComponent();
        }

        private void LoadSettings() {
            if (Owner is ProductRegistration2Window productWindow) {
                LoadSettingsFromWindow(productWindow.ProductPrintSettings, productWindow.labelSettingFilePath);
            }
            else if (Owner is RePrintWindow rePrintWindow) {
                LoadSettingsFromWindow(rePrintWindow.ProductPrintSettings, rePrintWindow.labelSettingFilePath);
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

        private void SetLabelSettings(BarcodeLayoutSettings labelSettings) {
            BarcodeHeightTextBox.Text = labelSettings.BarcodeHeight.ToString();
            BarcodeMagnitudeTextBox.Text = labelSettings.BarcodeMagnitude.ToString();
            BarcodeCopiesPerLabelTextBox.Text = labelSettings.CopiesPerLabel.ToString();
            BarcodeFormatTextBox.Text = labelSettings.Format;

            if (labelSettings.TextFont != null) {
                BarcodeFontDialog.Font = labelSettings.TextFont;
                BarcodeFontTextBox.Text = $"{BarcodeFontDialog.Font.Name} {BarcodeFontDialog.Font.SizeInPoints}pt";
            }

            BarcodePostionXTextBox.Text = labelSettings.BarcodePositionX.ToString();
            BarcodePostionYTextBox.Text = labelSettings.BarcodePositionY.ToString();
            BarcodeCenterCheckBox.Checked = labelSettings.AlignBarcodeXCenter;
            BarcodePostionXTextBox.Enabled = !BarcodeCenterCheckBox.Checked;

            FontPostionXTextBox.Text = labelSettings.TextPositionX.ToString();
            FontPostionYTextBox.Text = labelSettings.TextPositionY.ToString();
            FontCenterCheckBox.Checked = labelSettings.AlignTextXCenter;
            FontPostionXTextBox.Enabled = !FontCenterCheckBox.Checked;
        }
        private void SetPageSettings(BarcodePageSettings pageSettings) {
            BarcodeLabelWidthTextBox.Text = pageSettings.LabelWidth.ToString();
            BarcodeLabelHeightTextBox.Text = pageSettings.LabelHeight.ToString();
            BarcodeLabelsPerRowTextBox.Text = pageSettings.LabelsPerRow.ToString();
            BarcodeLabelsPerColumnTextBox.Text = pageSettings.LabelsPerColumn.ToString();
            BarcodeLabelMarginXTextBox.Text = pageSettings.MarginX.ToString();
            BarcodeLabelMarginYTextBox.Text = pageSettings.MarginY.ToString();
            BarcodeLabelIntervalXTextBox.Text = pageSettings.IntervalX.ToString();
            BarcodeLabelIntervalYTextBox.Text = pageSettings.IntervalY.ToString();
            BarcodeLabelHeaderTextBox.Text = pageSettings.HeaderText;
            BarcodeLabelHeaderPostionXTextBox.Text = pageSettings.HeaderPositionX.ToString();
            BarcodeLabelHeaderPostionYTextBox.Text = pageSettings.HeaderPositionY.ToString();

            if (pageSettings.HeaderFont != null) {
                BarcodeLabelHeaderFontDialog.Font = pageSettings.HeaderFont;
                BarcodeLabelHeaderFontTextBox.Text = $"{BarcodeLabelHeaderFontDialog.Font.Name} {BarcodeLabelHeaderFontDialog.Font.SizeInPoints}pt";
            }
        }

        private bool ValidateInput() {
            // double 型の検証
            var doubleTextBoxes = new Dictionary<TextBox, string>
            {
                { BarcodeLabelWidthTextBox, "ラベル幅" },
                { BarcodeHeightTextBox, "ラベル高さ" },
                { BarcodeLabelMarginXTextBox, "マージンX" },
                { BarcodeLabelMarginYTextBox, "マージンY" },
                { BarcodeLabelIntervalXTextBox, "間隔X" },
                { BarcodeLabelIntervalYTextBox, "間隔Y" },
                { BarcodePostionXTextBox, "バーコード位置X" },
                { BarcodePostionYTextBox, "バーコード位置Y" },
                { BarcodeMagnitudeTextBox, "バーコード倍率" },
                { FontPostionXTextBox, "フォント位置X" },
                { FontPostionYTextBox, "フォント位置Y" }
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
                { BarcodeLabelHeaderPostionXTextBox, "ヘッダー位置X" },
                { BarcodeLabelHeaderPostionYTextBox, "ヘッダー位置Y" }
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
            ProductPrintSettings.BarcodePageSettings.HeaderText = BarcodeLabelHeaderTextBox.Text;
            ProductPrintSettings.BarcodePageSettings.HeaderFont = BarcodeLabelHeaderFontDialog.Font;

            ProductPrintSettings.BarcodeLayoutSettings.BarcodeHeight = ParseDouble(BarcodeHeightTextBox.Text);
            ProductPrintSettings.BarcodeLayoutSettings.BarcodePositionX = ParseDouble(BarcodePostionXTextBox.Text);
            ProductPrintSettings.BarcodeLayoutSettings.BarcodePositionY = ParseDouble(BarcodePostionYTextBox.Text);
            ProductPrintSettings.BarcodeLayoutSettings.AlignBarcodeXCenter = BarcodeCenterCheckBox.Checked;
            ProductPrintSettings.BarcodeLayoutSettings.BarcodeMagnitude = ParseDouble(BarcodeMagnitudeTextBox.Text);
            ProductPrintSettings.BarcodeLayoutSettings.Format = BarcodeFormatTextBox.Text;
            ProductPrintSettings.BarcodeLayoutSettings.TextFont = BarcodeFontDialog.Font;
            ProductPrintSettings.BarcodeLayoutSettings.TextPositionX = ParseDouble(FontPostionXTextBox.Text);
            ProductPrintSettings.BarcodeLayoutSettings.TextPositionY = ParseDouble(FontPostionYTextBox.Text);
            ProductPrintSettings.BarcodeLayoutSettings.AlignTextXCenter = FontCenterCheckBox.Checked;
            ProductPrintSettings.BarcodeLayoutSettings.CopiesPerLabel = ParseInt(BarcodeCopiesPerLabelTextBox.Text);
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
                using (var swBarcode = new StreamWriter(_productPrintSettingFilePath, false, System.Text.Encoding.UTF8)) {
                    var serializerBarcode = new System.Xml.Serialization.XmlSerializer(typeof(ProductPrintSettings));
                    serializerBarcode.Serialize(swBarcode, ProductPrintSettings);
                }
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
