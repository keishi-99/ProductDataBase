using ProductDatabase.Product;

namespace ProductDatabase {
    public partial class ProductBarcodeSettingsWindow : Form {

        private readonly CBarcodeProLabelSettings _barcodeProLabelSettings;
        private readonly CBarcodeProPageSettings _barcodeProPageSettings;

        public ProductBarcodeSettingsWindow() {
            InitializeComponent();

            _barcodeProLabelSettings = new CBarcodeProLabelSettings();
            _barcodeProPageSettings = new CBarcodeProPageSettings();
        }

        public CSettingsBarcodePro SettingsBarcodePro { get; set; } = new CSettingsBarcodePro();
        private string _barcodeSettingFilePath = string.Empty;

        private void PageSettingsBarcodeLoad(object sender, EventArgs e) {
            if (Owner is ProductRegistration2Window productWindow) {
                LoadSettingsFromProductRegistration2Window(productWindow);
            }
            else if (Owner is RePrintWindow rePrintWindow) {
                LoadSettingsFromRePaintWindow(rePrintWindow);
            }
            else {
                MessageBox.Show("この画面を開くには正しいウィンドウから開いてください。");
                return;
            }
        }
        private void LoadSettingsFromProductRegistration2Window(ProductRegistration2Window window) {
            var pageSettings = window.SettingsBarcodePro.BarcodeProPageSettings;
            var labelSettings = window.SettingsBarcodePro.BarcodeProLabelSettings;
            _barcodeSettingFilePath = window.barcodeSettingFilePath;

            SetPageSettings(pageSettings);
            SetLabelSettings(labelSettings);
        }
        private void LoadSettingsFromRePaintWindow(RePrintWindow window) {
            var pageSettings = window.SettingsBarcodePro.BarcodeProPageSettings;
            var labelSettings = window.SettingsBarcodePro.BarcodeProLabelSettings;
            _barcodeSettingFilePath = window.barcodeSettingFilePath;

            SetPageSettings(pageSettings);
            SetLabelSettings(labelSettings);
        }
        private void SetLabelSettings(CBarcodeProLabelSettings labelSettings) {
            BarcodeHeightTextBox.Text = labelSettings.BarcodeHeight.ToString();
            BarcodeMagnitudeTextBox.Text = labelSettings.BarcodeMagnitude.ToString();
            BarcodeQuantityTextBox.Text = labelSettings.NumLabels.ToString();
            BarcodeFormatTextBox.Text = labelSettings.Format;

            if (labelSettings.Font != null) {
                BarcodeFontDialog.Font = labelSettings.Font;
                BarcodeFontTextBox.Text = $"{BarcodeFontDialog.Font.Name} {BarcodeFontDialog.Font.SizeInPoints}pt";
            }

            BarcodePostionXTextBox.Text = labelSettings.BarcodePosX.ToString();
            BarcodePostionYTextBox.Text = labelSettings.BarcodePosY.ToString();
            BarcodeCenterCheckBox.Checked = labelSettings.AlignBarcodeCenter;
            BarcodePostionXTextBox.Enabled = !BarcodeCenterCheckBox.Checked;

            FontPostionXTextBox.Text = labelSettings.StringPosX.ToString();
            FontPostionYTextBox.Text = labelSettings.StringPosY.ToString();
            FontCenterCheckBox.Checked = labelSettings.AlignStringCenter;
            FontPostionXTextBox.Enabled = !FontCenterCheckBox.Checked;
        }
        private void SetPageSettings(CBarcodeProPageSettings pageSettings) {
            BarcodeLabelWidthTextBox.Text = pageSettings.SizeX.ToString();
            BarcodeLabelHeightTextBox.Text = pageSettings.SizeY.ToString();
            BarcodeQuantityXTextBox.Text = pageSettings.NumLabelsX.ToString();
            BarcodeQuantityYTextBox.Text = pageSettings.NumLabelsY.ToString();
            BarcodePageOffsetXTextBox.Text = pageSettings.OffsetX.ToString();
            BarcodePageOffsetYTextBox.Text = pageSettings.OffsetY.ToString();
            BarcodeLabelIntervalXTextBox.Text = pageSettings.IntervalX.ToString();
            BarcodeLabelIntervalYTextBox.Text = pageSettings.IntervalY.ToString();
            BarcodeHeaderStringTextBox.Text = pageSettings.HeaderString;
            BarcodeHeaderPostionXTextBox.Text = pageSettings.HeaderPos.X.ToString();
            BarcodeHeaderPostionYTextBox.Text = pageSettings.HeaderPos.Y.ToString();

            if (pageSettings.HeaderFooterFont != null) {
                BarcodeHeaderFontDialog.Font = pageSettings.HeaderFooterFont;
                BarcodeHeaderFooterFontTextBox.Text = $"{BarcodeHeaderFontDialog.Font.Name} {BarcodeHeaderFontDialog.Font.SizeInPoints}pt";
            }
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            double sizeX, sizeY, offsetX, offsetY, intervalX, intervalY, barcodePosX, barcodePosY, barcodeHeight, barcodeMagnitude, stringPosX, stringPosY;
            int numLabelsX, numLabelsY, headerPosX, headerPosY, numLabels;

            try {
                sizeX = double.Parse(BarcodeLabelWidthTextBox.Text);
                sizeY = double.Parse(BarcodeLabelHeightTextBox.Text);
                numLabelsX = int.Parse(BarcodeQuantityXTextBox.Text);
                numLabelsY = int.Parse(BarcodeQuantityYTextBox.Text);
                offsetX = double.Parse(BarcodePageOffsetXTextBox.Text);
                offsetY = double.Parse(BarcodePageOffsetYTextBox.Text);
                intervalX = double.Parse(BarcodeLabelIntervalXTextBox.Text);
                intervalY = double.Parse(BarcodeLabelIntervalYTextBox.Text);
                headerPosX = int.Parse(BarcodeHeaderPostionXTextBox.Text);
                headerPosY = int.Parse(BarcodeHeaderPostionYTextBox.Text);

                barcodeHeight = double.Parse(BarcodeHeightTextBox.Text);
                barcodePosX = double.Parse(BarcodePostionXTextBox.Text);
                barcodePosY = double.Parse(BarcodePostionYTextBox.Text);
                barcodeMagnitude = double.Parse(BarcodeMagnitudeTextBox.Text);
                stringPosX = double.Parse(FontPostionXTextBox.Text);
                stringPosY = double.Parse(FontPostionYTextBox.Text);
                numLabels = int.Parse(BarcodeQuantityTextBox.Text);
            } catch (Exception ex) {
                MessageBox.Show($"入力値が不正です。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
                return;
            }

            _barcodeProPageSettings.SizeX = sizeX;
            _barcodeProPageSettings.SizeY = sizeY;
            _barcodeProPageSettings.NumLabelsX = numLabelsX;
            _barcodeProPageSettings.NumLabelsY = numLabelsY;
            _barcodeProPageSettings.OffsetX = offsetX;
            _barcodeProPageSettings.OffsetY = offsetY;
            _barcodeProPageSettings.IntervalX = intervalX;
            _barcodeProPageSettings.IntervalY = intervalY;
            _barcodeProPageSettings.HeaderPos = new Point(headerPosX, headerPosY);
            _barcodeProPageSettings.HeaderString = BarcodeHeaderStringTextBox.Text;
            _barcodeProPageSettings.HeaderFooterFont = BarcodeHeaderFontDialog.Font;

            _barcodeProLabelSettings.BarcodeHeight = barcodeHeight;
            _barcodeProLabelSettings.BarcodePosX = barcodePosX;
            _barcodeProLabelSettings.BarcodePosY = barcodePosY;
            _barcodeProLabelSettings.AlignBarcodeCenter = BarcodeCenterCheckBox.Checked;
            _barcodeProLabelSettings.BarcodeMagnitude = barcodeMagnitude;
            _barcodeProLabelSettings.Format = BarcodeFormatTextBox.Text;
            _barcodeProLabelSettings.Font = BarcodeFontDialog.Font;
            _barcodeProLabelSettings.StringPosX = stringPosX;
            _barcodeProLabelSettings.StringPosY = stringPosY;
            _barcodeProLabelSettings.AlignStringCenter = FontCenterCheckBox.Checked;
            _barcodeProLabelSettings.NumLabels = int.Parse(BarcodeQuantityTextBox.Text);

            SettingsBarcodePro.BarcodeProPageSettings = _barcodeProPageSettings;
            SettingsBarcodePro.BarcodeProLabelSettings = _barcodeProLabelSettings;

            try {
                using var swBarcode = new StreamWriter(_barcodeSettingFilePath, false, new System.Text.UTF8Encoding(false));
                var serializerBarcode = new System.Xml.Serialization.XmlSerializer(typeof(CSettingsBarcodePro));
                serializerBarcode.Serialize(swBarcode, SettingsBarcodePro);
            } catch (Exception ex) {
                MessageBox.Show($"設定の保存中にエラーが発生しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
        private void BtnHeaderFooterFont_Click(object sender, EventArgs e) {
            var r = BarcodeHeaderFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }

            BarcodeHeaderFooterFontTextBox.Text = $"{BarcodeHeaderFontDialog.Font.Name} {BarcodeHeaderFontDialog.Font.SizeInPoints}pt";
        }
        private void BtnBarcodeFont_Click(object sender, EventArgs e) {
            var r = BarcodeFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }

            BarcodeFontTextBox.Text = $"{BarcodeFontDialog.Font.Name} {BarcodeFontDialog.Font.SizeInPoints}pt";
        }
        private void BarcodeCenterCheckBox_CheckedChanged(object sender, EventArgs e) {
            BarcodePostionXTextBox.Enabled = !BarcodeCenterCheckBox.Checked;
        }
        private void ProductBarcodePrintSetting_Load(object sender, EventArgs e) { PageSettingsBarcodeLoad(sender, e); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
        private void FontCenterCheckBox_CheckedChanged(object sender, EventArgs e) {
            FontPostionXTextBox.Enabled = !FontCenterCheckBox.Checked;

        }
    }
}
