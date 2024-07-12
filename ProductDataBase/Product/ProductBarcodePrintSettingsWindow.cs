using ProductDatabase.Product;

namespace ProductDatabase {
    public partial class ProductBarcodePrintSettingsWindow : Form {

        private CBarcodeProLabelSettings _barcodeProLabelSettings;
        private CBarcodeProPageSettings _barcodeProPageSettings;

        public ProductBarcodePrintSettingsWindow() {
            InitializeComponent();

            _barcodeProLabelSettings = new CBarcodeProLabelSettings();
            _barcodeProPageSettings = new CBarcodeProPageSettings();
        }

        private void PageSettingsBarcodeLoad(object sender, EventArgs e) {
            _barcodeProPageSettings = ((ProductRegistration2Window)Owner!).SettingsBarcodePro!.BarcodeProPageSettings;
            BarcodeLabelWidthTextBox.Text = _barcodeProPageSettings.SizeX.ToString();
            BarcodeLabelHeightTextBox.Text = _barcodeProPageSettings.SizeY.ToString();
            BarcodeQuantityXTextBox.Text = _barcodeProPageSettings.NumLabelsX.ToString();
            BarcodeQuantityYTextBox.Text = _barcodeProPageSettings.NumLabelsY.ToString();
            BarcodePageOffsetXTextBox.Text = _barcodeProPageSettings.OffsetX.ToString();
            BarcodePageOffsetYTextBox.Text = _barcodeProPageSettings.OffsetY.ToString();
            BarcodeLabelIntervalXTextBox.Text = _barcodeProPageSettings.IntervalX.ToString();
            BarcodeLabelIntervalYTextBox.Text = _barcodeProPageSettings.IntervalY.ToString();
            BarcodeHeaderStringTextBox.Text = _barcodeProPageSettings.HeaderString;
            BarcodeHeaderPostionXTextBox.Text = _barcodeProPageSettings.HeaderPos.X.ToString();
            BarcodeHeaderPostionYTextBox.Text = _barcodeProPageSettings.HeaderPos.Y.ToString();

            if (_barcodeProPageSettings.HeaderFooterFont != null) {
                BarcodeHeaderFontDialog.Font = _barcodeProPageSettings.HeaderFooterFont;
                BarcodeHeaderFooterFontTextBox.Text = $"{BarcodeHeaderFontDialog.Font.Name} {BarcodeHeaderFontDialog.Font.SizeInPoints}pt";
            }

            _barcodeProLabelSettings = ((ProductRegistration2Window)Owner).SettingsBarcodePro!.BarcodeProLabelSettings;
            BarcodeHeightTextBox.Text = _barcodeProLabelSettings.BarcodeHeight.ToString();
            BarcodeMagnitudeTextBox.Text = _barcodeProLabelSettings.BarcodeMagnitude.ToString();
            BarcodeQuantityTextBox.Text = _barcodeProLabelSettings.NumLabels.ToString();
            BarcodeFormatTextBox.Text = _barcodeProLabelSettings.Format;

            if (_barcodeProLabelSettings.Font != null) {
                BarcodeFontDialog.Font = _barcodeProLabelSettings.Font;
                BarcodeFontTextBox.Text = $"{BarcodeFontDialog.Font.Name} {BarcodeFontDialog.Font.SizeInPoints}pt";
            }

            BarcodePostionXTextBox.Text = _barcodeProLabelSettings.BarcodePosX.ToString();
            BarcodePostionYTextBox.Text = _barcodeProLabelSettings.BarcodePosY.ToString();
            BarcodeCenterCheckBox.Checked = _barcodeProLabelSettings.AlignBarcodeCenter;

            BarcodePostionXTextBox.Enabled = !BarcodeCenterCheckBox.Checked;

            FontPostionXTextBox.Text = _barcodeProLabelSettings.StringPosX.ToString();
            FontPostionYTextBox.Text = _barcodeProLabelSettings.StringPosY.ToString();
            FontCenterCheckBox.Checked = _barcodeProLabelSettings.AlignStringCenter;
            FontPostionXTextBox.Enabled = !FontCenterCheckBox.Checked;
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            double sizeX, sizeY, offsetX, offsetY, intervalX, intervalY, barcodePosX, barcodePosY, barcodeHeight, barcodeMagnitude, stringPosX, stringPosY;
            int numLabelsX, numLabelsY, headerPosX, headerPosY, footerPosX = 0, footerPosY = 0, numLabels;

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
                MessageBox.Show($"入力値が不正です。{Environment.NewLine}{ex.Message}");
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
            _barcodeProPageSettings.FooterPos = new Point(footerPosX, footerPosY);
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
