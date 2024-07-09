using ProductDatabase.Product;

namespace ProductDatabase {
    public partial class ProductBarcodePrintSettingsWindow : Form {

        private CBarcodeProLabelSettings barcodeProLabelSettings;
        private CBarcodeProPageSettings barcodeProPageSettings;

        public ProductBarcodePrintSettingsWindow() {
            InitializeComponent();

            barcodeProLabelSettings = new CBarcodeProLabelSettings();
            barcodeProPageSettings = new CBarcodeProPageSettings();
        }

        private void PageSettingsBarcodeLoad(object sender, EventArgs e) {
            barcodeProPageSettings = ((ProductRegistration2Window)Owner!).SettingsBarcodePro!.BarcodeProPageSettings;
            BarcodeLabelWidthTextBox.Text = barcodeProPageSettings.SizeX.ToString();
            BarcodeLabelHeightTextBox.Text = barcodeProPageSettings.SizeY.ToString();
            BarcodeQuantityXTextBox.Text = barcodeProPageSettings.NumLabelsX.ToString();
            BarcodeQuantityYTextBox.Text = barcodeProPageSettings.NumLabelsY.ToString();
            BarcodePageOffsetXTextBox.Text = barcodeProPageSettings.OffsetX.ToString();
            BarcodePageOffsetYTextBox.Text = barcodeProPageSettings.OffsetY.ToString();
            BarcodeLabelIntervalXTextBox.Text = barcodeProPageSettings.IntervalX.ToString();
            BarcodeLabelIntervalYTextBox.Text = barcodeProPageSettings.IntervalY.ToString();
            BarcodeHeaderStringTextBox.Text = barcodeProPageSettings.HeaderString;
            BarcodeHeaderPostionXTextBox.Text = barcodeProPageSettings.HeaderPos.X.ToString();
            BarcodeHeaderPostionYTextBox.Text = barcodeProPageSettings.HeaderPos.Y.ToString();

            if (barcodeProPageSettings.HeaderFooterFont != null) {
                BarcodeHeaderFontDialog.Font = barcodeProPageSettings.HeaderFooterFont;
                BarcodeHeaderFooterFontTextBox.Text = $"{BarcodeHeaderFontDialog.Font.Name} {BarcodeHeaderFontDialog.Font.SizeInPoints}pt";
            }

            barcodeProLabelSettings = ((ProductRegistration2Window)Owner).SettingsBarcodePro!.BarcodeProLabelSettings;
            BarcodeHeightTextBox.Text = barcodeProLabelSettings.BarcodeHeight.ToString();
            BarcodeMagnitudeTextBox.Text = barcodeProLabelSettings.BarcodeMagnitude.ToString();
            BarcodeQuantityTextBox.Text = barcodeProLabelSettings.NumLabels.ToString();
            BarcodeFormatTextBox.Text = barcodeProLabelSettings.Format;

            if (barcodeProLabelSettings.Font != null) {
                BarcodeFontDialog.Font = barcodeProLabelSettings.Font;
                BarcodeFontTextBox.Text = $"{BarcodeFontDialog.Font.Name} {BarcodeFontDialog.Font.SizeInPoints}pt";
            }

            BarcodePostionXTextBox.Text = barcodeProLabelSettings.BarcodePosX.ToString();
            BarcodePostionYTextBox.Text = barcodeProLabelSettings.BarcodePosY.ToString();
            BarcodeCenterCheckBox.Checked = barcodeProLabelSettings.AlignBarcodeCenter;

            if (BarcodeCenterCheckBox.Checked) { BarcodePostionXTextBox.Enabled = false; }
            else { BarcodePostionXTextBox.Enabled = true; }

            FontPostionXTextBox.Text = barcodeProLabelSettings.StringPosX.ToString();
            FontPostionYTextBox.Text = barcodeProLabelSettings.StringPosY.ToString();
            FontCenterCheckBox.Checked = barcodeProLabelSettings.AlignStringCenter;
            if (FontCenterCheckBox.Checked) { FontPostionXTextBox.Enabled = false; }
            else { FontPostionXTextBox.Enabled = true; }
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            double _sizeX, _sizeY, _offsetX, _offsetY, _intervalX, _intervalY, _barcodePosX, _barcodePosY, _barcodeHeight, _barcodeMagnitude, _stringPosX, _stringPosY;
            int _numLabelsX, _numLabelsY, _headerPosX, _headerPosY, _footerPosX = 0, _footerPosY = 0, _numLabels;

            try {
                _sizeX = double.Parse(BarcodeLabelWidthTextBox.Text);
                _sizeY = double.Parse(BarcodeLabelHeightTextBox.Text);
                _numLabelsX = int.Parse(BarcodeQuantityXTextBox.Text);
                _numLabelsY = int.Parse(BarcodeQuantityYTextBox.Text);
                _offsetX = double.Parse(BarcodePageOffsetXTextBox.Text);
                _offsetY = double.Parse(BarcodePageOffsetYTextBox.Text);
                _intervalX = double.Parse(BarcodeLabelIntervalXTextBox.Text);
                _intervalY = double.Parse(BarcodeLabelIntervalYTextBox.Text);
                _headerPosX = int.Parse(BarcodeHeaderPostionXTextBox.Text);
                _headerPosY = int.Parse(BarcodeHeaderPostionYTextBox.Text);

                _barcodeHeight = double.Parse(BarcodeHeightTextBox.Text);
                _barcodePosX = double.Parse(BarcodePostionXTextBox.Text);
                _barcodePosY = double.Parse(BarcodePostionYTextBox.Text);
                _barcodeMagnitude = double.Parse(BarcodeMagnitudeTextBox.Text);
                _stringPosX = double.Parse(FontPostionXTextBox.Text);
                _stringPosY = double.Parse(FontPostionYTextBox.Text);
                _numLabels = int.Parse(BarcodeQuantityTextBox.Text);
            } catch (Exception ex) {
                MessageBox.Show($"入力値が不正です。{Environment.NewLine}{ex.Message}");
                DialogResult = DialogResult.None;
                return;
            }

            barcodeProPageSettings.SizeX = _sizeX;
            barcodeProPageSettings.SizeY = _sizeY;
            barcodeProPageSettings.NumLabelsX = _numLabelsX;
            barcodeProPageSettings.NumLabelsY = _numLabelsY;
            barcodeProPageSettings.OffsetX = _offsetX;
            barcodeProPageSettings.OffsetY = _offsetY;
            barcodeProPageSettings.IntervalX = _intervalX;
            barcodeProPageSettings.IntervalY = _intervalY;
            barcodeProPageSettings.HeaderPos = new Point(_headerPosX, _headerPosY);
            barcodeProPageSettings.FooterPos = new Point(_footerPosX, _footerPosY);
            barcodeProPageSettings.HeaderString = BarcodeHeaderStringTextBox.Text;
            barcodeProPageSettings.HeaderFooterFont = BarcodeHeaderFontDialog.Font;

            barcodeProLabelSettings.BarcodeHeight = _barcodeHeight;
            barcodeProLabelSettings.BarcodePosX = _barcodePosX;
            barcodeProLabelSettings.BarcodePosY = _barcodePosY;
            barcodeProLabelSettings.AlignBarcodeCenter = BarcodeCenterCheckBox.Checked;
            barcodeProLabelSettings.BarcodeMagnitude = _barcodeMagnitude;
            barcodeProLabelSettings.Format = BarcodeFormatTextBox.Text;
            barcodeProLabelSettings.Font = BarcodeFontDialog.Font;
            barcodeProLabelSettings.StringPosX = _stringPosX;
            barcodeProLabelSettings.StringPosY = _stringPosY;
            barcodeProLabelSettings.AlignStringCenter = FontCenterCheckBox.Checked;
            barcodeProLabelSettings.NumLabels = int.Parse(BarcodeQuantityTextBox.Text);

            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnHeaderFooterFont_Click(object sender, EventArgs e) {
            DialogResult _r = BarcodeHeaderFontDialog.ShowDialog();
            if (_r == DialogResult.Cancel) return;
            BarcodeHeaderFooterFontTextBox.Text = $"{BarcodeHeaderFontDialog.Font.Name} {BarcodeHeaderFontDialog.Font.SizeInPoints}pt";
        }

        private void BtnBarcodeFont_Click(object sender, EventArgs e) {
            DialogResult _r = BarcodeFontDialog.ShowDialog();
            if (_r == DialogResult.Cancel) return;
            BarcodeFontTextBox.Text = $"{BarcodeFontDialog.Font.Name} {BarcodeFontDialog.Font.SizeInPoints}pt";
        }

        private void BarcodeCenterCheckBox_CheckedChanged(object sender, EventArgs e) {
            if (BarcodeCenterCheckBox.Checked) {
                BarcodePostionXTextBox.Enabled = false;
            }
            else {
                BarcodePostionXTextBox.Enabled = true;
            }
        }

        private void ProductBarcodePrintSetting_Load(object sender, EventArgs e) { PageSettingsBarcodeLoad(sender, e); }

        private void CloseButton_Click(object sender, EventArgs e) { Close(); }

        private void FontCenterCheckBox_CheckedChanged(object sender, EventArgs e) {
            if (FontCenterCheckBox.Checked) {
                FontPostionXTextBox.Enabled = false;
            }
            else {
                FontPostionXTextBox.Enabled = true;
            }

        }
    }
}
