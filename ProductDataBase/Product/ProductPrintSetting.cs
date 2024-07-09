using ProductDatabase.Product;

namespace ProductDatabase {
    public partial class ProductPrintSetting : Form {

        private CLabelProLabelSettings labelProLabelSettings;
        private CLabelProPageSettings labelProPageSettings;

        public ProductPrintSetting() {
            InitializeComponent();

            labelProLabelSettings = new CLabelProLabelSettings();
            labelProPageSettings = new CLabelProPageSettings();
        }

        private void PageSettingsLabelLoad(object sender, EventArgs e) {
            labelProPageSettings = ((ProductRegistration2Window)Owner!).SettingsLabelPro!._labelProPageSettings;
            LabelWidthTextBox.Text = labelProPageSettings.SizeX.ToString();
            LabelHeightTextBox.Text = labelProPageSettings.SizeY.ToString();
            QuantityXTextBox.Text = labelProPageSettings.NumLabelsX.ToString();
            QuantityYTextBox.Text = labelProPageSettings.NumLabelsY.ToString();
            PageOffsetXTextBox.Text = labelProPageSettings.OffsetX.ToString();
            PageOffsetYTextBox.Text = labelProPageSettings.OffsetY.ToString();
            LabelIntervalXTextBox.Text = labelProPageSettings.IntervalX.ToString();
            LabelIntervalYTextBox.Text = labelProPageSettings.IntervalY.ToString();
            HeaderStringTextBox.Text = labelProPageSettings.HeaderString;
            HeaderPostionXTextBox.Text = labelProPageSettings.HeaderPos.X.ToString();
            HeaderPostionYTextBox.Text = labelProPageSettings.HeaderPos.Y.ToString();

            if (labelProPageSettings.HeaderFooterFont != null) {
                HeaderFontDialog.Font = labelProPageSettings.HeaderFooterFont;
                HeaderFooterFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
            }

            labelProLabelSettings = ((ProductRegistration2Window)Owner).SettingsLabelPro!._labelProLabelSettings;
            PrintTextHeightTextBox.Text = labelProLabelSettings.BarcodeHeight.ToString();
            PrintTextMagnitudeTextBox.Text = labelProLabelSettings.BarcodeMagnitude.ToString();
            PrintTextQuantityTextBox.Text = labelProLabelSettings.NumLabels.ToString();
            PrintTextFormatTextBox.Text = labelProLabelSettings.Format;

            if (labelProLabelSettings.Font != null) {
                TextFontDialog.Font = labelProLabelSettings.Font;
                PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
            }

            PrintTextPostionXTextBox.Text = labelProLabelSettings.StringPosX.ToString();
            PrintTextPostionYTextBox.Text = labelProLabelSettings.StringPosY.ToString();
            BarcodeCenterCheckBox.Checked = labelProLabelSettings.AlignStringCenter;

            if (BarcodeCenterCheckBox.Checked) {
                PrintTextPostionXTextBox.Enabled = false;
            }
            else {
                PrintTextPostionXTextBox.Enabled = true;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            double _sizeX, _sizeY, _offsetX, _offsetY, _intervalX, _intervalY, _barcodeHeight, _barcodeMagnitude, _stringPosX, _stringPosY;
            int _numLabelsX, _numLabelsY, _headerPosX, _headerPosY, _footerPosX = 0, _footerPosY = 0, _numLabels;

            try {
                _sizeX = double.Parse(LabelWidthTextBox.Text);
                _sizeY = double.Parse(LabelHeightTextBox.Text);
                _numLabelsX = int.Parse(QuantityXTextBox.Text);
                _numLabelsY = int.Parse(QuantityYTextBox.Text);
                _offsetX = double.Parse(PageOffsetXTextBox.Text);
                _offsetY = double.Parse(PageOffsetYTextBox.Text);
                _intervalX = double.Parse(LabelIntervalXTextBox.Text);
                _intervalY = double.Parse(LabelIntervalYTextBox.Text);
                _headerPosX = int.Parse(HeaderPostionXTextBox.Text);
                _headerPosY = int.Parse(HeaderPostionYTextBox.Text);

                _barcodeHeight = double.Parse(PrintTextHeightTextBox.Text);
                _stringPosX = double.Parse(PrintTextPostionXTextBox.Text);
                _stringPosY = double.Parse(PrintTextPostionYTextBox.Text);
                _barcodeMagnitude = double.Parse(PrintTextMagnitudeTextBox.Text);
                _numLabels = int.Parse(PrintTextQuantityTextBox.Text);
            } catch (Exception ex) {
                MessageBox.Show($"入力値が不正です。{Environment.NewLine}{ex.Message}");
                DialogResult = DialogResult.None;
                return;
            }

            labelProPageSettings.SizeX = _sizeX;
            labelProPageSettings.SizeY = _sizeY;
            labelProPageSettings.NumLabelsX = _numLabelsX;
            labelProPageSettings.NumLabelsY = _numLabelsY;
            labelProPageSettings.OffsetX = _offsetX;
            labelProPageSettings.OffsetY = _offsetY;
            labelProPageSettings.IntervalX = _intervalX;
            labelProPageSettings.IntervalY = _intervalY;
            labelProPageSettings.HeaderPos = new Point(_headerPosX, _headerPosY);
            labelProPageSettings.FooterPos = new Point(_footerPosX, _footerPosY);
            labelProPageSettings.HeaderString = HeaderStringTextBox.Text;
            labelProPageSettings.HeaderFooterFont = HeaderFontDialog.Font;

            labelProLabelSettings.BarcodeHeight = _barcodeHeight;
            labelProLabelSettings.BarcodePosX = 0;
            labelProLabelSettings.BarcodePosY = 0;
            labelProLabelSettings.BarcodeMagnitude = _barcodeMagnitude;
            labelProLabelSettings.Format = PrintTextFormatTextBox.Text;
            labelProLabelSettings.Font = TextFontDialog.Font;
            labelProLabelSettings.StringPosX = _stringPosX;
            labelProLabelSettings.StringPosY = _stringPosY;
            labelProLabelSettings.AlignStringCenter = BarcodeCenterCheckBox.Checked;
            labelProLabelSettings.NumLabels = int.Parse(PrintTextQuantityTextBox.Text);

            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnHeaderFooterFont_Click(object sender, EventArgs e) {
            DialogResult _r = HeaderFontDialog.ShowDialog();
            if (_r == DialogResult.Cancel) return;
            HeaderFooterFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
        }

        private void BtnBarcodeFont_Click(object sender, EventArgs e) {
            DialogResult _r = TextFontDialog.ShowDialog();
            if (_r == DialogResult.Cancel) return;
            PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
        }

        private void BarcodeCenterCheckBox_CheckedChanged(object sender, EventArgs e) {
            if (BarcodeCenterCheckBox.Checked) {
                PrintTextPostionXTextBox.Enabled = false;
            }
            else {
                PrintTextPostionXTextBox.Enabled = true;
            }
        }

        private void ProductPrintSetting_Load(object sender, EventArgs e) { PageSettingsLabelLoad(sender, e); }

        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
