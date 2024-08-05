using ProductDatabase.Product;

namespace ProductDatabase {
    public partial class ProductPrintSetting : Form {

        private CLabelProLabelSettings _labelProLabelSettings;
        private CLabelProPageSettings _labelProPageSettings;

        public ProductPrintSetting() {
            InitializeComponent();

            _labelProLabelSettings = new CLabelProLabelSettings();
            _labelProPageSettings = new CLabelProPageSettings();
        }

        private void PageSettingsLabelLoad(object sender, EventArgs e) {
            _labelProPageSettings = ((ProductRegistration2Window)Owner!).SettingsLabelPro!.LabelProPageSettings;
            LabelWidthTextBox.Text = _labelProPageSettings.SizeX.ToString();
            LabelHeightTextBox.Text = _labelProPageSettings.SizeY.ToString();
            QuantityXTextBox.Text = _labelProPageSettings.NumLabelsX.ToString();
            QuantityYTextBox.Text = _labelProPageSettings.NumLabelsY.ToString();
            PageOffsetXTextBox.Text = _labelProPageSettings.OffsetX.ToString();
            PageOffsetYTextBox.Text = _labelProPageSettings.OffsetY.ToString();
            LabelIntervalXTextBox.Text = _labelProPageSettings.IntervalX.ToString();
            LabelIntervalYTextBox.Text = _labelProPageSettings.IntervalY.ToString();
            HeaderStringTextBox.Text = _labelProPageSettings.HeaderString;
            HeaderPostionXTextBox.Text = _labelProPageSettings.HeaderPos.X.ToString();
            HeaderPostionYTextBox.Text = _labelProPageSettings.HeaderPos.Y.ToString();

            if (_labelProPageSettings.HeaderFooterFont != null) {
                HeaderFontDialog.Font = _labelProPageSettings.HeaderFooterFont;
                HeaderFooterFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
            }

            _labelProLabelSettings = ((ProductRegistration2Window)Owner).SettingsLabelPro!.LabelProLabelSettings;
            PrintTextHeightTextBox.Text = _labelProLabelSettings.BarcodeHeight.ToString();
            PrintTextMagnitudeTextBox.Text = _labelProLabelSettings.BarcodeMagnitude.ToString();
            PrintTextQuantityTextBox.Text = _labelProLabelSettings.NumLabels.ToString();
            PrintTextFormatTextBox.Text = _labelProLabelSettings.Format;

            if (_labelProLabelSettings.Font != null) {
                TextFontDialog.Font = _labelProLabelSettings.Font;
                PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
            }

            PrintTextPostionXTextBox.Text = _labelProLabelSettings.StringPosX.ToString();
            PrintTextPostionYTextBox.Text = _labelProLabelSettings.StringPosY.ToString();
            BarcodeCenterCheckBox.Checked = _labelProLabelSettings.AlignStringCenter;

            PrintTextPostionXTextBox.Enabled = !BarcodeCenterCheckBox.Checked;
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            double sizeX, sizeY, offsetX, offsetY, intervalX, intervalY, barcodeHeight, barcodeMagnitude, stringPosX, stringPosY;
            int numLabelsX, numLabelsY, headerPosX, headerPosY, footerPosX = 0, footerPosY = 0, numLabels;

            try {
                sizeX = double.Parse(LabelWidthTextBox.Text);
                sizeY = double.Parse(LabelHeightTextBox.Text);
                numLabelsX = int.Parse(QuantityXTextBox.Text);
                numLabelsY = int.Parse(QuantityYTextBox.Text);
                offsetX = double.Parse(PageOffsetXTextBox.Text);
                offsetY = double.Parse(PageOffsetYTextBox.Text);
                intervalX = double.Parse(LabelIntervalXTextBox.Text);
                intervalY = double.Parse(LabelIntervalYTextBox.Text);
                headerPosX = int.Parse(HeaderPostionXTextBox.Text);
                headerPosY = int.Parse(HeaderPostionYTextBox.Text);

                barcodeHeight = double.Parse(PrintTextHeightTextBox.Text);
                stringPosX = double.Parse(PrintTextPostionXTextBox.Text);
                stringPosY = double.Parse(PrintTextPostionYTextBox.Text);
                barcodeMagnitude = double.Parse(PrintTextMagnitudeTextBox.Text);
                numLabels = int.Parse(PrintTextQuantityTextBox.Text);
            } catch (Exception ex) {
                MessageBox.Show($"入力値が不正です。{Environment.NewLine}{ex.Message}");
                DialogResult = DialogResult.None;
                return;
            }

            _labelProPageSettings.SizeX = sizeX;
            _labelProPageSettings.SizeY = sizeY;
            _labelProPageSettings.NumLabelsX = numLabelsX;
            _labelProPageSettings.NumLabelsY = numLabelsY;
            _labelProPageSettings.OffsetX = offsetX;
            _labelProPageSettings.OffsetY = offsetY;
            _labelProPageSettings.IntervalX = intervalX;
            _labelProPageSettings.IntervalY = intervalY;
            _labelProPageSettings.HeaderPos = new Point(headerPosX, headerPosY);
            _labelProPageSettings.FooterPos = new Point(footerPosX, footerPosY);
            _labelProPageSettings.HeaderString = HeaderStringTextBox.Text;
            _labelProPageSettings.HeaderFooterFont = HeaderFontDialog.Font;

            _labelProLabelSettings.BarcodeHeight = barcodeHeight;
            _labelProLabelSettings.BarcodePosX = 0;
            _labelProLabelSettings.BarcodePosY = 0;
            _labelProLabelSettings.BarcodeMagnitude = barcodeMagnitude;
            _labelProLabelSettings.Format = PrintTextFormatTextBox.Text;
            _labelProLabelSettings.Font = TextFontDialog.Font;
            _labelProLabelSettings.StringPosX = stringPosX;
            _labelProLabelSettings.StringPosY = stringPosY;
            _labelProLabelSettings.AlignStringCenter = BarcodeCenterCheckBox.Checked;
            _labelProLabelSettings.NumLabels = int.Parse(PrintTextQuantityTextBox.Text);

            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnHeaderFooterFont_Click(object sender, EventArgs e) {
            var r = HeaderFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }

            HeaderFooterFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
        }

        private void BtnBarcodeFont_Click(object sender, EventArgs e) {
            var r = TextFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }

            PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
        }

        private void BarcodeCenterCheckBox_CheckedChanged(object sender, EventArgs e) {
            PrintTextPostionXTextBox.Enabled = !BarcodeCenterCheckBox.Checked;
        }

        private void ProductPrintSetting_Load(object sender, EventArgs e) { PageSettingsLabelLoad(sender, e); }

        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
