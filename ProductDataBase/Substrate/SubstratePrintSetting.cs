using ProductDatabase.Substrate;

namespace ProductDatabase {
    public partial class SubstratePrintSetting : Form {

        private CLabelSubLabelSettings _labelSubLabelSettings;
        private CLabelSubPageSettings _labelSubPageSettings;

        public SubstratePrintSetting() {
            InitializeComponent();

            _labelSubLabelSettings = new CLabelSubLabelSettings();
            _labelSubPageSettings = new CLabelSubPageSettings();
        }

        private void PageSettingsLabelLoad(object sender, EventArgs e) {
            _labelSubPageSettings = ((SubstrateRegistrationWindow)Owner!).SettingsLabelSub!.LabelSubPageSettings;
            LabelWidthTextBox.Text = _labelSubPageSettings.SizeX.ToString();
            LabelHeightTextBox.Text = _labelSubPageSettings.SizeY.ToString();
            QuantityXTextBox.Text = _labelSubPageSettings.NumLabelsX.ToString();
            QuantityYTextBox.Text = _labelSubPageSettings.NumLabelsY.ToString();
            PageOffsetXTextBox.Text = _labelSubPageSettings.OffsetX.ToString();
            PageOffsetYTextBox.Text = _labelSubPageSettings.OffsetY.ToString();
            LabelIntervalXTextBox.Text = _labelSubPageSettings.IntervalX.ToString();
            LabelIntervalYTextBox.Text = _labelSubPageSettings.IntervalY.ToString();
            HeaderStringTextBox.Text = _labelSubPageSettings.HeaderString;
            HeaderPostionXTextBox.Text = _labelSubPageSettings.HeaderPos.X.ToString();
            HeaderPostionYTextBox.Text = _labelSubPageSettings.HeaderPos.Y.ToString();

            if (_labelSubPageSettings.HeaderFooterFont != null) {
                HeaderFontDialog.Font = _labelSubPageSettings.HeaderFooterFont;
                HeaderFooterFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
            }

            _labelSubLabelSettings = ((SubstrateRegistrationWindow)Owner).SettingsLabelSub!.LabelSubLabelSettings;
            PrintTextHeightTextBox.Text = _labelSubLabelSettings.BarcodeHeight.ToString();
            PrintTextMagnitudeTextBox.Text = _labelSubLabelSettings.BarcodeMagnitude.ToString();
            PrintTextQuantityTextBox.Text = _labelSubLabelSettings.NumLabels.ToString();
            PrintTextFormatTextBox.Text = _labelSubLabelSettings.Format;

            if (_labelSubLabelSettings.Font != null) {
                TextFontDialog.Font = _labelSubLabelSettings.Font;
                PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
            }

            PrintTextPostionXTextBox.Text = _labelSubLabelSettings.StringPosX.ToString();
            PrintTextPostionYTextBox.Text = _labelSubLabelSettings.StringPosY.ToString();
            PrintTextCenterCheckBox.Checked = _labelSubLabelSettings.AlignStringCenter;

            PrintTextPostionXTextBox.Enabled = !PrintTextCenterCheckBox.Checked;
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            double sizeX, sizeY, offsetX, offsetY, intervalX, intervalY, barcodeHeight, barcodeMagnitude, stringPosX, stringPosY;
            int numLabelsX, numLabelsY, headerPosX, headerPosY, footerPosX = 0, footerPosY = 0, numLabels;

            try {
                sizeX = Double.Parse(LabelWidthTextBox.Text);
                sizeY = Double.Parse(LabelHeightTextBox.Text);
                numLabelsX = Int32.Parse(QuantityXTextBox.Text);
                numLabelsY = Int32.Parse(QuantityYTextBox.Text);
                offsetX = Double.Parse(PageOffsetXTextBox.Text);
                offsetY = Double.Parse(PageOffsetYTextBox.Text);
                intervalX = Double.Parse(LabelIntervalXTextBox.Text);
                intervalY = Double.Parse(LabelIntervalYTextBox.Text);
                headerPosX = Int32.Parse(HeaderPostionXTextBox.Text);
                headerPosY = Int32.Parse(HeaderPostionYTextBox.Text);

                barcodeHeight = Double.Parse(PrintTextHeightTextBox.Text);
                stringPosX = Double.Parse(PrintTextPostionXTextBox.Text);
                stringPosY = Double.Parse(PrintTextPostionYTextBox.Text);
                barcodeMagnitude = Double.Parse(PrintTextMagnitudeTextBox.Text);
                numLabels = Int32.Parse(PrintTextQuantityTextBox.Text);
            } catch (Exception ex) {
                MessageBox.Show($"入力値が不正です。{Environment.NewLine}{ex.Message}");
                DialogResult = DialogResult.None;
                return;
            }

            _labelSubPageSettings.SizeX = sizeX;
            _labelSubPageSettings.SizeY = sizeY;
            _labelSubPageSettings.NumLabelsX = numLabelsX;
            _labelSubPageSettings.NumLabelsY = numLabelsY;
            _labelSubPageSettings.OffsetX = offsetX;
            _labelSubPageSettings.OffsetY = offsetY;
            _labelSubPageSettings.IntervalX = intervalX;
            _labelSubPageSettings.IntervalY = intervalY;
            _labelSubPageSettings.HeaderPos = new Point(headerPosX, headerPosY);
            _labelSubPageSettings.FooterPos = new Point(footerPosX, footerPosY);
            _labelSubPageSettings.HeaderString = HeaderStringTextBox.Text;
            _labelSubPageSettings.HeaderFooterFont = HeaderFontDialog.Font;

            _labelSubLabelSettings.BarcodeHeight = barcodeHeight;
            _labelSubLabelSettings.BarcodePosX = 0;
            _labelSubLabelSettings.BarcodePosY = 0;
            _labelSubLabelSettings.BarcodeMagnitude = barcodeMagnitude;
            _labelSubLabelSettings.Format = PrintTextFormatTextBox.Text;
            _labelSubLabelSettings.Font = TextFontDialog.Font;
            _labelSubLabelSettings.StringPosX = stringPosX;
            _labelSubLabelSettings.StringPosY = stringPosY;
            _labelSubLabelSettings.AlignStringCenter = PrintTextCenterCheckBox.Checked;
            _labelSubLabelSettings.NumLabels = Int32.Parse(PrintTextQuantityTextBox.Text);

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
            PrintTextPostionXTextBox.Enabled = !PrintTextCenterCheckBox.Checked;
        }

        private void SubstratePrintSetting_Load(object sender, EventArgs e) { PageSettingsLabelLoad(sender, e); }

        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
