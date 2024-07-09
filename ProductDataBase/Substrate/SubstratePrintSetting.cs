using ProductDatabase.Substrate;

namespace ProductDatabase {
    public partial class SubstratePrintSetting : Form {

        private CLabelSubLabelSettings labelSubLabelSettings;
        private CLabelSubPageSettings labelSubPageSettings;

        public SubstratePrintSetting() {
            InitializeComponent();

            labelSubLabelSettings = new CLabelSubLabelSettings();
            labelSubPageSettings = new CLabelSubPageSettings();
        }

        private void PageSettingsLabelLoad(object sender, EventArgs e) {
            labelSubPageSettings = ((SubstrateRegistrationWindow)Owner!).SettingsLabelSub!.LabelSubPageSettings;
            LabelWidthTextBox.Text = labelSubPageSettings.SizeX.ToString();
            LabelHeightTextBox.Text = labelSubPageSettings.SizeY.ToString();
            QuantityXTextBox.Text = labelSubPageSettings.NumLabelsX.ToString();
            QuantityYTextBox.Text = labelSubPageSettings.NumLabelsY.ToString();
            PageOffsetXTextBox.Text = labelSubPageSettings.OffsetX.ToString();
            PageOffsetYTextBox.Text = labelSubPageSettings.OffsetY.ToString();
            LabelIntervalXTextBox.Text = labelSubPageSettings.IntervalX.ToString();
            LabelIntervalYTextBox.Text = labelSubPageSettings.IntervalY.ToString();
            HeaderStringTextBox.Text = labelSubPageSettings.HeaderString;
            HeaderPostionXTextBox.Text = labelSubPageSettings.HeaderPos.X.ToString();
            HeaderPostionYTextBox.Text = labelSubPageSettings.HeaderPos.Y.ToString();

            if (labelSubPageSettings.HeaderFooterFont != null) {
                HeaderFontDialog.Font = labelSubPageSettings.HeaderFooterFont;
                HeaderFooterFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
            }

            labelSubLabelSettings = ((SubstrateRegistrationWindow)Owner).SettingsLabelSub!.LabelSubLabelSettings;
            PrintTextHeightTextBox.Text = labelSubLabelSettings.BarcodeHeight.ToString();
            PrintTextMagnitudeTextBox.Text = labelSubLabelSettings.BarcodeMagnitude.ToString();
            PrintTextQuantityTextBox.Text = labelSubLabelSettings.NumLabels.ToString();
            PrintTextFormatTextBox.Text = labelSubLabelSettings.Format;

            if (labelSubLabelSettings.Font != null) {
                TextFontDialog.Font = labelSubLabelSettings.Font;
                PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
            }

            PrintTextPostionXTextBox.Text = labelSubLabelSettings.StringPosX.ToString();
            PrintTextPostionYTextBox.Text = labelSubLabelSettings.StringPosY.ToString();
            PrintTextCenterCheckBox.Checked = labelSubLabelSettings.AlignStringCenter;

            if (PrintTextCenterCheckBox.Checked) {
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

            labelSubPageSettings.SizeX = _sizeX;
            labelSubPageSettings.SizeY = _sizeY;
            labelSubPageSettings.NumLabelsX = _numLabelsX;
            labelSubPageSettings.NumLabelsY = _numLabelsY;
            labelSubPageSettings.OffsetX = _offsetX;
            labelSubPageSettings.OffsetY = _offsetY;
            labelSubPageSettings.IntervalX = _intervalX;
            labelSubPageSettings.IntervalY = _intervalY;
            labelSubPageSettings.HeaderPos = new Point(_headerPosX, _headerPosY);
            labelSubPageSettings.FooterPos = new Point(_footerPosX, _footerPosY);
            labelSubPageSettings.HeaderString = HeaderStringTextBox.Text;
            labelSubPageSettings.HeaderFooterFont = HeaderFontDialog.Font;

            labelSubLabelSettings.BarcodeHeight = _barcodeHeight;
            labelSubLabelSettings.BarcodePosX = 0;
            labelSubLabelSettings.BarcodePosY = 0;
            labelSubLabelSettings.BarcodeMagnitude = _barcodeMagnitude;
            labelSubLabelSettings.Format = PrintTextFormatTextBox.Text;
            labelSubLabelSettings.Font = TextFontDialog.Font;
            labelSubLabelSettings.StringPosX = _stringPosX;
            labelSubLabelSettings.StringPosY = _stringPosY;
            labelSubLabelSettings.AlignStringCenter = PrintTextCenterCheckBox.Checked;
            labelSubLabelSettings.NumLabels = int.Parse(PrintTextQuantityTextBox.Text);

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
            if (PrintTextCenterCheckBox.Checked) {
                PrintTextPostionXTextBox.Enabled = false;
            }
            else {
                PrintTextPostionXTextBox.Enabled = true;
            }
        }

        private void SubstratePrintSetting_Load(object sender, EventArgs e) { PageSettingsLabelLoad(sender, e); }

        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
