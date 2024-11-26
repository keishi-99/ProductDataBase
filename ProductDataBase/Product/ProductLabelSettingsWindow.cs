using ProductDatabase.Product;

namespace ProductDatabase {
    public partial class ProductLabelSettingsWindow : Form {

        private readonly CLabelProLabelSettings _labelProLabelSettings;
        private readonly CLabelProPageSettings _labelProPageSettings;

        public ProductLabelSettingsWindow() {
            InitializeComponent();

            _labelProLabelSettings = new CLabelProLabelSettings();
            _labelProPageSettings = new CLabelProPageSettings();
        }

        private void PageSettingsLabelLoad(object sender, EventArgs e) {
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
            var pageSettings = window.SettingsLabelPro.LabelProPageSettings;
            var labelSettings = window.SettingsLabelPro.LabelProLabelSettings;

            SetPageSettings(pageSettings);
            SetLabelSettings(labelSettings);
        }
        private void LoadSettingsFromRePaintWindow(RePrintWindow window) {
            var pageSettings = window.SettingsLabelPro.LabelProPageSettings;
            var labelSettings = window.SettingsLabelPro.LabelProLabelSettings;

            SetPageSettings(pageSettings);
            SetLabelSettings(labelSettings);
        }
        private void SetLabelSettings(CLabelProLabelSettings labelSettings) {
            PrintTextHeightTextBox.Text = labelSettings.BarcodeHeight.ToString();
            PrintTextMagnitudeTextBox.Text = labelSettings.BarcodeMagnitude.ToString();
            PrintTextQuantityTextBox.Text = labelSettings.NumLabels.ToString();
            PrintTextFormatTextBox.Text = labelSettings.Format;

            if (labelSettings.Font != null) {
                TextFontDialog.Font = labelSettings.Font;
                PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
            }

            PrintTextPostionXTextBox.Text = labelSettings.StringPosX.ToString();
            PrintTextPostionYTextBox.Text = labelSettings.StringPosY.ToString();
            BarcodeCenterCheckBox.Checked = labelSettings.AlignStringCenter;

            PrintTextPostionXTextBox.Enabled = !BarcodeCenterCheckBox.Checked;
        }
        private void SetPageSettings(CLabelProPageSettings pageSettings) {
            LabelWidthTextBox.Text = pageSettings.SizeX.ToString();
            LabelHeightTextBox.Text = pageSettings.SizeY.ToString();
            QuantityXTextBox.Text = pageSettings.NumLabelsX.ToString();
            QuantityYTextBox.Text = pageSettings.NumLabelsY.ToString();
            PageOffsetXTextBox.Text = pageSettings.OffsetX.ToString();
            PageOffsetYTextBox.Text = pageSettings.OffsetY.ToString();
            LabelIntervalXTextBox.Text = pageSettings.IntervalX.ToString();
            LabelIntervalYTextBox.Text = pageSettings.IntervalY.ToString();
            HeaderStringTextBox.Text = pageSettings.HeaderString;
            HeaderPostionXTextBox.Text = pageSettings.HeaderPos.X.ToString();
            HeaderPostionYTextBox.Text = pageSettings.HeaderPos.Y.ToString();

            if (pageSettings.HeaderFooterFont != null) {
                HeaderFontDialog.Font = pageSettings.HeaderFooterFont;
                HeaderFooterFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
            }
        }
        private void BtnOK_Click(object sender, EventArgs e) {
            decimal sizeX, sizeY, offsetX, offsetY, intervalX, intervalY, barcodeHeight, barcodeMagnitude, stringPosX, stringPosY;
            int numLabelsX, numLabelsY, headerPosX, headerPosY, footerPosX = 0, footerPosY = 0, numLabels;

            try {
                sizeX = decimal.Parse(LabelWidthTextBox.Text);
                sizeY = decimal.Parse(LabelHeightTextBox.Text);
                numLabelsX = int.Parse(QuantityXTextBox.Text);
                numLabelsY = int.Parse(QuantityYTextBox.Text);
                offsetX = decimal.Parse(PageOffsetXTextBox.Text);
                offsetY = decimal.Parse(PageOffsetYTextBox.Text);
                intervalX = decimal.Parse(LabelIntervalXTextBox.Text);
                intervalY = decimal.Parse(LabelIntervalYTextBox.Text);
                headerPosX = int.Parse(HeaderPostionXTextBox.Text);
                headerPosY = int.Parse(HeaderPostionYTextBox.Text);

                barcodeHeight = decimal.Parse(PrintTextHeightTextBox.Text);
                stringPosX = decimal.Parse(PrintTextPostionXTextBox.Text);
                stringPosY = decimal.Parse(PrintTextPostionYTextBox.Text);
                barcodeMagnitude = decimal.Parse(PrintTextMagnitudeTextBox.Text);
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
