using LabelGenerator;

namespace ProductDatabase {
    public partial class ProductPrintSetting : Form {

        private CLabelProLabelSettings LabelProLabelSettings;
        private CLabelProPageSettings LabelProPageSettings;

        public ProductPrintSetting() {
            InitializeComponent();

            LabelProLabelSettings = new CLabelProLabelSettings();
            LabelProPageSettings = new CLabelProPageSettings();
        }

        private void PageSettingsLabelLoad(object sender, EventArgs e) {
            LabelProPageSettings = ((ProductRegistration2Window)Owner!).SettingsLabelPro!.LabelProPageSettings;
            LabelWidthTextBox.Text = LabelProPageSettings.SizeX.ToString();
            LabelHeightTextBox.Text = LabelProPageSettings.SizeY.ToString();
            QuantityXTextBox.Text = LabelProPageSettings.NumLabelsX.ToString();
            QuantityYTextBox.Text = LabelProPageSettings.NumLabelsY.ToString();
            PageOffsetXTextBox.Text = LabelProPageSettings.OffsetX.ToString();
            PageOffsetYTextBox.Text = LabelProPageSettings.OffsetY.ToString();
            LabelIntervalXTextBox.Text = LabelProPageSettings.IntervalX.ToString();
            LabelIntervalYTextBox.Text = LabelProPageSettings.IntervalY.ToString();
            HeaderStringTextBox.Text = LabelProPageSettings.HeaderString;
            HeaderPostionXTextBox.Text = LabelProPageSettings.HeaderPos.X.ToString();
            HeaderPostionYTextBox.Text = LabelProPageSettings.HeaderPos.Y.ToString();

            if (LabelProPageSettings.HeaderFooterFont != null) {
                HeaderFontDialog.Font = LabelProPageSettings.HeaderFooterFont;
                HeaderFooterFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
            }

            LabelProLabelSettings = ((ProductRegistration2Window)Owner).SettingsLabelPro!.LabelProLabelSettings;
            PrintTextHeightTextBox.Text = LabelProLabelSettings.BarcodeHeight.ToString();
            PrintTextMagnitudeTextBox.Text = LabelProLabelSettings.BarcodeMagnitude.ToString();
            PrintTextQuantityTextBox.Text = LabelProLabelSettings.NumLabels.ToString();
            PrintTextFormatTextBox.Text = LabelProLabelSettings.Format;

            if (LabelProLabelSettings.Font != null) {
                TextFontDialog.Font = LabelProLabelSettings.Font;
                PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
            }

            PrintTextPostionXTextBox.Text = LabelProLabelSettings.StringPosX.ToString();
            PrintTextPostionYTextBox.Text = LabelProLabelSettings.StringPosY.ToString();
            BarcodeCenterCheckBox.Checked = LabelProLabelSettings.AlignStringCenter;

            if (BarcodeCenterCheckBox.Checked) {
                PrintTextPostionXTextBox.Enabled = false;
            }
            else {
                PrintTextPostionXTextBox.Enabled = true;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            double SizeX, SizeY, OffsetX, OffsetY, IntervalX, IntervalY, BarcodeHeight, BarcodeMagnitude, StringPosX, StringPosY;
            int NumLabelsX, NumLabelsY, HeaderPosX, HeaderPosY, FooterPosX = 0, FooterPosY = 0, NumLabels;

            try {
                SizeX = double.Parse(LabelWidthTextBox.Text);
                SizeY = double.Parse(LabelHeightTextBox.Text);
                NumLabelsX = int.Parse(QuantityXTextBox.Text);
                NumLabelsY = int.Parse(QuantityYTextBox.Text);
                OffsetX = double.Parse(PageOffsetXTextBox.Text);
                OffsetY = double.Parse(PageOffsetYTextBox.Text);
                IntervalX = double.Parse(LabelIntervalXTextBox.Text);
                IntervalY = double.Parse(LabelIntervalYTextBox.Text);
                HeaderPosX = int.Parse(HeaderPostionXTextBox.Text);
                HeaderPosY = int.Parse(HeaderPostionYTextBox.Text);

                BarcodeHeight = double.Parse(PrintTextHeightTextBox.Text);
                StringPosX = double.Parse(PrintTextPostionXTextBox.Text);
                StringPosY = double.Parse(PrintTextPostionYTextBox.Text);
                BarcodeMagnitude = double.Parse(PrintTextMagnitudeTextBox.Text);
                NumLabels = int.Parse(PrintTextQuantityTextBox.Text);
            } catch (Exception ex) {
                MessageBox.Show($"入力値が不正です。{Environment.NewLine}{ex.Message}");
                DialogResult = DialogResult.None;
                return;
            }

            LabelProPageSettings.SizeX = SizeX;
            LabelProPageSettings.SizeY = SizeY;
            LabelProPageSettings.NumLabelsX = NumLabelsX;
            LabelProPageSettings.NumLabelsY = NumLabelsY;
            LabelProPageSettings.OffsetX = OffsetX;
            LabelProPageSettings.OffsetY = OffsetY;
            LabelProPageSettings.IntervalX = IntervalX;
            LabelProPageSettings.IntervalY = IntervalY;
            LabelProPageSettings.HeaderPos = new Point(HeaderPosX, HeaderPosY);
            LabelProPageSettings.FooterPos = new Point(FooterPosX, FooterPosY);
            LabelProPageSettings.HeaderString = HeaderStringTextBox.Text;
            LabelProPageSettings.HeaderFooterFont = HeaderFontDialog.Font;

            LabelProLabelSettings.BarcodeHeight = BarcodeHeight;
            LabelProLabelSettings.BarcodePosX = 0; // BarcodePosX is not set in the original VB.NET code
            LabelProLabelSettings.BarcodePosY = 0; // BarcodePosY is not set in the original VB.NET code
            LabelProLabelSettings.BarcodeMagnitude = BarcodeMagnitude;
            LabelProLabelSettings.Format = PrintTextFormatTextBox.Text;
            LabelProLabelSettings.Font = TextFontDialog.Font;
            LabelProLabelSettings.StringPosX = StringPosX;
            LabelProLabelSettings.StringPosY = StringPosY;
            LabelProLabelSettings.AlignStringCenter = BarcodeCenterCheckBox.Checked;
            LabelProLabelSettings.NumLabels = int.Parse(PrintTextQuantityTextBox.Text);

            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnHeaderFooterFont_Click(object sender, EventArgs e) {
            DialogResult r = HeaderFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) return;
            HeaderFooterFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
        }

        private void BtnBarcodeFont_Click(object sender, EventArgs e) {
            DialogResult r = TextFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) return;
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
