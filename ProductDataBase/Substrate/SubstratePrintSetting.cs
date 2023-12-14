using ProductDatabase.Substrate;

namespace ProductDatabase
{
    public partial class SubstratePrintSetting : Form {

        private CLabelSubLabelSettings LabelSubLabelSettings;
        private CLabelSubPageSettings LabelSubPageSettings;

        public SubstratePrintSetting() {
            InitializeComponent();

            LabelSubLabelSettings = new CLabelSubLabelSettings();
            LabelSubPageSettings = new CLabelSubPageSettings();
        }

        private void PageSettingsLabelLoad(object sender, EventArgs e) {
            LabelSubPageSettings = ((SubstrateRegistrationWindow)Owner!).SettingsLabelSub!.LabelSubPageSettings;
            LabelWidthTextBox.Text = LabelSubPageSettings.SizeX.ToString();
            LabelHeightTextBox.Text = LabelSubPageSettings.SizeY.ToString();
            QuantityXTextBox.Text = LabelSubPageSettings.NumLabelsX.ToString();
            QuantityYTextBox.Text = LabelSubPageSettings.NumLabelsY.ToString();
            PageOffsetXTextBox.Text = LabelSubPageSettings.OffsetX.ToString();
            PageOffsetYTextBox.Text = LabelSubPageSettings.OffsetY.ToString();
            LabelIntervalXTextBox.Text = LabelSubPageSettings.IntervalX.ToString();
            LabelIntervalYTextBox.Text = LabelSubPageSettings.IntervalY.ToString();
            HeaderStringTextBox.Text = LabelSubPageSettings.HeaderString;
            HeaderPostionXTextBox.Text = LabelSubPageSettings.HeaderPos.X.ToString();
            HeaderPostionYTextBox.Text = LabelSubPageSettings.HeaderPos.Y.ToString();

            if (LabelSubPageSettings.HeaderFooterFont != null) {
                HeaderFontDialog.Font = LabelSubPageSettings.HeaderFooterFont;
                HeaderFooterFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
            }

            LabelSubLabelSettings = ((SubstrateRegistrationWindow)Owner).SettingsLabelSub!.LabelSubLabelSettings;
            PrintTextHeightTextBox.Text = LabelSubLabelSettings.BarcodeHeight.ToString();
            PrintTextMagnitudeTextBox.Text = LabelSubLabelSettings.BarcodeMagnitude.ToString();
            PrintTextQuantityTextBox.Text = LabelSubLabelSettings.NumLabels.ToString();
            PrintTextFormatTextBox.Text = LabelSubLabelSettings.Format;

            if (LabelSubLabelSettings.Font != null) {
                TextFontDialog.Font = LabelSubLabelSettings.Font;
                PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
            }

            PrintTextPostionXTextBox.Text = LabelSubLabelSettings.StringPosX.ToString();
            PrintTextPostionYTextBox.Text = LabelSubLabelSettings.StringPosY.ToString();
            PrintTextCenterCheckBox.Checked = LabelSubLabelSettings.AlignStringCenter;

            if (PrintTextCenterCheckBox.Checked) {
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

            LabelSubPageSettings.SizeX = SizeX;
            LabelSubPageSettings.SizeY = SizeY;
            LabelSubPageSettings.NumLabelsX = NumLabelsX;
            LabelSubPageSettings.NumLabelsY = NumLabelsY;
            LabelSubPageSettings.OffsetX = OffsetX;
            LabelSubPageSettings.OffsetY = OffsetY;
            LabelSubPageSettings.IntervalX = IntervalX;
            LabelSubPageSettings.IntervalY = IntervalY;
            LabelSubPageSettings.HeaderPos = new Point(HeaderPosX, HeaderPosY);
            LabelSubPageSettings.FooterPos = new Point(FooterPosX, FooterPosY);
            LabelSubPageSettings.HeaderString = HeaderStringTextBox.Text;
            LabelSubPageSettings.HeaderFooterFont = HeaderFontDialog.Font;

            LabelSubLabelSettings.BarcodeHeight = BarcodeHeight;
            LabelSubLabelSettings.BarcodePosX = 0; // BarcodePosX is not set in the original VB.NET code
            LabelSubLabelSettings.BarcodePosY = 0; // BarcodePosY is not set in the original VB.NET code
            LabelSubLabelSettings.BarcodeMagnitude = BarcodeMagnitude;
            LabelSubLabelSettings.Format = PrintTextFormatTextBox.Text;
            LabelSubLabelSettings.Font = TextFontDialog.Font;
            LabelSubLabelSettings.StringPosX = StringPosX;
            LabelSubLabelSettings.StringPosY = StringPosY;
            LabelSubLabelSettings.AlignStringCenter = PrintTextCenterCheckBox.Checked;
            LabelSubLabelSettings.NumLabels = int.Parse(PrintTextQuantityTextBox.Text);

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
