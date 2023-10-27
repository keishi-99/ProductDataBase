using LabelGenerator;

namespace ProductDataBase {
    public partial class SubstratePrintSetting : Form {

        private CLabelSubLabelSettings LabelSubLabelSettings;
        private CLabelSubPageSettings LabelSubPageSettings;

        public SubstratePrintSetting() {
            InitializeComponent();

            LabelSubLabelSettings = new CLabelSubLabelSettings();
            LabelSubPageSettings = new CLabelSubPageSettings();
        }

        private void FrmPageSettings_Label_Load(object sender, EventArgs e) {
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
                SubstratePrintFontDialog.Font = LabelSubPageSettings.HeaderFooterFont;
                HeaderFooterFontTextBox.Text = $"{SubstratePrintFontDialog.Font.Name} {SubstratePrintFontDialog.Font.SizeInPoints}pt";
            }

            LabelSubLabelSettings = ((SubstrateRegistrationWindow)Owner).SettingsLabelSub!.LabelSubLabelSettings;
            BarcodeHeightTextBox.Text = LabelSubLabelSettings.BarcodeHeight.ToString();
            BarcodeMagnitudeTextBox.Text = LabelSubLabelSettings.BarcodeMagnitude.ToString();
            BarcodeFormatTextBox.Text = LabelSubLabelSettings.Format;

            if (LabelSubLabelSettings.Font != null) {
                SubstrateBarcodeFontDialog.Font = LabelSubLabelSettings.Font;
                BarcodeFontTextBox.Text = $"{SubstrateBarcodeFontDialog.Font.Name} {SubstrateBarcodeFontDialog.Font.SizeInPoints}pt";
            }

            BarcodePostionXTextBox.Text = LabelSubLabelSettings.StringPosX.ToString();
            BarcodePostionYTextBox.Text = LabelSubLabelSettings.StringPosY.ToString();
            BarcodeCenterCheckBox.Checked = LabelSubLabelSettings.AlignStringCenter;

            if (BarcodeCenterCheckBox.Checked) {
                BarcodePostionXTextBox.Enabled = false;
            }
            else {
                BarcodePostionXTextBox.Enabled = true;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            double SizeX, SizeY, OffsetX, OffsetY, IntervalX, IntervalY, BarcodeHeight, BarcodeMagnitude, StringPosX, StringPosY;
            int NumLabelsX, NumLabelsY, HeaderPosX, HeaderPosY, FooterPosX = 0, FooterPosY = 0;

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

                BarcodeHeight = double.Parse(BarcodeHeightTextBox.Text);
                StringPosX = double.Parse(BarcodePostionXTextBox.Text);
                StringPosY = double.Parse(BarcodePostionYTextBox.Text);
                BarcodeMagnitude = double.Parse(BarcodeMagnitudeTextBox.Text);
            } catch (Exception ex) {
                MessageBox.Show($"入力値が不正です。{Environment.NewLine}{ex.Message}");
                this.DialogResult = DialogResult.None;
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
            LabelSubPageSettings.HeaderFooterFont = SubstratePrintFontDialog.Font;

            LabelSubLabelSettings.BarcodeHeight = BarcodeHeight;
            LabelSubLabelSettings.BarcodePosX = 0; // BarcodePosX is not set in the original VB.NET code
            LabelSubLabelSettings.BarcodePosY = 0; // BarcodePosY is not set in the original VB.NET code
            LabelSubLabelSettings.BarcodeMagnitude = BarcodeMagnitude;
            LabelSubLabelSettings.Format = BarcodeFormatTextBox.Text;
            LabelSubLabelSettings.Font = SubstrateBarcodeFontDialog.Font;
            LabelSubLabelSettings.StringPosX = StringPosX;
            LabelSubLabelSettings.StringPosY = StringPosY;
            LabelSubLabelSettings.AlignStringCenter = BarcodeCenterCheckBox.Checked;
            LabelSubLabelSettings.NumLabels = int.Parse(BarcodeQuantityTextBox.Text);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnHeaderFooterFont_Click(object sender, EventArgs e) {
            DialogResult r = SubstratePrintFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) return;
            HeaderFooterFontTextBox.Text = $"{SubstratePrintFontDialog.Font.Name} {SubstratePrintFontDialog.Font.SizeInPoints}pt";
        }

        private void BtnBarcodeFont_Click(object sender, EventArgs e) {
            DialogResult r = SubstrateBarcodeFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) return;
            BarcodeFontTextBox.Text = $"{SubstrateBarcodeFontDialog.Font.Name} {SubstrateBarcodeFontDialog.Font.SizeInPoints}pt";
        }

        private void BarcodeCenterCheckBox_CheckedChanged(object sender, EventArgs e) {
            if (BarcodeCenterCheckBox.Checked) {
                BarcodePostionXTextBox.Enabled = false;
            }
            else {
                BarcodePostionXTextBox.Enabled = true;
            }
        }

        private void SubstratePrintSetting_Load(object sender, EventArgs e) { FrmPageSettings_Label_Load(sender, e); }
    }
}
