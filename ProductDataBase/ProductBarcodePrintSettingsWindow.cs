using LabelGenerator;

namespace ProductDataBase {
    public partial class ProductBarcodePrintSettingsWindow : Form {

        private CBarcodeProLabelSettings BarcodeProLabelSettings;
        private CBarcodeProPageSettings BarcodeProPageSettings;

        public ProductBarcodePrintSettingsWindow() {
            InitializeComponent();

            BarcodeProLabelSettings = new CBarcodeProLabelSettings();
            BarcodeProPageSettings = new CBarcodeProPageSettings();
        }

        private void PageSettingsBarcodeLoad(object sender, EventArgs e) {
            BarcodeProPageSettings = ((ProductRegistration2Window)Owner!).SettingsBarcodePro!.BarcodeProPageSettings;
            BarcodeLabelWidthTextBox.Text = BarcodeProPageSettings.SizeX.ToString();
            BarcodeLabelHeightTextBox.Text = BarcodeProPageSettings.SizeY.ToString();
            BarcodeQuantityXTextBox.Text = BarcodeProPageSettings.NumLabelsX.ToString();
            BarcodeQuantityYTextBox.Text = BarcodeProPageSettings.NumLabelsY.ToString();
            BarcodePageOffsetXTextBox.Text = BarcodeProPageSettings.OffsetX.ToString();
            BarcodePageOffsetYTextBox.Text = BarcodeProPageSettings.OffsetY.ToString();
            BarcodeLabelIntervalXTextBox.Text = BarcodeProPageSettings.IntervalX.ToString();
            BarcodeLabelIntervalYTextBox.Text = BarcodeProPageSettings.IntervalY.ToString();
            BarcodeHeaderStringTextBox.Text = BarcodeProPageSettings.HeaderString;
            BarcodeHeaderPostionXTextBox.Text = BarcodeProPageSettings.HeaderPos.X.ToString();
            BarcodeHeaderPostionYTextBox.Text = BarcodeProPageSettings.HeaderPos.Y.ToString();

            if (BarcodeProPageSettings.HeaderFooterFont != null) {
                BarcodeHeaderFontDialog.Font = BarcodeProPageSettings.HeaderFooterFont;
                BarcodeHeaderFooterFontTextBox.Text = $"{BarcodeHeaderFontDialog.Font.Name} {BarcodeHeaderFontDialog.Font.SizeInPoints}pt";
            }

            BarcodeProLabelSettings = ((ProductRegistration2Window)Owner).SettingsBarcodePro!.BarcodeProLabelSettings;
            BarcodeHeightTextBox.Text = BarcodeProLabelSettings.BarcodeHeight.ToString();
            BarcodeMagnitudeTextBox.Text = BarcodeProLabelSettings.BarcodeMagnitude.ToString();
            BarcodeQuantityTextBox.Text = BarcodeProLabelSettings.NumLabels.ToString();
            BarcodeFormatTextBox.Text = BarcodeProLabelSettings.Format;

            if (BarcodeProLabelSettings.Font != null) {
                BarcodeFontDialog.Font = BarcodeProLabelSettings.Font;
                BarcodeFontTextBox.Text = $"{BarcodeFontDialog.Font.Name} {BarcodeFontDialog.Font.SizeInPoints}pt";
            }

            BarcodePostionXTextBox.Text = BarcodeProLabelSettings.StringPosX.ToString();
            BarcodePostionYTextBox.Text = BarcodeProLabelSettings.StringPosY.ToString();
            BarcodeCenterCheckBox.Checked = BarcodeProLabelSettings.AlignStringCenter;

            if (BarcodeCenterCheckBox.Checked) {
                BarcodePostionXTextBox.Enabled = false;
            }
            else {
                BarcodePostionXTextBox.Enabled = true;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            double SizeX, SizeY, OffsetX, OffsetY, IntervalX, IntervalY, BarcodeHeight, BarcodeMagnitude, StringPosX, StringPosY;
            int NumLabelsX, NumLabelsY, HeaderPosX, HeaderPosY, FooterPosX = 0, FooterPosY = 0, NumLabels;

            try {
                SizeX = double.Parse(BarcodeLabelWidthTextBox.Text);
                SizeY = double.Parse(BarcodeLabelHeightTextBox.Text);
                NumLabelsX = int.Parse(BarcodeQuantityXTextBox.Text);
                NumLabelsY = int.Parse(BarcodeQuantityYTextBox.Text);
                OffsetX = double.Parse(BarcodePageOffsetXTextBox.Text);
                OffsetY = double.Parse(BarcodePageOffsetYTextBox.Text);
                IntervalX = double.Parse(BarcodeLabelIntervalXTextBox.Text);
                IntervalY = double.Parse(BarcodeLabelIntervalYTextBox.Text);
                HeaderPosX = int.Parse(BarcodeHeaderPostionXTextBox.Text);
                HeaderPosY = int.Parse(BarcodeHeaderPostionYTextBox.Text);

                BarcodeHeight = double.Parse(BarcodeHeightTextBox.Text);
                StringPosX = double.Parse(BarcodePostionXTextBox.Text);
                StringPosY = double.Parse(BarcodePostionYTextBox.Text);
                BarcodeMagnitude = double.Parse(BarcodeMagnitudeTextBox.Text);
                NumLabels = int.Parse(BarcodeQuantityTextBox.Text);
            } catch (Exception ex) {
                MessageBox.Show($"入力値が不正です。{Environment.NewLine}{ex.Message}");
                DialogResult = DialogResult.None;
                return;
            }

            BarcodeProPageSettings.SizeX = SizeX;
            BarcodeProPageSettings.SizeY = SizeY;
            BarcodeProPageSettings.NumLabelsX = NumLabelsX;
            BarcodeProPageSettings.NumLabelsY = NumLabelsY;
            BarcodeProPageSettings.OffsetX = OffsetX;
            BarcodeProPageSettings.OffsetY = OffsetY;
            BarcodeProPageSettings.IntervalX = IntervalX;
            BarcodeProPageSettings.IntervalY = IntervalY;
            BarcodeProPageSettings.HeaderPos = new Point(HeaderPosX, HeaderPosY);
            BarcodeProPageSettings.FooterPos = new Point(FooterPosX, FooterPosY);
            BarcodeProPageSettings.HeaderString = BarcodeHeaderStringTextBox.Text;
            BarcodeProPageSettings.HeaderFooterFont = BarcodeHeaderFontDialog.Font;

            BarcodeProLabelSettings.BarcodeHeight = BarcodeHeight;
            BarcodeProLabelSettings.BarcodePosX = 0; // BarcodePosX is not set in the original VB.NET code
            BarcodeProLabelSettings.BarcodePosY = 0; // BarcodePosY is not set in the original VB.NET code
            BarcodeProLabelSettings.BarcodeMagnitude = BarcodeMagnitude;
            BarcodeProLabelSettings.Format = BarcodeFormatTextBox.Text;
            BarcodeProLabelSettings.Font = BarcodeFontDialog.Font;
            BarcodeProLabelSettings.StringPosX = StringPosX;
            BarcodeProLabelSettings.StringPosY = StringPosY;
            BarcodeProLabelSettings.AlignStringCenter = BarcodeCenterCheckBox.Checked;
            BarcodeProLabelSettings.NumLabels = int.Parse(BarcodeQuantityTextBox.Text);

            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnHeaderFooterFont_Click(object sender, EventArgs e) {
            DialogResult r = BarcodeHeaderFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) return;
            BarcodeHeaderFooterFontTextBox.Text = $"{BarcodeHeaderFontDialog.Font.Name} {BarcodeHeaderFontDialog.Font.SizeInPoints}pt";
        }

        private void BtnBarcodeFont_Click(object sender, EventArgs e) {
            DialogResult r = BarcodeFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) return;
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
    }
}
