using ProductDatabase.Substrate;

namespace ProductDatabase {
    public partial class SubstrateLabelSettingsWindow : Form {

        private CLabelSubLabelSettings _labelSubLabelSettings;
        private CLabelSubPageSettings _labelSubPageSettings;

        public SubstrateLabelSettingsWindow() {
            InitializeComponent();

            _labelSubLabelSettings = new CLabelSubLabelSettings();
            _labelSubPageSettings = new CLabelSubPageSettings();
        }

        private void PageSettingsLabelLoad(object sender, EventArgs e) {
            _labelSubPageSettings = ((SubstrateRegistrationWindow)Owner!).SettingsLabelSub!.LabelSubPageSettings;
            LabelWidthTextBox.Text = _labelSubPageSettings.LabelWidth.ToString();
            LabelHeightTextBox.Text = _labelSubPageSettings.LabelHeight.ToString();
            QuantityXTextBox.Text = _labelSubPageSettings.LabelCountX.ToString();
            QuantityYTextBox.Text = _labelSubPageSettings.LabelCountY.ToString();
            PageOffsetXTextBox.Text = _labelSubPageSettings.MarginX.ToString();
            PageOffsetYTextBox.Text = _labelSubPageSettings.MarginY.ToString();
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
            PrintTextQuantityTextBox.Text = _labelSubLabelSettings.NumLabels.ToString();
            PrintTextFormatTextBox.Text = _labelSubLabelSettings.Format;

            if (_labelSubLabelSettings.Font != null) {
                TextFontDialog.Font = _labelSubLabelSettings.Font;
                PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
            }

            PrintTextPostionXTextBox.Text = _labelSubLabelSettings.StringPosX.ToString();
            PrintTextPostionYTextBox.Text = _labelSubLabelSettings.StringPosY.ToString();
            PrintTextXCenterCheckBox.Checked = _labelSubLabelSettings.AlignStringXCenter;
            PrintTextYCenterCheckBox.Checked = _labelSubLabelSettings.AlignStringYCenter;

            PrintTextPostionXTextBox.Enabled = !PrintTextXCenterCheckBox.Checked;
        }

        private void BtnOK_Click(object sender, EventArgs e) {
            double sizeX, sizeY, offsetX, offsetY, intervalX, intervalY, stringPosX, stringPosY;
            int numLabelsX, numLabelsY, headerPosX, headerPosY, numLabels;

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

                stringPosX = double.Parse(PrintTextPostionXTextBox.Text);
                stringPosY = double.Parse(PrintTextPostionYTextBox.Text);
                numLabels = int.Parse(PrintTextQuantityTextBox.Text);
            } catch (Exception ex) {
                MessageBox.Show($"入力値が不正です。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
                return;
            }

            _labelSubPageSettings.LabelWidth = sizeX;
            _labelSubPageSettings.LabelHeight = sizeY;
            _labelSubPageSettings.LabelCountX = numLabelsX;
            _labelSubPageSettings.LabelCountY = numLabelsY;
            _labelSubPageSettings.MarginX = offsetX;
            _labelSubPageSettings.MarginY = offsetY;
            _labelSubPageSettings.IntervalX = intervalX;
            _labelSubPageSettings.IntervalY = intervalY;
            _labelSubPageSettings.HeaderPos = new Point(headerPosX, headerPosY);
            _labelSubPageSettings.HeaderString = HeaderStringTextBox.Text;
            _labelSubPageSettings.HeaderFooterFont = HeaderFontDialog.Font;

            _labelSubLabelSettings.Format = PrintTextFormatTextBox.Text;
            _labelSubLabelSettings.Font = TextFontDialog.Font;
            _labelSubLabelSettings.StringPosX = stringPosX;
            _labelSubLabelSettings.StringPosY = stringPosY;
            _labelSubLabelSettings.AlignStringXCenter = PrintTextXCenterCheckBox.Checked;
            _labelSubLabelSettings.AlignStringYCenter = PrintTextYCenterCheckBox.Checked;
            _labelSubLabelSettings.NumLabels = int.Parse(PrintTextQuantityTextBox.Text);

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
        private void PrintTextFontButton_Click(object sender, EventArgs e) {
            var r = TextFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }

            PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
        }
        private void PrintTextXCenterCheckBox_CheckedChanged(object sender, EventArgs e) { PrintTextPostionXTextBox.Enabled = !PrintTextXCenterCheckBox.Checked; }
        private void PrintTextYCenterCheckBox_CheckedChanged(object sender, EventArgs e) { PrintTextPostionYTextBox.Enabled = !PrintTextYCenterCheckBox.Checked; }
        private void SubstratePrintSetting_Load(object sender, EventArgs e) { PageSettingsLabelLoad(sender, e); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
