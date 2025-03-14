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

        public CSettingsLabelPro SettingsLabelPro { get; set; } = new CSettingsLabelPro();
        private string _labelSettingFilePath = string.Empty;

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
            _labelSettingFilePath = window.labelSettingFilePath;

            SetPageSettings(pageSettings);
            SetLabelSettings(labelSettings);
        }
        private void LoadSettingsFromRePaintWindow(RePrintWindow window) {
            var pageSettings = window.SettingsLabelPro.LabelProPageSettings;
            var labelSettings = window.SettingsLabelPro.LabelProLabelSettings;
            _labelSettingFilePath = window.labelSettingFilePath;

            SetPageSettings(pageSettings);
            SetLabelSettings(labelSettings);
        }
        private void SetLabelSettings(CLabelProLabelSettings labelSettings) {
            PrintTextQuantityTextBox.Text = labelSettings.NumLabels.ToString();
            PrintTextFormatTextBox.Text = labelSettings.Format;

            if (labelSettings.Font != null) {
                TextFontDialog.Font = labelSettings.Font;
                PrintTextFontTextBox.Text = $"{TextFontDialog.Font.Name} {TextFontDialog.Font.SizeInPoints}pt";
            }

            PrintTextPostionXTextBox.Text = labelSettings.StringPosX.ToString();
            PrintTextPostionYTextBox.Text = labelSettings.StringPosY.ToString();
            TextXCenterCheckBox.Checked = labelSettings.AlignStringXCenter;
            TextYCenterCheckBox.Checked = labelSettings.AlignStringYCenter;

            PrintTextPostionXTextBox.Enabled = !TextXCenterCheckBox.Checked;
        }
        private void SetPageSettings(CLabelProPageSettings pageSettings) {
            LabelWidthTextBox.Text = pageSettings.LabelWidth.ToString();
            LabelHeightTextBox.Text = pageSettings.LabelHeight.ToString();
            QuantityXTextBox.Text = pageSettings.LabelCountX.ToString();
            QuantityYTextBox.Text = pageSettings.LabelCountY.ToString();
            PageOffsetXTextBox.Text = pageSettings.MarginX.ToString();
            PageOffsetYTextBox.Text = pageSettings.MarginY.ToString();
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

            _labelProPageSettings.LabelWidth = sizeX;
            _labelProPageSettings.LabelHeight = sizeY;
            _labelProPageSettings.LabelCountX = numLabelsX;
            _labelProPageSettings.LabelCountY = numLabelsY;
            _labelProPageSettings.MarginX = offsetX;
            _labelProPageSettings.MarginY = offsetY;
            _labelProPageSettings.IntervalX = intervalX;
            _labelProPageSettings.IntervalY = intervalY;
            _labelProPageSettings.HeaderPos = new Point(headerPosX, headerPosY);
            _labelProPageSettings.HeaderString = HeaderStringTextBox.Text;
            _labelProPageSettings.HeaderFooterFont = HeaderFontDialog.Font;

            _labelProLabelSettings.Format = PrintTextFormatTextBox.Text;
            _labelProLabelSettings.Font = TextFontDialog.Font;
            _labelProLabelSettings.StringPosX = stringPosX;
            _labelProLabelSettings.StringPosY = stringPosY;
            _labelProLabelSettings.AlignStringXCenter = TextXCenterCheckBox.Checked;
            _labelProLabelSettings.AlignStringYCenter = TextYCenterCheckBox.Checked;
            _labelProLabelSettings.NumLabels = int.Parse(PrintTextQuantityTextBox.Text);

            SettingsLabelPro.LabelProPageSettings = _labelProPageSettings;
            SettingsLabelPro.LabelProLabelSettings = _labelProLabelSettings;

            try {
                using var swLabel = new StreamWriter(_labelSettingFilePath, false, new System.Text.UTF8Encoding(false));
                var serializerLabel = new System.Xml.Serialization.XmlSerializer(typeof(CSettingsLabelPro));
                serializerLabel.Serialize(swLabel, SettingsLabelPro);
            } catch (Exception ex) {
                MessageBox.Show($"設定の保存中にエラーが発生しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
                return;
            }

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
        private void TextXCenterCheckBox_CheckedChanged(object sender, EventArgs e) { PrintTextPostionXTextBox.Enabled = !TextXCenterCheckBox.Checked; }
        private void TextYCenterCheckBox_CheckedChanged(object sender, EventArgs e) { PrintTextPostionYTextBox.Enabled = !TextYCenterCheckBox.Checked; }
        private void ProductPrintSetting_Load(object sender, EventArgs e) { PageSettingsLabelLoad(sender, e); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
