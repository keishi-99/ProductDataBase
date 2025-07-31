using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using static ProductDatabase.Other.PrintOptions;

namespace ProductDatabase {
    public partial class SubstratePrintSettingsWindow : Form {

        public DocumentPrintSettings SubstratePrintSettings { get; set; } = new DocumentPrintSettings();
        public static readonly string s_substratePrintSettingFilePath = Path.Combine(Environment.CurrentDirectory, "config", "Substrate", "SubstrateConfig.json");
        private JsonSerializerOptions? _jsonSerializerOptions;

        public SubstratePrintSettingsWindow() {
            InitializeComponent();
        }

        private void LoadSettings() {
            if (Owner is SubstrateRegistrationWindow substrateWindow) {
                LoadSettingsFromWindow(substrateWindow.SubstratePrintSettings);
            }
            else {
                MessageBox.Show("この画面を開くには正しいウィンドウから開いてください。");
                Close();
            }
        }
        private void LoadSettingsFromWindow(DocumentPrintSettings settings) {
            SubstratePrintSettings = settings;

            if (SubstratePrintSettings.LabelPageSettings == null || SubstratePrintSettings.LabelLayoutSettings == null) {
                MessageBox.Show("設定が読み込まれませんでした。");
                Close();
                return;
            }

            SetPageSettings(SubstratePrintSettings.LabelPageSettings);
            SetLabelSettings(SubstratePrintSettings.LabelLayoutSettings);
        }

        private void SetPageSettings(PrintPageSettings pageSettings) {
            LabelWidthTextBox.Text = pageSettings.LabelWidth.ToString();
            LabelHeightTextBox.Text = pageSettings.LabelHeight.ToString();
            LabelsPerRowTextBox.Text = pageSettings.LabelsPerRow.ToString();
            LabelsPerColumnTextBox.Text = pageSettings.LabelsPerColumn.ToString();
            LabelMarginXTextBox.Text = pageSettings.MarginX.ToString();
            LabelMarginYTextBox.Text = pageSettings.MarginY.ToString();
            LabelIntervalXTextBox.Text = pageSettings.IntervalX.ToString();
            LabelIntervalYTextBox.Text = pageSettings.IntervalY.ToString();
            HeaderTextTextBox.Text = pageSettings.HeaderTextFormat;
            HeaderPositionXTextBox.Text = pageSettings.HeaderPositionX.ToString();
            HeaderPositionYTextBox.Text = pageSettings.HeaderPositionY.ToString();

            if (pageSettings.HeaderFont != null) {
                HeaderFontDialog.Font = pageSettings.HeaderFont;
                HeaderFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
            }
        }
        private void SetLabelSettings(PrintLayoutSettings labelSettings) {
            CopiesPerLabelTextBox.Text = labelSettings.CopiesPerLabel.ToString();
            LabelFormatTextBox.Text = labelSettings.TextFormat;

            if (labelSettings.TextFont != null) {
                LabelFontDialog.Font = labelSettings.TextFont;
                PrintTextFontTextBox.Text = $"{LabelFontDialog.Font.Name} {LabelFontDialog.Font.SizeInPoints}pt";
            }

            LabelTextPositionXTextBox.Text = labelSettings.TextPositionX.ToString();
            LabelTextPositionYTextBox.Text = labelSettings.TextPositionY.ToString();
            AlignTextXCenterCheckBox.Checked = labelSettings.AlignTextCenterX;
            AlignTextYCenterCheckBox.Checked = labelSettings.AlignTextCenterY;

            LabelTextPositionXTextBox.Enabled = !AlignTextXCenterCheckBox.Checked;
        }

        private bool ValidateInput() {
            // double 型の検証
            var doubleTextBoxes = new Dictionary<TextBox, string>
            {
                { LabelWidthTextBox, "ラベル幅" },
                { LabelHeightTextBox, "ラベル高さ" },
                { LabelMarginXTextBox, "マージンX" },
                { LabelMarginYTextBox, "マージンY" },
                { LabelIntervalXTextBox, "間隔X" },
                { LabelIntervalYTextBox, "間隔Y" },
                { LabelTextPositionXTextBox, "フォント位置X" },
                { LabelTextPositionYTextBox, "フォント位置Y" },
                { HeaderPositionXTextBox, "ヘッダー位置X" },
                { HeaderPositionYTextBox, "ヘッダー位置Y" },
            };

            foreach (var textBox in doubleTextBoxes) {
                if (!double.TryParse(textBox.Key.Text, out var value)) {
                    MessageBox.Show($"{textBox.Value} に不正な数値が入力されました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // 範囲チェック
                if (value < 0) {
                    MessageBox.Show($"{textBox.Value} は 0 以上である必要があります。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            // int 型の検証
            var intTextBoxes = new Dictionary<TextBox, string>
            {
                {LabelsPerRowTextBox , "ラベル行数" },
                { LabelsPerColumnTextBox, "ラベル列数" },
                { CopiesPerLabelTextBox, "ラベルごとのコピー数" },
            };

            foreach (var textBox in intTextBoxes) {
                if (!int.TryParse(textBox.Key.Text, out var value)) {
                    MessageBox.Show($"{textBox.Value} に不正な整数が入力されました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                // 範囲チェック
                if (value <= 0) {
                    MessageBox.Show($"{textBox.Value} は 1 以上である必要があります。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            return true;
        }
        private void UpdateSubstratePrintSettings() {
            if (SubstratePrintSettings.LabelPageSettings != null) {
                SubstratePrintSettings.LabelPageSettings.LabelWidth = ParseDouble(LabelWidthTextBox.Text);
                SubstratePrintSettings.LabelPageSettings.LabelHeight = ParseDouble(LabelHeightTextBox.Text);
                SubstratePrintSettings.LabelPageSettings.LabelsPerRow = ParseInt(LabelsPerRowTextBox.Text);
                SubstratePrintSettings.LabelPageSettings.LabelsPerColumn = ParseInt(LabelsPerColumnTextBox.Text);
                SubstratePrintSettings.LabelPageSettings.MarginX = ParseDouble(LabelMarginXTextBox.Text);
                SubstratePrintSettings.LabelPageSettings.MarginY = ParseDouble(LabelMarginYTextBox.Text);
                SubstratePrintSettings.LabelPageSettings.IntervalX = ParseDouble(LabelIntervalXTextBox.Text);
                SubstratePrintSettings.LabelPageSettings.IntervalY = ParseDouble(LabelIntervalYTextBox.Text);
                SubstratePrintSettings.LabelPageSettings.HeaderPositionX = ParseDouble(HeaderPositionXTextBox.Text);
                SubstratePrintSettings.LabelPageSettings.HeaderPositionY = ParseDouble(HeaderPositionYTextBox.Text);
                SubstratePrintSettings.LabelPageSettings.HeaderTextFormat = HeaderTextTextBox.Text;
                SubstratePrintSettings.LabelPageSettings.HeaderFont = HeaderFontDialog.Font;
            }

            if (SubstratePrintSettings.LabelLayoutSettings != null) {
                SubstratePrintSettings.LabelLayoutSettings.TextFont = LabelFontDialog.Font;
                SubstratePrintSettings.LabelLayoutSettings.TextPositionX = ParseDouble(LabelTextPositionXTextBox.Text);
                SubstratePrintSettings.LabelLayoutSettings.TextPositionY = ParseDouble(LabelTextPositionYTextBox.Text);
                SubstratePrintSettings.LabelLayoutSettings.AlignTextCenterX = AlignTextXCenterCheckBox.Checked;
                SubstratePrintSettings.LabelLayoutSettings.AlignTextCenterY = AlignTextYCenterCheckBox.Checked;
                SubstratePrintSettings.LabelLayoutSettings.CopiesPerLabel = ParseInt(CopiesPerLabelTextBox.Text);
                SubstratePrintSettings.LabelLayoutSettings.TextFormat = LabelFormatTextBox.Text;
                SubstratePrintSettings.LabelLayoutSettings.TextFont = LabelFontDialog.Font;
                SubstratePrintSettings.LabelLayoutSettings.TextPositionX = ParseDouble(LabelTextPositionXTextBox.Text);
                SubstratePrintSettings.LabelLayoutSettings.TextPositionY = ParseDouble(LabelTextPositionYTextBox.Text);
                SubstratePrintSettings.LabelLayoutSettings.AlignTextCenterX = AlignTextXCenterCheckBox.Checked;
                SubstratePrintSettings.LabelLayoutSettings.AlignTextCenterY = AlignTextYCenterCheckBox.Checked;
                SubstratePrintSettings.LabelLayoutSettings.CopiesPerLabel = ParseInt(CopiesPerLabelTextBox.Text);
            }
        }

        private static double ParseDouble(string text) {
            if (double.TryParse(text, out var value)) {
                return value;
            }
            else {
                MessageBox.Show($"不正な数値が入力されました: {text}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }
        private static int ParseInt(string text) {
            if (int.TryParse(text, out var value)) {
                return value;
            }
            else {
                MessageBox.Show($"不正な数値が入力されました: {text}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }
        private void SaveSubstratePrintSettings() {
            try {
                // JsonSerializerOptions のインスタンスをキャッシュ
                _jsonSerializerOptions ??= new JsonSerializerOptions {
                    WriteIndented = true,
                    PropertyNamingPolicy = null,
                    Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                };
                var jsonString = JsonSerializer.Serialize(SubstratePrintSettings, _jsonSerializerOptions);
                File.WriteAllText(s_substratePrintSettingFilePath, jsonString);
                DialogResult = DialogResult.OK;
                Close();
            } catch (Exception ex) {
                MessageBox.Show($"設定の保存中にエラーが発生しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
            }
        }

        private void BtnHeaderFont_Click(object sender, EventArgs e) {
            var r = HeaderFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }
            HeaderFontTextBox.Text = $"{HeaderFontDialog.Font.Name} {HeaderFontDialog.Font.SizeInPoints}pt";
        }
        private void PrintTextFontButton_Click(object sender, EventArgs e) {
            var r = LabelFontDialog.ShowDialog();
            if (r == DialogResult.Cancel) {
                return;
            }
            PrintTextFontTextBox.Text = $"{LabelFontDialog.Font.Name} {LabelFontDialog.Font.SizeInPoints}pt";
        }
        private void PrintTextXCenterCheckBox_CheckedChanged(object sender, EventArgs e) { LabelTextPositionXTextBox.Enabled = !AlignTextXCenterCheckBox.Checked; }
        private void PrintTextYCenterCheckBox_CheckedChanged(object sender, EventArgs e) { LabelTextPositionYTextBox.Enabled = !AlignTextYCenterCheckBox.Checked; }
        private void SubstratePrintSetting_Load(object sender, EventArgs e) { LoadSettings(); }
        private void OKButton_Click(object sender, EventArgs e) {
            if (!ValidateInput()) {
                return;
            }
            UpdateSubstratePrintSettings();
            SaveSubstratePrintSettings();
        }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
