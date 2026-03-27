using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using ProductDatabase.Models;
using static ProductDatabase.Data.ProductRepository;
using static ProductDatabase.Print.PrintManager;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase {
    public partial class PrintSettingsWindow : Form {

        public DocumentPrintSettings DocumentPrintSettings { get; set; } = new DocumentPrintSettings();
        private string _documentPrintSettingFilePath = string.Empty;
        private JsonSerializerOptions? _jsonSerializerOptions;

        public ProductMaster ProductMaster { get; set; } = new ProductMaster();
        public SubstrateMaster SubstrateMaster { get; set; } = new SubstrateMaster();
        public AppSettings AppSettings { get; set; } = new AppSettings();

        private bool _isLabelPrint;
        private bool _isBarcodePrint;
        private bool _isNameplatePrint;

        public PrintSettingsWindow() {
            InitializeComponent();
        }

        // 呼び出し元ウィンドウの種類に応じて印刷設定とパスを読み込みPropertyGridに反映する
        private void LoadSettings() {
            Font = new System.Drawing.Font(AppSettings.FontName, AppSettings.FontSize);
            switch (Owner) {
                case ProductRegistration2Window productWindow:
                    _isLabelPrint = ProductMaster.IsLabelPrint;
                    _isBarcodePrint = ProductMaster.IsBarcodePrint;
                    _isNameplatePrint = ProductMaster.IsNameplatePrint;
                    LoadSettingsFromWindow(productWindow.ProductPrintSettings, productWindow.PrintSettingPath);
                    break;
                case SubstrateRegistrationWindow substrateWindow:
                    _isLabelPrint = SubstrateMaster.IsLabelPrint;
                    LoadSettingsFromWindow(substrateWindow.SubstratePrintSettings, substrateWindow.PrintSettingPath);
                    break;
                case RePrintWindow rePrintWindow:
                    _isLabelPrint = ProductMaster.IsLabelPrint;
                    _isBarcodePrint = ProductMaster.IsBarcodePrint;
                    _isNameplatePrint = ProductMaster.IsNameplatePrint;
                    LoadSettingsFromWindow(rePrintWindow.ProductPrintSettings, rePrintWindow.PrintSettingPath);
                    break;
                default:
                    MessageBox.Show("この画面を開くには正しいウィンドウから開いてください。");
                    Close();
                    break;
            }
        }
        // 設定ファイルパスとDocumentPrintSettingsをセットし対象シリアルタイプの設定をPropertyGridに表示する
        private void LoadSettingsFromWindow(DocumentPrintSettings settings, string filePath) {
            DocumentPrintSettings = settings;
            _documentPrintSettingFilePath = filePath;

            PrintPropertyGrid.SelectedObject = CurrentSerialType switch {
                SerialType.Label => DocumentPrintSettings.LabelPrintSettings,
                SerialType.Barcode => DocumentPrintSettings.BarcodePrintSettings,
                SerialType.Nameplate => DocumentPrintSettings.NameplatePrintSettings,
                SerialType.Substrate => DocumentPrintSettings.LabelPrintSettings,
                _ => null
            };
        }
        // 現在の印刷設定をJSONにシリアライズして設定ファイルに保存しウィンドウを閉じる
        private void SaveProductPrintSettings() {
            try {
                DocumentPrintSettings.SetSettingsType(_isLabelPrint, _isBarcodePrint, _isNameplatePrint);
                _jsonSerializerOptions ??= new JsonSerializerOptions {
                    WriteIndented = true,
                    PropertyNamingPolicy = null,
                    Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                };
                var jsonString = JsonSerializer.Serialize(DocumentPrintSettings, _jsonSerializerOptions);
                File.WriteAllText(_documentPrintSettingFilePath, jsonString);
                Close();
            } catch (Exception ex) {
                MessageBox.Show($"設定の保存中にエラーが発生しました。{Environment.NewLine}{ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProductBarcodePrintSetting_Load(object sender, EventArgs e) { LoadSettings(); }
        private void OKButton_Click(object sender, EventArgs e) { SaveProductPrintSettings(); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }

    }
}
