using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using static ProductDatabase.MainWindow;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase {
    public partial class PrintSettingsWindow : Form {

        public DocumentPrintSettings DocumentPrintSettings { get; set; } = new DocumentPrintSettings();
        private string _documentPrintSettingFilePath = string.Empty;
        private JsonSerializerOptions? _jsonSerializerOptions;

        public ProductInformation ProductInfo { get; set; } = new ProductInformation();

        public string serialType = string.Empty;

        public PrintSettingsWindow() {
            InitializeComponent();
        }

        private void LoadSettings() {
            Font = new System.Drawing.Font(ProductInfo.FontName, ProductInfo.FontSize);
            switch (Owner) {
                case ProductRegistration2Window productWindow:
                    ProductInfo = productWindow.ProductInfo;
                    LoadSettingsFromWindow(productWindow.ProductPrintSettings, productWindow.printSettingPath);
                    break;
                case SubstrateRegistrationWindow substrateWindow:
                    ProductInfo = substrateWindow.ProductInfo;
                    LoadSettingsFromWindow(substrateWindow.SubstratePrintSettings, substrateWindow.printSettingPath);
                    break;
                case RePrintWindow rePrintWindow:
                    LoadSettingsFromWindow(rePrintWindow.ProductPrintSettings, rePrintWindow.printSettingPath);
                    break;
                default:
                    MessageBox.Show("この画面を開くには正しいウィンドウから開いてください。");
                    Close();
                    break;
            }
        }
        private void LoadSettingsFromWindow(DocumentPrintSettings settings, string filePath) {
            DocumentPrintSettings = settings;
            _documentPrintSettingFilePath = filePath;

            PrintPropertyGrid.SelectedObject = serialType switch {
                "Label" => DocumentPrintSettings.LabelPrintSettings,
                "Barcode" => DocumentPrintSettings.BarcodePrintSettings,
                "Substrate" => DocumentPrintSettings.LabelPrintSettings,
                _ => null
            };
        }
        private void SaveProductPrintSettings() {
            try {
                DocumentPrintSettings.SetSettingsType(ProductInfo.IsLabelPrint, ProductInfo.IsBarcodePrint);
                // JsonSerializerOptions のインスタンスをキャッシュ
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
