using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ProductDatabase.Print {
    // 印刷設定クラス群（DocumentPrintSettings / PrintSettingsBase / 各印刷設定 / FontJsonConverter）
    public class PrintOptions {

        public class DocumentPrintSettings {
            public LabelPrintSettings? LabelPrintSettings { get; set; }
            public BarcodePrintSettings? BarcodePrintSettings { get; set; }
            public NameplatePrintSettings? NameplatePrintSettings { get; set; }

            public DocumentPrintSettings() {
                LabelPrintSettings = new LabelPrintSettings();
                BarcodePrintSettings = new BarcodePrintSettings();
                NameplatePrintSettings = new NameplatePrintSettings();
            }

            // 印刷フラグに応じてラベル・バーコード・銘版の設定インスタンスをnullまたは既存値に設定する
            public void SetSettingsType(bool isLabelPrint, bool isBarcodePrint, bool isNameplatePrint) {
                LabelPrintSettings = isLabelPrint ? (LabelPrintSettings ?? new LabelPrintSettings()) : null;
                BarcodePrintSettings = isBarcodePrint ? (BarcodePrintSettings ?? new BarcodePrintSettings()) : null;
                NameplatePrintSettings = isNameplatePrint ? (NameplatePrintSettings ?? new NameplatePrintSettings()) : null;
            }
        }

        public abstract class PrintSettingsBase {
            [Category("\t用紙設定"), DisplayName("ラベル幅 (mm)")]
            public double LabelWidth { get; set; }
            [Category("\t用紙設定"), DisplayName("ラベル高さ (mm)")]
            public double LabelHeight { get; set; }
            [Category("\t用紙設定"), DisplayName("配置数 (行)")]
            public int LabelsPerRow { get; set; }
            [Category("\t用紙設定"), DisplayName("配置数 (列)")]
            public int LabelsPerColumn { get; set; }
            [Category("\t用紙設定"), DisplayName("余白左 (mm)")]
            public double MarginX { get; set; }
            [Category("\t用紙設定"), DisplayName("余白上 (mm)")]
            public double MarginY { get; set; }
            [Category("\t用紙設定"), DisplayName("ラベル間隔 横 (mm)")]
            public double IntervalX { get; set; }
            [Category("\t用紙設定"), DisplayName("ラベル間隔 縦 (mm)")]
            public double IntervalY { get; set; }
            [Category("\t用紙設定"), DisplayName("ヘッダー開始位置 左 (mm)")]
            public double HeaderPositionX { get; set; }
            [Category("\t用紙設定"), DisplayName("ヘッダー開始位置 上 (mm)")]
            public double HeaderPositionY { get; set; }
            [Category("\t用紙設定"), DisplayName("ヘッダー フォーマット"), Description("{D}(印刷日)  {M}(製番)  {O}(注番)  {T}(型式)  {P}(製品名)  {N}(台数)  {U}(担当者)")]
            public string HeaderTextFormat { get; set; } = string.Empty;

            [Category("\t用紙設定"), DisplayName("ヘッダー フォント"), JsonConverter(typeof(FontJsonConverter))]
            public Font HeaderFont { get; set; } = SystemFonts.DefaultFont;

            [Category("印字設定"), DisplayName("発行枚数")]
            public int CopiesPerLabel { get; set; }
            [Category("印字設定"), DisplayName("印字位置 左 (mm)")]
            public double TextPositionX { get; set; }
            [Category("印字設定"), DisplayName("印字位置 上 (mm)")]
            public double TextPositionY { get; set; }
            [Category("印字設定"), DisplayName("中心に配置 (横)")]
            public bool AlignTextCenterX { get; set; }
            [Category("印字設定"), DisplayName("中心に配置 (縦)")]
            public bool AlignTextCenterY { get; set; }
            [Category("印字設定"), DisplayName("印字フォーマット"), Description("{T}(型式)  {R}(リビジョン)  {S}(シリアル)  {Y}(製造年)  {M}(製造月(1桁))  {MM}(製造月(2桁))")]
            public string LabelTextFormat { get; set; } = string.Empty;

            [Category("印字設定"), DisplayName("印字フォント"), JsonConverter(typeof(FontJsonConverter))]
            public Font TextFont { get; set; } = SystemFonts.DefaultFont;
        }

        public class LabelPrintSettings : PrintSettingsBase {
            [Category("O-Lesラベル設定"), DisplayName("O-Lesラベル発行枚数")]
            public int CopiesPerOLesLabel { get; set; }
            [Category("O-Lesラベル設定"), DisplayName("O-Lesラベル 印字フォーマット"), Description("{OT}(O-Les接頭)  {SA}(O-Les接尾)  {T}(型式)  {R}(リビジョン)  {S}(シリアル)  {Y}(製造年)  {M}(製造月(1桁))  {MM}(製造月(2桁))")]
            public string OLesLabelTextFormat { get; set; } = string.Empty;
        }

        public class BarcodePrintSettings : PrintSettingsBase {
            [Category("印字設定"), DisplayName("バーコード幅 (mm)")]
            public double BarcodeWidth { get; set; }
            [Category("印字設定"), DisplayName("バーコード高さ (mm)")]
            public double BarcodeHeight { get; set; }
            [Category("印字設定"), DisplayName("バーコード位置 左 (mm)")]
            public double BarcodePositionX { get; set; }
            [Category("印字設定"), DisplayName("バーコード位置 上 (mm)")]
            public double BarcodePositionY { get; set; }
            [Category("印字設定"), DisplayName("中心に配置 (横)")]
            public bool AlignBarcodeCenterX { get; set; }
        }

        public class NameplatePrintSettings {
            [Category("印字設定"), DisplayName("発行枚数")]
            public int CopiesPerLabel { get; set; }
            [Category("印字設定"), DisplayName("印字フォーマット"), Description("{T}(型式)  {R}(リビジョン)  {S}(シリアル)  {Y}(製造年)  {M}(製造月(1桁))  {MM}(製造月(2桁))")]
            public string NameplateTextFormat { get; set; } = string.Empty;
            [Category("印字設定"), DisplayName("テンプレートファイルのパス"), Editor(typeof(System.Windows.Forms.Design.FileNameEditor), typeof(System.Drawing.Design.UITypeEditor))]
            public string TemplatePath { get; set; } = string.Empty;
        }

        public class FontJsonConverter : JsonConverter<Font> {
            // JSON文字列をFont型に変換して返す
            public override Font Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) {
                var fontString = reader.GetString()!;
                return (Font)TypeDescriptor.GetConverter(typeof(Font)).ConvertFromString(fontString)!;
            }

            // Font型をJSON文字列に変換して書き込む
            public override void Write(Utf8JsonWriter writer, Font value, JsonSerializerOptions options) {
                var fontString = TypeDescriptor.GetConverter(typeof(Font)).ConvertToString(value)!;
                writer.WriteStringValue(fontString);
            }
        }
    }
}
