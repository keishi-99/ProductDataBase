using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ProductDatabase.Product {
    public class ProductPrintSettings {
        public PrintPageSettings? LabelPageSettings { get; set; }
        public PrintLayoutSettings? LabelLayoutSettings { get; set; }
        public PrintPageSettings? BarcodePageSettings { get; set; }
        public PrintLayoutSettings? BarcodeLayoutSettings { get; set; }

        public ProductPrintSettings() {
            LabelPageSettings = new PrintPageSettings();
            LabelLayoutSettings = new PrintLayoutSettings();
            BarcodePageSettings = new PrintPageSettings();
            BarcodeLayoutSettings = new PrintLayoutSettings();
        }

        public bool ShouldSerializeLabelPageSettings() {
            return LabelPageSettings != null;
        }
        public bool ShouldSerializeLabelLayoutSettings() {
            return LabelLayoutSettings != null;
        }
        public bool ShouldSerializeBarcodePageSettings() {
            return BarcodePageSettings != null;
        }
        public bool ShouldSerializeBarcodeLayoutSettings() {
            return BarcodeLayoutSettings != null;
        }

        public void SetSettingsType(bool isLabelPrint, bool isBarcodePrint) {
            if (!isLabelPrint) {
                LabelPageSettings = null;
                LabelLayoutSettings = null;
            }
            else {
                LabelPageSettings ??= new PrintPageSettings();
                LabelLayoutSettings ??= new PrintLayoutSettings();
            }

            if (!isBarcodePrint) {
                BarcodePageSettings = null;
                BarcodeLayoutSettings = null;
            }
            else {
                BarcodePageSettings ??= new PrintPageSettings();
                BarcodeLayoutSettings ??= new PrintLayoutSettings();
            }
        }
    }

    public class PrintPageSettings {
        public int LabelsPerRow { get; set; }
        public int LabelsPerColumn { get; set; }
        public double LabelWidth { get; set; }
        public double LabelHeight { get; set; }
        public double MarginX { get; set; }
        public double MarginY { get; set; }
        public double IntervalX { get; set; }
        public double IntervalY { get; set; }
        public string HeaderTextFormat { get; set; } = string.Empty;
        public double HeaderPositionX { get; set; }
        public double HeaderPositionY { get; set; }

        [JsonConverter(typeof(FontJsonConverter))]
        public Font HeaderFont { get; set; } = SystemFonts.DefaultFont;
    }
    public class PrintLayoutSettings {
        public string TextFormat { get; set; } = string.Empty;
        public bool AlignTextCenterX { get; set; }
        public bool AlignTextCenterY { get; set; }
        public double TextPositionX { get; set; }
        public double TextPositionY { get; set; }
        public bool AlignBarcodeCenterX { get; set; }
        public double BarcodeHeight { get; set; }
        public double BarcodeWidth { get; set; }
        public double BarcodePositionX { get; set; }
        public double BarcodePositionY { get; set; }
        public int CopiesPerLabel { get; set; }

        [JsonConverter(typeof(FontJsonConverter))]
        public Font TextFont { get; set; } = SystemFonts.DefaultFont;
    }
    public class FontJsonConverter : JsonConverter<Font> {
        public override Font Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) {
            var fontString = reader.GetString()!;
            return (Font)TypeDescriptor.GetConverter(typeof(Font)).ConvertFromString(fontString)!;
        }

        public override void Write(Utf8JsonWriter writer, Font value, JsonSerializerOptions options) {
            var fontString = TypeDescriptor.GetConverter(typeof(Font)).ConvertToString(value)!;
            writer.WriteStringValue(fontString);
        }
    }
}
