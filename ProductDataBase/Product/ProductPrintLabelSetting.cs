using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ProductDatabase.Product {
    public class ProductPrintSettings {
        public LabelPageSettings LabelPageSettings { get; set; }
        public LabelLayoutSettings LabelLayoutSettings { get; set; }
        public BarcodePageSettings BarcodePageSettings { get; set; }
        public BarcodeLayoutSettings BarcodeLayoutSettings { get; set; }

        public ProductPrintSettings() {
            LabelPageSettings = new LabelPageSettings();
            LabelLayoutSettings = new LabelLayoutSettings();
            BarcodePageSettings = new BarcodePageSettings();
            BarcodeLayoutSettings = new BarcodeLayoutSettings();
        }
    }

    public class LabelPageSettings {
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
        public Font HeaderFont { get; set; } = new Font("ＭＳ Ｐ明朝", 5.25F);
    }
    public class LabelLayoutSettings {
        public string TextFormat { get; set; } = string.Empty;
        public double TextPositionX { get; set; }
        public double TextPositionY { get; set; }
        public bool AlignTextXCenter { get; set; }
        public bool AlignTextYCenter { get; set; }
        public int CopiesPerLabel { get; set; }

        [JsonConverter(typeof(FontJsonConverter))]
        public Font TextFont { get; set; } = new Font("ＭＳ Ｐ明朝", 5.25F);
    }

    public class BarcodePageSettings {
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
        public Font HeaderFont { get; set; } = new Font("ＭＳ Ｐ明朝", 5.25F);
    }
    public class BarcodeLayoutSettings {
        public string TextFormat { get; set; } = string.Empty;
        public double BarcodeHeight { get; set; }
        public double BarcodePositionX { get; set; }
        public double BarcodePositionY { get; set; }
        public double BarcodeMagnitude { get; set; }
        public double TextPositionX { get; set; }
        public double TextPositionY { get; set; }
        public bool AlignTextXCenter { get; set; }
        public bool AlignBarcodeXCenter { get; set; }
        public int CopiesPerLabel { get; set; }

        [JsonConverter(typeof(FontJsonConverter))]
        public Font TextFont { get; set; } = new Font("ＭＳ Ｐ明朝", 5.25F);
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
