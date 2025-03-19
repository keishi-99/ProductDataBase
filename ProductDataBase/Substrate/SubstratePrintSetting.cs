using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ProductDatabase.Substrate {
    public class SubstratePrintSettings {
        public LabelPageSettings LabelPageSettings { get; set; }
        public LabelLayoutSettings LabelLayoutSettings { get; set; }

        public SubstratePrintSettings() {
            LabelPageSettings = new LabelPageSettings();
            LabelLayoutSettings = new LabelLayoutSettings();
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
        public Font HeaderFont { get; set; } = SystemFonts.DefaultFont;
    }

    public class LabelLayoutSettings {
        public string TextFormat { get; set; } = string.Empty;
        public double TextPositionX { get; set; }
        public double TextPositionY { get; set; }
        public bool AlignTextXCenter { get; set; }
        public bool AlignTextYCenter { get; set; }
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
