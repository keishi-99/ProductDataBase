using System.ComponentModel;
using System.Xml.Serialization;

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
        public string HeaderText { get; set; } = string.Empty;
        public double HeaderPositionX { get; set; }
        public double HeaderPositionY { get; set; }

        [XmlIgnore]
        public Font HeaderFont { get; set; } = new Font("ＭＳ Ｐ明朝", 5.25F);

        public string HeaderFontAsString {
            get => FontConverter.ConvertToString(HeaderFont);
            set => HeaderFont = FontConverter.ConvertFromString(value);
        }
    }

    public class LabelLayoutSettings {
        public string Format { get; set; } = string.Empty;
        public double TextPositionX { get; set; }
        public double TextPositionY { get; set; }
        public bool AlignTextXCenter { get; set; }
        public bool AlignTextYCenter { get; set; }
        public int CopiesPerLabel { get; set; }

        [XmlIgnore]
        public Font TextFont { get; set; } = new Font("ＭＳ Ｐ明朝", 5.25F);

        public string TextFontAsString {
            get => FontConverter.ConvertToString(TextFont);
            set => TextFont = FontConverter.ConvertFromString(value);
        }
    }
    public static class FontConverter {
        public static string ConvertToString(Font font) {
            return TypeDescriptor.GetConverter(typeof(Font)).ConvertToString(font)!;
        }

        public static Font ConvertFromString(string value) {
            return (Font)TypeDescriptor.GetConverter(typeof(Font)).ConvertFromString(value)!;
        }
    }
}
