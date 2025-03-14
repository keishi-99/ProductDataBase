using System.ComponentModel;
using System.Xml.Serialization;

namespace ProductDatabase.Product {
    public class CSettingsLabelPro {
        public CLabelProPageSettings LabelProPageSettings { get; set; }
        public CLabelProLabelSettings LabelProLabelSettings { get; set; }

        public CSettingsLabelPro() {
            LabelProPageSettings = new CLabelProPageSettings();
            LabelProLabelSettings = new CLabelProLabelSettings();
        }
    }

    public class CLabelProPageSettings {
        public int LabelCountX { get; set; }
        public int LabelCountY { get; set; }
        public double LabelWidth { get; set; }
        public double LabelHeight { get; set; }
        public double MarginX { get; set; }
        public double MarginY { get; set; }
        public double IntervalX { get; set; }
        public double IntervalY { get; set; }
        public string HeaderString { get; set; } = string.Empty;
        public Point HeaderPos { get; set; }

        [XmlIgnore]
        public Font HeaderFooterFont { get; set; } = new Font("ＭＳ Ｐ明朝", 5.25F);

        public string FontAsString {
            get => ConvertToString(HeaderFooterFont);
            set => HeaderFooterFont = ConvertFromString<Font>(value);
        }

        public static string ConvertToString<tValue>(tValue value) {
            return TypeDescriptor.GetConverter(typeof(tValue)).ConvertToString(value)!;
        }

        public static tValue ConvertFromString<tValue>(string value) {
            return (tValue)TypeDescriptor.GetConverter(typeof(tValue)).ConvertFromString(value)!;
        }
    }

    public class CLabelProLabelSettings {
        public string Format { get; set; } = string.Empty;
        public double StringPosX { get; set; }
        public double StringPosY { get; set; }
        public bool AlignStringXCenter { get; set; }
        public bool AlignStringYCenter { get; set; }
        public int NumLabels { get; set; }
        [XmlIgnore]
        public Font Font { get; set; } = new Font("ＭＳ Ｐ明朝", 5.25F);

        [Browsable(false)]
        public string FontAsString {
            get => ConvertToString(Font);
            set => Font = ConvertFromString<Font>(value);
        }

        public static string ConvertToString<tValue>(tValue value) {
            return TypeDescriptor.GetConverter(typeof(tValue)).ConvertToString(value)!;
        }

        public static tValue ConvertFromString<tValue>(string value) {
            return (tValue)TypeDescriptor.GetConverter(typeof(tValue)).ConvertFromString(value)!;
        }
    }
}
