using System.ComponentModel;
using System.Xml.Serialization;

namespace ProductDatabase.Substrate {
    public class CSettingsLabelSub {
        public CLabelSubPageSettings LabelSubPageSettings { get; set; }
        public CLabelSubLabelSettings LabelSubLabelSettings { get; set; }

        public CSettingsLabelSub() {
            LabelSubPageSettings = new CLabelSubPageSettings();
            LabelSubLabelSettings = new CLabelSubLabelSettings();
        }
    }

    public class CLabelSubPageSettings {
        public int NumLabelsX { get; set; }
        public int NumLabelsY { get; set; }
        public double SizeX { get; set; }
        public double SizeY { get; set; }
        public double OffsetX { get; set; }
        public double OffsetY { get; set; }
        public double IntervalX { get; set; }
        public double IntervalY { get; set; }
        public string HeaderString { get; set; } = string.Empty;
        public Point HeaderPos { get; set; } = Point.Empty;

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

    public class CLabelSubLabelSettings {
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
