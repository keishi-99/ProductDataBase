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
        public int NumLabelsX { get; set; }
        public int NumLabelsY { get; set; }
        public decimal SizeX { get; set; }
        public decimal SizeY { get; set; }
        public decimal OffsetX { get; set; }
        public decimal OffsetY { get; set; }
        public decimal IntervalX { get; set; }
        public decimal IntervalY { get; set; }
        public string HeaderString { get; set; } = string.Empty;
        public string FooterString { get; set; } = string.Empty;
        public Point HeaderPos { get; set; }
        public Point FooterPos { get; set; }

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
        public decimal StringHeight { get; set; }
        public decimal StringMagnitude { get; set; }
        public decimal StringPosX { get; set; }
        public decimal StringPosY { get; set; }
        public bool AlignStringCenter { get; set; }
        public bool AlignBarcodeCenter { get; set; }
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
