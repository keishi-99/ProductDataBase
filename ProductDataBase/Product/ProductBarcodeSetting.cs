using System.ComponentModel;
using System.Xml.Serialization;

namespace ProductDatabase.Product {
    public class CSettingsBarcodePro {
        public CBarcodeProPageSettings BarcodeProPageSettings { get; set; }
        public CBarcodeProLabelSettings BarcodeProLabelSettings { get; set; }

        public CSettingsBarcodePro() {
            BarcodeProPageSettings = new CBarcodeProPageSettings();
            BarcodeProLabelSettings = new CBarcodeProLabelSettings();
        }
    }

    public class CBarcodeProPageSettings {
        public int NumLabelsX { get; set; }
        public int NumLabelsY { get; set; }
        public decimal SizeX { get; set; }
        public decimal SizeY { get; set; }
        public decimal OffsetX { get; set; }
        public decimal OffsetY { get; set; }
        public decimal IntervalX { get; set; }
        public decimal IntervalY { get; set; }
        public string HeaderString { get; set; } = string.Empty;
        public Point HeaderPos { get; set; }

        [XmlIgnore]
        public Font HeaderFooterFont { get; set; } = new Font("Arial", 5.25F);

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

    public class CBarcodeProLabelSettings {
        public decimal BarcodeHeight { get; set; }
        public decimal BarcodePosX { get; set; }
        public decimal BarcodePosY { get; set; }
        public decimal BarcodeMagnitude { get; set; }
        public string Format { get; set; } = string.Empty;
        public decimal StringPosX { get; set; }
        public decimal StringPosY { get; set; }
        public bool AlignStringCenter { get; set; }
        public bool AlignBarcodeCenter { get; set; }
        public int NumLabels { get; set; }
        [XmlIgnore]
        public Font Font { get; set; } = new Font("Arial", 5.25F);

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
