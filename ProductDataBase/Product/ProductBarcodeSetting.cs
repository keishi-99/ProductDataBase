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
        public double SizeX { get; set; }
        public double SizeY { get; set; }
        public double OffsetX { get; set; }
        public double OffsetY { get; set; }
        public double IntervalX { get; set; }
        public double IntervalY { get; set; }
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
        public double BarcodeHeight { get; set; }
        public double BarcodePosX { get; set; }
        public double BarcodePosY { get; set; }
        public double BarcodeMagnitude { get; set; }
        public string Format { get; set; } = string.Empty;
        public double StringPosX { get; set; }
        public double StringPosY { get; set; }
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
