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
        public int NumLabelsX { get; set; } = 10;
        public int NumLabelsY { get; set; } = 31;
        public decimal SizeX { get; set; } = 15;
        public decimal SizeY { get; set; } = 5;
        public decimal OffsetX { get; set; } = 12;
        public decimal OffsetY { get; set; } = 11;
        public decimal IntervalX { get; set; } = 4;
        public decimal IntervalY { get; set; } = 4;
        public string HeaderString { get; set; } = "%D 製番 %M 注番 %O %T 台数 %N 担当者 %U";
        public string FooterString { get; set; } = String.Empty;
        public Point HeaderPos { get; set; } = Point.Empty;
        public Point FooterPos { get; set; } = Point.Empty;

        [XmlIgnore]
        public Font HeaderFooterFont { get; set; } = new Font("Arial", 6);

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
        public decimal BarcodeHeight { get; set; } = 1;
        public decimal BarcodePosX { get; set; } = 4;
        public decimal BarcodePosY { get; set; } = 1;
        public decimal BarcodeMagnitude { get; set; } = 1.0M;
        public string Format { get; set; } = "%T%Y%MM%S%R";
        public decimal StringPosX { get; set; } = 4;
        public decimal StringPosY { get; set; } = 1;
        public bool AlignStringCenter { get; set; } = true;
        public bool AlignBarcodeCenter { get; set; } = true;
        public int NumLabels { get; set; } = 1;
        [XmlIgnore]
        public Font Font { get; set; } = new Font("Arial", 6);

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
