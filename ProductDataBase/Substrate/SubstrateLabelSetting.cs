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
        public int NumLabelsX { get; set; } = 10;
        public int NumLabelsY { get; set; } = 31;
        public double SizeX { get; set; } = 15;
        public double SizeY { get; set; } = 5;
        public double OffsetX { get; set; } = 12;
        public double OffsetY { get; set; } = 11;
        public double IntervalX { get; set; } = 4;
        public double IntervalY { get; set; } = 4;
        public string HeaderString { get; set; } = "%D 製番 %M 注番 %O %T 台数 %N 担当者 %U";
        public string FooterString { get; set; } = string.Empty;
        public Point HeaderPos { get; set; } = Point.Empty;
        public Point FooterPos { get; set; } = Point.Empty;

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
        public double BarcodeHeight { get; set; } = 1;
        public double BarcodePosX { get; set; } = 4;
        public double BarcodePosY { get; set; } = 1;
        public double BarcodeMagnitude { get; set; } = 1.0;
        public string Format { get; set; } = "%S";
        public double StringPosX { get; set; } = 4;
        public double StringPosY { get; set; } = 1;
        public bool AlignStringCenter { get; set; } = true;
        public bool AlignBarcodeCenter { get; set; } = true;
        public int NumLabels { get; set; } = 1;
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
