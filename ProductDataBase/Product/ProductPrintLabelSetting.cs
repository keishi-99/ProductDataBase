using System.ComponentModel;
using System.Xml.Serialization;

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

    public class BarcodePageSettings {
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
        public Font HeaderFont { get; set; } = new Font("Arial", 5.25F);

        public string HeaderFontAsString {
            get => FontConverter.ConvertToString(HeaderFont);
            set => HeaderFont = FontConverter.ConvertFromString(value);
        }
    }
    public class BarcodeLayoutSettings {
        public string Format { get; set; } = string.Empty;
        public double BarcodeHeight { get; set; }
        public double BarcodePositionX { get; set; }
        public double BarcodePositionY { get; set; }
        public double BarcodeMagnitude { get; set; }
        public double TextPositionX { get; set; }
        public double TextPositionY { get; set; }
        public bool AlignTextXCenter { get; set; }
        public bool AlignBarcodeXCenter { get; set; }
        public int CopiesPerLabel { get; set; }
        [XmlIgnore]
        public Font TextFont { get; set; } = new Font("Arial", 5.25F);

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
