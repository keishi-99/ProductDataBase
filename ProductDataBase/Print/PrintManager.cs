using bpac;
using ProductDatabase.Models;
using System.ComponentModel;
using System.Drawing.Printing;
using System.Text.Json;
using System.Text.Json.Serialization;
using ZXing;
using ZXing.Windows.Compatibility;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase.Print {
    // シリアル印刷管理クラス
    public static class PrintManager {

        public static ProductMaster ProductMaster { get; private set; } = default!;
        public static ProductRegisterWork ProductRegisterWork { get; private set; } = default!;

        public static SubstrateMaster SubstrateMaster { get; private set; } = default!;
        public static SubstrateRegisterWork SubstrateRegisterWork { get; private set; } = default!;

        public static DocumentPrintSettings DocumentPrintSettings { get; private set; } = new();

        // 各種設定へのアクセスプロパティ
        public static LabelPrintSettings LabelPrintSettings => DocumentPrintSettings.LabelPrintSettings ?? throw new InvalidOperationException("LabelPrintSettings が null です。");
        public static BarcodePrintSettings BarcodePrintSettings => DocumentPrintSettings.BarcodePrintSettings ?? throw new InvalidOperationException("BarcodePrintSettings が null です。");
        private static DocumentPrintSettings PrintSettings => DocumentPrintSettings ?? throw new InvalidOperationException("DocumentPrintSettings が null です。");

        // 印刷状態を保持するプロパティ
        public static int PrintCount { get; private set; }
        public static int CopiesRemainingPerSerial { get; private set; }
        public static int PageCount { get; private set; }

        public static int SerialPrintType { get; private set; }

        private record SerialEntry(string Serial, int Copies, bool IsOLes);
        private static List<SerialEntry> _serialEntries = [];

        private static bool _isUnderlinePrint;
        private static bool _isLast4Digits;

        // 4桁以上の製品型式の下4桁を取得するプロパティ
        public static string Last4ProductModel =>
            ProductMaster.ProductModel.Length >= 4
                ? ProductMaster.ProductModel[^4..]
                : string.Empty;

        public static SerialType CurrentSerialType { get; set; }
        public enum SerialType { Label, OLesLabel, Barcode, Nameplate, Substrate }

        // 製品印刷に必要なマスター・作業データ・シリアルリストを初期化する
        // olesLabelList を指定した場合は Label・OLes を交互に含むリストを内部で生成する
        public static void ProductInitialize(ProductMaster productMaster, ProductRegisterWork productRegisterWork, DocumentPrintSettings productPrintSettings, List<string> labelList, List<string>? olesLabelList = null) {
            ProductMaster = productMaster;
            ProductRegisterWork = productRegisterWork;
            DocumentPrintSettings = productPrintSettings ?? throw new ArgumentNullException(nameof(productPrintSettings));

            var copiesPerLabel = productPrintSettings.LabelPrintSettings?.CopiesPerLabel ?? 1;
            var copiesPerOles = productPrintSettings.LabelPrintSettings?.CopiesPerOLesLabel ?? 0;

            if (olesLabelList != null && copiesPerOles > 0) {
                if (olesLabelList.Count != labelList.Count) {
                    throw new ArgumentException($"ラベルリストとO-Lesリストの件数が一致しません（Label:{labelList.Count}, OLes:{olesLabelList.Count}）。");
                }
                var entries = new List<SerialEntry>(labelList.Count + olesLabelList.Count);
                for (var i = 0; i < labelList.Count; i++) {
                    entries.Add(new SerialEntry(labelList[i], copiesPerLabel, false));
                    entries.Add(new SerialEntry(olesLabelList[i], copiesPerOles, true));
                }
                _serialEntries = entries;
            }
            else {
                _serialEntries = labelList.Select(s => new SerialEntry(s, copiesPerLabel, false)).ToList();
            }

            PageCount = 1;
            PrintCount = 0;
            SerialPrintType = productMaster.SerialPrintType;

            _isUnderlinePrint = ProductMaster.IsUnderlinePrint;
            _isLast4Digits = ProductMaster.IsLast4Digits;
        }
        // 基板印刷に必要なマスター・作業データ・シリアルリストを初期化する
        public static void SubstrateInitialize(SubstrateMaster substrateMaster, SubstrateRegisterWork substrateRegisterWork, DocumentPrintSettings documentPrintSettings, List<string> serialList) {
            SubstrateMaster = substrateMaster;
            SubstrateRegisterWork = substrateRegisterWork;
            DocumentPrintSettings = documentPrintSettings ?? throw new ArgumentNullException(nameof(documentPrintSettings));

            var copiesPerLabel = documentPrintSettings.LabelPrintSettings?.CopiesPerLabel ?? 1;
            _serialEntries = serialList.Select(s => new SerialEntry(s, copiesPerLabel, false)).ToList();
            PageCount = 1;
            PrintCount = 0;
            SerialPrintType = substrateMaster.SerialPrintType;
        }

        // ミリメートルをピクセルに変換するヘルパーメソッド
        private static float ConvertMmToPixel(double mm, float dpi) {
            const float MmPerInch = 25.4f;
            return (float)(mm / MmPerInch * dpi);
        }

        // 1ページ分のシリアルラベルを印刷（またはプレビュー）し次ページが必要ならtrueを返す
        public static bool PrintSerialCommon(PrintPageEventArgs e, bool isPreview, int startLine, SerialType serialType) {
            try {
                if (e.Graphics is null) { throw new InvalidOperationException("Graphics オブジェクトを取得できません。"); }

                e.Graphics.PageUnit = GraphicsUnit.Pixel;

                var dpiX = e.Graphics.DpiX;
                var dpiY = e.Graphics.DpiY;

                PrintSettingsBase printSettings = serialType switch {
                    SerialType.Label => PrintSettings.LabelPrintSettings!,
                    SerialType.Barcode => PrintSettings.BarcodePrintSettings!,
                    SerialType.Substrate => PrintSettings.LabelPrintSettings!,
                    _ => throw new ArgumentException($"不明なシリアルタイプ: {serialType}")
                };

                var labelCountX = printSettings.LabelsPerColumn;
                var labelCountY = printSettings.LabelsPerRow;
                var labelWidthPx = ConvertMmToPixel(printSettings.LabelWidth, dpiX);
                var labelHeightPx = ConvertMmToPixel(printSettings.LabelHeight, dpiY);
                var marginXPx = ConvertMmToPixel(printSettings.MarginX, dpiX);
                var marginYPx = ConvertMmToPixel(printSettings.MarginY, dpiY);
                var intervalXPx = ConvertMmToPixel(printSettings.IntervalX, dpiX);
                var intervalYPx = ConvertMmToPixel(printSettings.IntervalY, dpiY);
                var headerPositionXPx = ConvertMmToPixel(printSettings.HeaderPositionX, dpiX);
                var headerPositionYPx = ConvertMmToPixel(printSettings.HeaderPositionY, dpiY);
                var headerString = ConvertHeaderString(serialType, printSettings.HeaderTextFormat);
                var headerFont = printSettings.HeaderFont;

                CopiesRemainingPerSerial = _serialEntries[PrintCount].Copies;
                if (labelCountX == 0 || labelCountY == 0 || printSettings.CopiesPerLabel == 0) {
                    throw new Exception("印刷設定が異常です。");
                }

                var (hardMarginX, hardMarginY) = isPreview
                    ? (0, 0)
                    : (e.PageSettings.HardMarginX * dpiX / 100.0f, e.PageSettings.HardMarginY * dpiY / 100.0f);

                //if (marginXPx <= hardMarginX) { throw new Exception($"左余白の値がプリンター左余白より小さい値になっています。\n{marginXPx} <= {hardMarginX}"); }
                //if (marginYPx <= hardMarginY) { throw new Exception($"上余白の値がプリンター上余白より小さい値になっています。\n{marginYPx} <= {hardMarginY}"); }

                if (PageCount >= 2) { startLine = 0; }

                var verticalOffsetPx = PageCount == 1 ? startLine * (intervalYPx + labelHeightPx) : 0;
                e.Graphics.DrawString(headerString, headerFont, Brushes.Gray, headerPositionXPx, verticalOffsetPx + headerPositionYPx - hardMarginY);

                for (var y = startLine; y < labelCountY; y++) {
                    for (var x = 0; x < labelCountX; x++) {
                        var posX = marginXPx - hardMarginX + (x * (intervalXPx + labelWidthPx));
                        var posY = marginYPx - hardMarginY + (y * (intervalYPx + labelHeightPx));

                        var isLastCopy = CopiesRemainingPerSerial == 1;

                        // 残り枚数に基づいてフォントとテキストを決定する
                        var isSecondToLastCopy = CopiesRemainingPerSerial == 2;

                        var entry = _serialEntries[PrintCount];

                        // 両方Trueの場合は末尾2枚目→下線、それ以外は最終コピーのみ下線
                        var fontUnderline = !entry.IsOLes && _isUnderlinePrint && (_isLast4Digits ? isSecondToLastCopy : isLastCopy);

                        var printText = !entry.IsOLes && _isLast4Digits && isLastCopy
                            ? Last4ProductModel
                            : entry.Serial;

                        using var labelImage = MakeLabelImage(printText, serialType, fontUnderline, labelWidthPx, labelHeightPx, dpiX, dpiY, isPreview);

                        e.Graphics.DrawImage(labelImage, posX, posY, labelWidthPx, labelHeightPx);

                        CopiesRemainingPerSerial--;

                        if (CopiesRemainingPerSerial <= 0) {
                            PrintCount++;

                            if (_serialEntries.Count <= PrintCount) {
                                DrawFinalRowMark(e.Graphics, y + 1, 0, posY, 0, labelHeightPx, headerFont);
                                return false;
                            }

                            CopiesRemainingPerSerial = _serialEntries[PrintCount].Copies;
                        }
                    }
                }

                if (_serialEntries.Count > PrintCount) {
                    PageCount++;
                    return true;
                }

                return false;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        // 1枚分のラベルBitmapをシリアルタイプに応じて生成しプレビュー時は枠線を描画する
        private static Bitmap MakeLabelImage(string text, SerialType serialType, bool fontUnderline, float labelWidthPx, float labelHeightPx, float dpiX, float dpiY, bool isPreview) {

            var widthPx = (int)Math.Round(labelWidthPx);
            var heightPx = (int)Math.Round(labelHeightPx);

            var labelImage = new Bitmap(widthPx, heightPx);
            labelImage.SetResolution(dpiX, dpiY);

            using (var g = Graphics.FromImage(labelImage)) {
                g.PageUnit = GraphicsUnit.Pixel;

                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                switch (serialType) {
                    case SerialType.Label:
                        DrawLabel(g, text, fontUnderline, labelWidthPx, labelHeightPx, dpiX, dpiY);
                        break;
                    case SerialType.Barcode:
                        DrawBarcode(g, text, labelWidthPx, labelHeightPx, dpiX, dpiY);
                        break;
                    case SerialType.Substrate:
                        DrawLabel(g, text, fontUnderline, labelWidthPx, labelHeightPx, dpiX, dpiY);
                        break;
                }

                if (isPreview) {
                    using var p = new Pen(Color.Black, ConvertMmToPixel(0.1, dpiX));
                    g.DrawRectangle(p, 0, 0, widthPx - 1, heightPx - 1);
                }
            }

            return labelImage;
        }
        // 設定フォントと位置でテキストラベルを描画する（下線オプションあり）
        private static void DrawLabel(Graphics g, string text, bool fontUnderline, float labelWidthPx, float labelHeightPx, float dpiX, float dpiY) {

            if (PrintSettings.LabelPrintSettings is null) { return; }
            var fontName = PrintSettings.LabelPrintSettings.TextFont.Name;
            var fontSize = PrintSettings.LabelPrintSettings.TextFont.SizeInPoints;
            var style = fontUnderline ? FontStyle.Underline : FontStyle.Regular;

            using var textFont = new Font(fontName, fontSize, style);
            using var sf = new StringFormat {
                Alignment = PrintSettings.LabelPrintSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near,
                LineAlignment = PrintSettings.LabelPrintSettings.AlignTextCenterY ? StringAlignment.Center : StringAlignment.Near
            };
            var textPosX = PrintSettings.LabelPrintSettings.AlignTextCenterX ? 0f : ConvertMmToPixel(PrintSettings.LabelPrintSettings.TextPositionX, dpiX);
            var textPosY = PrintSettings.LabelPrintSettings.AlignTextCenterY ? 0f : ConvertMmToPixel(PrintSettings.LabelPrintSettings.TextPositionY, dpiY);

            var layoutRect = new RectangleF(textPosX, textPosY, labelWidthPx - textPosX, labelHeightPx - textPosY);

            g.DrawString(text, textFont, Brushes.Black, layoutRect, sf);
        }
        // CODE128バーコードと下部テキストをラベルBitmapに描画する
        private static void DrawBarcode(Graphics g, string text, float labelWidthPx, float labelHeightPx, float dpiX, float dpiY) {

            if (PrintSettings.BarcodePrintSettings is null) { return; }
            var fontName = PrintSettings.BarcodePrintSettings.TextFont.Name;
            var fontSize = PrintSettings.BarcodePrintSettings.TextFont.SizeInPoints;

            using (var textFont = new Font(fontName, fontSize, FontStyle.Regular)) {
                using var sf = new StringFormat {
                    Alignment = PrintSettings.BarcodePrintSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near
                };

                var textPosX = PrintSettings.BarcodePrintSettings.AlignTextCenterX ? 0f : ConvertMmToPixel(PrintSettings.BarcodePrintSettings.TextPositionX, dpiX);
                var textPosY = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.TextPositionY, dpiY);

                var layoutRectString = new RectangleF(textPosX, textPosY, labelWidthPx - textPosX, labelHeightPx - textPosY);

                g.DrawString(text, textFont, Brushes.Black, layoutRectString, sf);
            }

            var barcodePixelWidth = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodeWidth, dpiX);
            var barcodePixelHeight = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodeHeight, dpiY);

            var qrWidthPx = (int)Math.Round(barcodePixelWidth);
            var qrHeightPx = (int)Math.Round(barcodePixelHeight);

            var writer = new BarcodeWriter<Bitmap> {
                Format = BarcodeFormat.CODE_128,
                Options = new ZXing.Common.EncodingOptions {
                    Width = qrWidthPx,
                    Height = qrHeightPx,
                    PureBarcode = true
                },
                Renderer = new BitmapRenderer()
            };

            using var barcodeBitmap = writer.Write(text);
            var barcodePosX = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodePositionX, dpiX);
            var barcodePosY = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodePositionY, dpiY);
            var barcodeWidthPx = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodeWidth, dpiX);
            var barcodeHeightPx = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodeHeight, dpiY);

            if (PrintSettings.BarcodePrintSettings.AlignBarcodeCenterX) {
                barcodePosX = (labelWidthPx / 2f) - (barcodeWidthPx / 2f);
            }

            var layoutRectBarcode = new RectangleF(barcodePosX, barcodePosY, barcodeWidthPx, barcodeHeightPx);
            g.DrawImage(barcodeBitmap, layoutRectBarcode);
        }
        // ヘッダーフォーマット文字列内のプレースホルダーをマスターデータの値に置換する
        private static string ConvertHeaderString(SerialType serialType, string s) {

            var map = serialType switch {
                SerialType.Label => CreateProductMap(),
                SerialType.Barcode => CreateProductMap(),
                SerialType.Substrate => CreateSubstrateMap(),
                _ => throw new Exception($"不明な SerialType: {serialType}")
            };

            foreach (var kv in map) {
                s = s.Replace(kv.Key, kv.Value);
            }

            return s;
        }
        // 製品マスター・作業データのプレースホルダーと実値のマッピング辞書を生成する
        private static Dictionary<string, string> CreateProductMap() {

            if (ProductMaster is null) {
                throw new Exception("ProductMaster nullです。");
            }

            if (ProductRegisterWork is null) {
                throw new Exception("ProductRegisterWork nullです。");
            }

            return new Dictionary<string, string> {
                ["{P}"] = ProductMaster.ProductName,
                ["{T}"] = ProductMaster.ProductModel,
                ["{D}"] = DateTime.Today.ToShortDateString(),
                ["{M}"] = ProductRegisterWork.ProductNumber,
                ["{O}"] = ProductRegisterWork.OrderNumber,
                ["{N}"] = ProductRegisterWork.Quantity.ToString(),
                ["{U}"] = ProductRegisterWork.Person,
            };
        }
        // 基板マスター・作業データのプレースホルダーと実値のマッピング辞書を生成する
        private static Dictionary<string, string> CreateSubstrateMap() {

            if (SubstrateMaster is null) {
                throw new Exception("SubstrateMaster nullです。");
            }

            if (SubstrateRegisterWork is null) {
                throw new Exception("SubstrateRegisterWork nullです。");
            }

            return new Dictionary<string, string> {
                ["{P}"] = SubstrateMaster.SubstrateName,
                ["{T}"] = SubstrateMaster.SubstrateModel,
                ["{D}"] = DateTime.Today.ToShortDateString(),
                ["{M}"] = SubstrateRegisterWork.ProductNumber,
                ["{O}"] = SubstrateRegisterWork.OrderNumber,
                ["{N}"] = SubstrateRegisterWork.AddQuantity.ToString(),
                ["{U}"] = SubstrateRegisterWork.Person,
            };
        }


        // 最終シリアル印刷後に次の開始行番号をページ上に描画する
        private static void DrawFinalRowMark(Graphics graphics, int rowNumber, float posX, float posY, float width, float height, Font font) {
            using var sf = new StringFormat {
                Alignment = StringAlignment.Near,
                LineAlignment = StringAlignment.Center
            };
            var layoutRect = new RectangleF(posX, posY, width, height);
            graphics.DrawString(rowNumber.ToString(), font, Brushes.Black, layoutRect, sf);
        }

        // b-pacを使用してNameplateテンプレートにシリアルをセットし指定枚数を連続印刷する
        public static void PrintUsingBPac(NameplatePrintSettings nameplatePrintSettings, List<string> serialList) {

            int copiesPerLabel = nameplatePrintSettings.CopiesPerLabel;
            var templatePath = nameplatePrintSettings.TemplatePath;

            Document? doc = null;
            try {
                doc = new Document();
                if (!doc.Open(templatePath)) {
                    throw new Exception("テンプレートファイルが開けませんでした。");
                }

                doc.StartPrint("", PrintOptionConstants.bpoDefault);

                foreach (var serialNumber in serialList) {
                    doc.GetObject("SerialNo").Text = serialNumber;
                    doc.PrintOut(copiesPerLabel, PrintOptionConstants.bpoDefault);
                }

                doc.EndPrint();
            } finally {
                if (doc != null) {
                    doc.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }
            }
        }
    }
    public class PrintOptions {

        public class DocumentPrintSettings {
            public LabelPrintSettings? LabelPrintSettings { get; set; }
            public BarcodePrintSettings? BarcodePrintSettings { get; set; }
            public NameplatePrintSettings? NameplatePrintSettings { get; set; }

            public DocumentPrintSettings() {
                LabelPrintSettings = new LabelPrintSettings();
                BarcodePrintSettings = new BarcodePrintSettings();
                NameplatePrintSettings = new NameplatePrintSettings();
            }

            // 印刷フラグに応じてラベル・バーコード・銘版の設定インスタンスをnullまたは既存値に設定する
            public void SetSettingsType(bool isLabelPrint, bool isBarcodePrint, bool isNameplatePrint) {
                LabelPrintSettings = isLabelPrint ? (LabelPrintSettings ?? new LabelPrintSettings()) : null;
                BarcodePrintSettings = isBarcodePrint ? (BarcodePrintSettings ?? new BarcodePrintSettings()) : null;
                NameplatePrintSettings = isNameplatePrint ? (NameplatePrintSettings ?? new NameplatePrintSettings()) : null;
            }
        }

        public abstract class PrintSettingsBase {
            [Category("\t用紙設定"), DisplayName("ラベル幅 (mm)")]
            public double LabelWidth { get; set; }
            [Category("\t用紙設定"), DisplayName("ラベル高さ (mm)")]
            public double LabelHeight { get; set; }
            [Category("\t用紙設定"), DisplayName("配置数 (行)")]
            public int LabelsPerRow { get; set; }
            [Category("\t用紙設定"), DisplayName("配置数 (列)")]
            public int LabelsPerColumn { get; set; }
            [Category("\t用紙設定"), DisplayName("余白左 (mm)")]
            public double MarginX { get; set; }
            [Category("\t用紙設定"), DisplayName("余白上 (mm)")]
            public double MarginY { get; set; }
            [Category("\t用紙設定"), DisplayName("ラベル間隔 横 (mm)")]
            public double IntervalX { get; set; }
            [Category("\t用紙設定"), DisplayName("ラベル間隔 縦 (mm)")]
            public double IntervalY { get; set; }
            [Category("\t用紙設定"), DisplayName("ヘッダー開始位置 左 (mm)")]
            public double HeaderPositionX { get; set; }
            [Category("\t用紙設定"), DisplayName("ヘッダー開始位置 上 (mm)")]
            public double HeaderPositionY { get; set; }
            [Category("\t用紙設定"), DisplayName("ヘッダー フォーマット"), Description("{D}(印刷日)  {M}(製番)  {O}(注番)  {T}(型式)  {P}(製品名)  {N}(台数)  {U}(担当者)")]
            public string HeaderTextFormat { get; set; } = string.Empty;

            [Category("\t用紙設定"), DisplayName("ヘッダー フォント"), JsonConverter(typeof(FontJsonConverter))]
            public Font HeaderFont { get; set; } = SystemFonts.DefaultFont;

            [Category("印字設定"), DisplayName("発行枚数")]
            public int CopiesPerLabel { get; set; }
            [Category("印字設定"), DisplayName("印字位置 左 (mm)")]
            public double TextPositionX { get; set; }
            [Category("印字設定"), DisplayName("印字位置 上 (mm)")]
            public double TextPositionY { get; set; }
            [Category("印字設定"), DisplayName("中心に配置 (横)")]
            public bool AlignTextCenterX { get; set; }
            [Category("印字設定"), DisplayName("中心に配置 (縦)")]
            public bool AlignTextCenterY { get; set; }
            [Category("印字設定"), DisplayName("印字フォーマット"), Description("{T}(型式)  {R}(リビジョン)  {S}(シリアル)  {Y}(製造年)  {M}(製造月(1桁))  {MM}(製造月(2桁))")]
            public string LabelTextFormat { get; set; } = string.Empty;

            [Category("印字設定"), DisplayName("印字フォント"), JsonConverter(typeof(FontJsonConverter))]
            public Font TextFont { get; set; } = SystemFonts.DefaultFont;

            [Category("O-Lesラベル設定"), DisplayName("O-Lesラベル発行枚数")]
            public int CopiesPerOLesLabel { get; set; }
            [Category("O-Lesラベル設定"), DisplayName("O-Lesラベル 印字フォーマット"), Description("{OT}(O-Les型式)  {T}(型式)  {R}(リビジョン)  {S}(シリアル)  {Y}(製造年)  {M}(製造月(1桁))  {MM}(製造月(2桁))")]
            public string OLesLabelTextFormat { get; set; } = string.Empty;

        }

        public class LabelPrintSettings : PrintSettingsBase {
        }

        public class BarcodePrintSettings : PrintSettingsBase {
            [Category("印字設定"), DisplayName("バーコード幅 (mm)")]
            public double BarcodeWidth { get; set; }
            [Category("印字設定"), DisplayName("バーコード高さ (mm)")]
            public double BarcodeHeight { get; set; }
            [Category("印字設定"), DisplayName("バーコード位置 左 (mm)")]
            public double BarcodePositionX { get; set; }
            [Category("印字設定"), DisplayName("バーコード位置 上 (mm)")]
            public double BarcodePositionY { get; set; }
            [Category("印字設定"), DisplayName("中心に配置 (横)")]
            public bool AlignBarcodeCenterX { get; set; }
        }

        public class NameplatePrintSettings {
            [Category("印字設定"), DisplayName("発行枚数")]
            public int CopiesPerLabel { get; set; }
            [Category("印字設定"), DisplayName("印字フォーマット"), Description("{T}(型式)  {R}(リビジョン)  {S}(シリアル)  {Y}(製造年)  {M}(製造月(1桁))  {MM}(製造月(2桁))")]
            public string TextFormat { get; set; } = string.Empty;
            [Category("印字設定"), DisplayName("テンプレートファイルのパス"), Editor(typeof(System.Windows.Forms.Design.FileNameEditor), typeof(System.Drawing.Design.UITypeEditor))]
            public string TemplatePath { get; set; } = string.Empty;
        }

        public class FontJsonConverter : JsonConverter<Font> {
            // JSON文字列をFont型に変換して返す
            public override Font Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) {
                var fontString = reader.GetString()!;
                return (Font)TypeDescriptor.GetConverter(typeof(Font)).ConvertFromString(fontString)!;
            }

            // Font型をJSON文字列に変換して書き込む
            public override void Write(Utf8JsonWriter writer, Font value, JsonSerializerOptions options) {
                var fontString = TypeDescriptor.GetConverter(typeof(Font)).ConvertToString(value)!;
                writer.WriteStringValue(fontString);
            }
        }
    }
}
